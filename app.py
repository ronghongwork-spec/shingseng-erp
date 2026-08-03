import base64
import concurrent.futures
import json
import os
import sys
from datetime import datetime

from nicegui import ui
import pandas as pd
import requests


def safe_text(value, default=""):
  """把值轉成安全的顯示字串，同時避免 pandas 的 NaN 陷阱。

  pandas 若某欄位全部是 None，會把該欄自動轉成 float64 的 NaN；但
  Python 的 bool(nan) 是 True，用 `value or default` 這種寫法會誤判
  NaN 為「有值」，後續 f-string 或字串操作就會出錯。這裡統一用
  pd.isna() 明確判斷。
  """
  if value is None:
    return default
  try:
    if pd.isna(value):
      return default
  except (TypeError, ValueError):
    pass
  return str(value)

# -------------------------------------------------------------------------
# 1. 鼎新 A1 API 串接與全量資料自動抓取
#    依「鼎新 A1 商務應用雲 POS API 串接手冊」版本 1.0.35
# -------------------------------------------------------------------------
# 本機開發時，若有安裝 python-dotenv 且專案根目錄有 .env 檔，會自動載入；
# 部署到 Render 時則直接讀取 Render 後台設定的環境變數，不需要 .env 檔。
try:
  from dotenv import load_dotenv
  load_dotenv()
except ImportError:
  pass

A1_BASE_URL = "http://a1external.digiwin.com"  # 正式區 URL（1.0.35 已取消試用區，範例改正式區）
API_KEY = os.environ.get("A1_API_KEY", "")            # 鼎新 A1 APIKey（海濤客食品廠2 專屬）
API_PASSWORD = os.environ.get("A1_API_PASSWORD", "")  # 鼎新 A1 Password
COMPANY_NAME = os.environ.get("A1_COMPANY_NAME", "海濤客食品廠2")  # 分公司/廠別標示，用於頁面標題

if not API_KEY or not API_PASSWORD:
  print(
      "警告：尚未設定環境變數 A1_API_KEY / A1_API_PASSWORD，"
      "將無法登入鼎新 A1，頁面會改用測試防呆資料。",
      file=sys.stderr,
  )

REQUEST_TIMEOUT = 15   # 秒，避免 A1 主機無回應時整頁卡死
STOCK_PAGE_SIZE = 100  # 手冊：StockBatch 每頁固定 100 筆（依 品號+倉庫 計算）
MAX_STOCK_PAGES = 1000  # 分頁安全上限，避免 More 一直為 true 造成無窮迴圈
ITEM_DETAIL_WORKERS = 8  # 平行抓取商品明細的執行緒數

# 鼎新 A1 目前的 POS API（手冊 1.0.35）沒有提供「組合品-組成明細」的查詢
# 端點，只能改用人工維護的 Excel 來補齊「主件/子件品號＋用量」關係，
# 再由本程式讀取、合併進「商品組合資訊」頁籤顯示。
#
# 部署到 Render 時請注意：Render 的一般檔案系統在「每次重新部署／服務
# 重啟」時都會被清空還原成 git 版本，並不是永久保存的硬碟。若要讓網頁
# 上傳的 Excel 檔案能長期保留，需要在 Render 後台為此服務加裝一個
# Persistent Disk（付費功能），掛載到一個固定路徑（例如 /var/data），
# 再把環境變數 A1_BOM_EXCEL_PATH 設成該路徑底下的檔案位置，例如：
#   A1_BOM_EXCEL_PATH=/var/data/商品組合明細.xlsx
# 若沒有加裝 Persistent Disk，網頁上傳的檔案在下次部署/重啟後就會消失，
# 需要重新上傳一次（但服務持續運作期間內都可以正常使用、下載、查詢）。
BOM_EXCEL_PATH = os.environ.get(
    "A1_BOM_EXCEL_PATH",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "商品組合明細.xlsx"),
)
BOM_EXCEL_SHEET_NAME = "組合明細"

# ---- Google Sheets 版 BOM（正式串接方式，逐步取代上面的 Excel 上傳） ----
# 需要在 Google Cloud 建立一個服務帳號(Service Account)，下載其 JSON 金鑰，
# 並把該服務帳號的 email 加入 Google Sheet 的「共用」名單（唯讀權限即可）。
# 三個環境變數都要設定，系統才會改用 Google Sheets；任一沒設定，就自動
# 退回上面的本機 Excel 上傳機制（過渡期相容，不會讓頁面壞掉）。
GOOGLE_SHEETS_CREDENTIALS_JSON = os.environ.get(
    "GOOGLE_SHEETS_CREDENTIALS_JSON", ""
)  # 服務帳號 JSON 金鑰的「內容」（整包貼進環境變數），不是檔案路徑
BOM_GOOGLE_SHEET_ID = os.environ.get("BOM_GOOGLE_SHEET_ID", "")
BOM_GOOGLE_SHEET_TAB = os.environ.get("BOM_GOOGLE_SHEET_TAB", "組合明細")


def get_a1_token():
  """透過 APIKey + Password 呼叫 Login，取得完整可用的 Authorization 值。

  手冊 Login[Post]：UserName 填 APIKey，Password 填 Password，兩者皆為必填，
  缺一則回 401001(帳號密碼空白) 或 401002(帳號密碼錯誤)。
  登入有效期限為 12 小時。
  """
  url = f"{A1_BASE_URL}/Login"
  headers = {"Content-Type": "application/json"}
  body = {"UserName": API_KEY, "Password": API_PASSWORD}

  try:
    response = requests.post(
        url, json=body, headers=headers, timeout=REQUEST_TIMEOUT
    )
    if response.status_code == 200:
      data = response.json()
      access_token = data.get("access_token")
      if not access_token:
        print(f"A1 登入回應缺少 access_token: {data}")
        return None
      # 手冊寫的是「Authorization Header 填入使用者金鑰」，不是標準 OAuth
      # 的 "Bearer <token>" 格式，這裡直接回傳原始 access_token 給後續呼叫使用。
      print(f"A1 登入成功，token 前 8 碼: {access_token[:8]}...")
      return access_token
    else:
      print(f"A1 登入失敗 [{response.status_code}]: {response.text}")
  except requests.exceptions.RequestException as e:
    print(f"A1 登入連線異常: {e}")
  return None


def fetch_warehouses(token):
  """動態取得所有未停用的倉庫列表（Warehouses[Get]）"""
  url = f"{A1_BASE_URL}/Warehouses"
  headers = {"Authorization": token}
  try:
    response = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
    if response.status_code == 200:
      return [w["Name"] for w in response.json()]
    print(f"取得倉庫列表失敗 [{response.status_code}]: {response.text}")
  except requests.exceptions.RequestException as e:
    print(f"取得倉庫列表失敗: {e}")
  return []


def fetch_categories(token):
  """動態取得所有商品分類列表（Categorys[Get]）

  回傳 {分類代號字串: 分類名稱}，代號統一轉字串，避免與商品明細的
  CategoryID 型別不一致時比對不到而全部落在「未分類」。
  """
  url = f"{A1_BASE_URL}/Categorys"
  headers = {"Authorization": token}
  try:
    response = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
    if response.status_code == 200:
      return {str(c["ID"]): c["Name"] for c in response.json()}
    print(f"取得商品分類失敗 [{response.status_code}]: {response.text}")
  except requests.exceptions.RequestException as e:
    print(f"取得商品分類失敗: {e}")
  return {}


def fetch_item_image_data_uri(item_id, headers, sequence=1):
  """依手冊 ItemImage[Get]（/ItemImage/{ItemID}/{Sequence}）取得單一商品圖片，
  將回傳的二進位圖檔轉成 base64 data URI，方便直接放進 <img> 顯示。

  手冊備註：若該品號沒有上傳過圖片，會回傳 400 400028（商品圖檔不存在），
  這是正常情況（大部分商品可能都還沒有圖），此處直接回傳 None，不視為錯誤。
  """
  url = f"{A1_BASE_URL}/ItemImage/{item_id}/{sequence}"
  try:
    response = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
    if response.status_code == 200:
      content_type = response.headers.get("Content-Type", "image/png")
      encoded = base64.b64encode(response.content).decode("utf-8")
      return f"data:{content_type};base64,{encoded}"
    if response.status_code != 400:
      print(f"取得商品圖片失敗 [{item_id}] [{response.status_code}]: {response.text}")
  except requests.exceptions.RequestException as e:
    print(f"取得商品圖片連線異常 [{item_id}]: {e}")
  return None


def fetch_items_map(token):
  """取得商品詳細資料（對應品名、分類、單位、平均成本、商品圖片等）

  手冊 Items[Get] 無傳入商品代號時，只回傳 ID/Name，要拿到 CategoryID、
  UnitName、StdPurPrice 等完整欄位，必須逐筆呼叫 Items/{ItemID}。
  商品數量多時逐一序列呼叫會很慢，這裡改用多執行緒平行抓取明細，
  並針對單筆失敗加入重試，避免暫時性網路錯誤讓某些商品被靜默漏掉。
  同時依手冊 ItemImage[Get] 一併嘗試抓取每個品號的圖片（第 1 張），
  沒有圖片的商品會是 None，前端會改用預留圖示顯示。
  """
  url = f"{A1_BASE_URL}/Items"
  headers = {"Authorization": token}
  items_dict = {}

  try:
    response = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
    if response.status_code != 200:
      print(f"取得商品列表失敗 [{response.status_code}]: {response.text}")
      return items_dict
    item_ids = [item.get("ID") for item in response.json() if item.get("ID")]
  except requests.exceptions.RequestException as e:
    print(f"取得商品列表失敗: {e}")
    return items_dict

  ITEM_DETAIL_RETRIES = 2

  def fetch_one(item_id):
    detail_url = f"{A1_BASE_URL}/Items/{item_id}"
    last_error = None
    for attempt in range(1, ITEM_DETAIL_RETRIES + 1):
      try:
        detail_res = requests.get(
            detail_url, headers=headers, timeout=REQUEST_TIMEOUT
        )
        if detail_res.status_code == 200:
          detail = detail_res.json()
          detail["ImageDataURI"] = fetch_item_image_data_uri(item_id, headers)
          return item_id, detail
        last_error = f"[{detail_res.status_code}]: {detail_res.text}"
      except requests.exceptions.RequestException as e:
        last_error = str(e)
    print(f"取得商品明細失敗（已重試{ITEM_DETAIL_RETRIES}次）[{item_id}] {last_error}")
    return item_id, None

  with concurrent.futures.ThreadPoolExecutor(
      max_workers=ITEM_DETAIL_WORKERS
  ) as executor:
    for item_id, detail in executor.map(fetch_one, item_ids):
      if detail:
        items_dict[item_id] = detail

  failed_count = len(item_ids) - len(items_dict)
  print(
      f"商品明細抓取完成：品號清單共 {len(item_ids)} 筆，"
      f"成功取得明細 {len(items_dict)} 筆，失敗 {failed_count} 筆"
  )

  return items_dict


def fetch_all_a1_inventory():
  """透過 StockBatch API，以「品號分批（ItemIDs）」的方式完整抓取庫存資料

  手冊 StockBatch[Post]：每頁固定 100 筆（依 品號+倉庫 組合計算），
  且支援 ItemIDs 參數一次帶入多個品號查詢（最多 100 筆）。

  實測發現：不指定品號、查「全部品號」時（不論有無搭配 WarehouseName），
  分頁的 More 旗標會提前變成 false，導致漏抓資料——即使該品號用
  Stock[Get]（單品查詢）直查是查得到、有庫存的。

  因此改為：先用 Items[Get] 取得完整品號清單，切成每批 100 個
  （符合手冊 ItemIDs 上限），逐批帶入 ItemIDs 查詢 StockBatch，
  每批各自分頁到底再合併。這樣每一批查詢的品號範圍都是我們自己
  明確指定的，不會受「全部品號」模式下 More 旗標異常的影響。
  """
  token = get_a1_token()

  if not token:
    print("無法取得 A1 Token，啟用測試防呆數據...")
    return get_mock_data(), [], [], {}

  headers = {
      "Content-Type": "application/json",
      "Authorization": token,
  }

  warehouses = fetch_warehouses(token)
  categories_map = fetch_categories(token)
  items_map = fetch_items_map(token)

  def fetch_stock_rows(item_ids_batch=None, warehouse_name=None):
    """依 ItemIDs（品號批次）、可選 WarehouseName，完整分頁抓取 StockBatch 原始資料列"""
    url = f"{A1_BASE_URL}/Stock/Batch"
    rows_collected = []
    pagination = 1
    more_data = True

    while more_data and pagination <= MAX_STOCK_PAGES:
      payload = {"Pagination": pagination}
      if item_ids_batch:
        payload["ItemIDs"] = item_ids_batch
      if warehouse_name:
        payload["WarehouseName"] = warehouse_name

      try:
        response = requests.post(
            url, json=payload, headers=headers, timeout=REQUEST_TIMEOUT
        )
        if response.status_code == 200:
          res_json = response.json()
          rows = (
              res_json if isinstance(res_json, list) else res_json.get("Data", [])
          )
          rows_collected.extend(rows)

          more_data = (
              res_json.get("More", False) if isinstance(res_json, dict) else False
          )
          pagination += 1
        else:
          print(
              f"StockBatch 抓取失敗 [{response.status_code}]: {response.text}"
          )
          break
      except requests.exceptions.RequestException as e:
        print(f"StockBatch 請求異常: {e}")
        break

    if pagination > MAX_STOCK_PAGES:
      print(f"警告：StockBatch 已達分頁安全上限 {MAX_STOCK_PAGES} 頁，資料可能未抓取完整")

    return rows_collected

  ITEM_BATCH_SIZE = 100  # 手冊：ItemIDs 一次最多可傳 100 筆
  all_item_ids = list(items_map.keys())
  raw_rows = []

  if all_item_ids:
    batches = [
        all_item_ids[i:i + ITEM_BATCH_SIZE]
        for i in range(0, len(all_item_ids), ITEM_BATCH_SIZE)
    ]
    for batch_index, batch in enumerate(batches, start=1):
      batch_rows = fetch_stock_rows(item_ids_batch=batch)
      print(
          f"StockBatch 品號批次 {batch_index}/{len(batches)}"
          f"（{len(batch)} 個品號）抓取完成，共 {len(batch_rows)} 筆"
      )
      raw_rows.extend(batch_rows)
  else:
    # 沒取到商品清單時，退回原本「不分品號」的查詢方式
    print("警告：未取得商品清單，改用不分品號的方式查詢 StockBatch")
    raw_rows = fetch_stock_rows()

  all_stock_data = []
  for row in raw_rows:
    item_id = row.get("ItemID")
    item_info = items_map.get(item_id, {})

    cat_id = item_info.get("CategoryID")
    cat_name = categories_map.get(str(cat_id), "未分類")

    all_stock_data.append({
        "倉庫名稱": row.get("WarehouseName"),
        "商品分類": cat_name,
        "品號": item_id,
        "品名": row.get("ItemName") or item_info.get("Name"),
        "單位": item_info.get("UnitName", "個"),
        "庫存數量": row.get("Qty", 0.0),
        "平均成本": item_info.get("StdPurPrice", 0.0),
        "圖片": item_info.get("ImageDataURI"),
        "商品型態": item_info.get("Type"),
    })

  # 手冊備註：StockBatch「若商品為新建，未在任何倉庫中有異動，則不會回傳」，
  # 這代表組合品（尤其是[先銷售自動組合]這種不佔實體庫存的類型）很可能完全不會
  # 出現在 StockBatch 的結果裡。為了讓「各商品分類」的統計不漏掉這些商品，
  # 這裡以商品主檔（items_map，來自 Items[Get]）為準，把 StockBatch 完全沒提到
  # 的品號也補一筆進表格，庫存數量顯示 0、倉庫顯示「(無庫存異動)」。
  covered_item_ids = {row["品號"] for row in all_stock_data}
  missing_items = [
      (item_id, info)
      for item_id, info in items_map.items()
      if item_id not in covered_item_ids
  ]
  for item_id, item_info in missing_items:
    cat_id = item_info.get("CategoryID")
    cat_name = categories_map.get(str(cat_id), "未分類")
    all_stock_data.append({
        "倉庫名稱": "(無庫存異動)",
        "商品分類": cat_name,
        "品號": item_id,
        "品名": item_info.get("Name"),
        "單位": item_info.get("UnitName", "個"),
        "庫存數量": 0.0,
        "平均成本": item_info.get("StdPurPrice", 0.0),
        "圖片": item_info.get("ImageDataURI"),
        "商品型態": item_info.get("Type"),
    })

  print(
      f"庫存資料彙整完成：StockBatch 回傳 {len(covered_item_ids)} 個不同品號，"
      f"另補上 {len(missing_items)} 個從未有庫存異動的品號，"
      f"總計 {len(all_stock_data)} 列（品號 x 倉庫）"
  )

  # 防呆機制：若 API 無資料或連線失敗，回傳範例資料
  if not all_stock_data:
    return get_mock_data(), warehouses, list(categories_map.values()), items_map

  return (
      pd.DataFrame(all_stock_data),
      warehouses,
      list(categories_map.values()),
      items_map,
  )


def fetch_stock_single_item(token, item_id):
  """Stock[Get]：查詢單一品號在所有倉庫的庫存量（不分頁，用於交叉驗證 StockBatch）

  回傳 (成功與否, 訊息, [{WarehouseName, Qty}, ...])
  """
  url = f"{A1_BASE_URL}/Stock/{item_id}"
  headers = {"Authorization": token}
  try:
    response = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
    if response.status_code == 200:
      data = response.json()
      rows = data if isinstance(data, list) else [data] if data else []
      return True, "查詢成功", rows
    return False, f"[{response.status_code}] {response.text}", []
  except requests.exceptions.RequestException as e:
    return False, str(e), []


def fetch_all_lot_nos(token):
  """ItemLotNos[Get]（不傳品號）：取得所有商品的批號資料，含有效日期。

  食品業批號管理很依賴這支——A1 商品主檔本身沒有「保存效期」欄位，
  但有租用批號模組、且商品有啟用批號管理的話，每一批號會各自記錄
  ExpiryDate（有效日期），可以用這個做效期預警。
  沒有租用批號模組，或商品都沒啟用批號管理時，這裡會回傳空列表，
  頁面會顯示「目前沒有批號資料」，不會出錯。
  """
  url = f"{A1_BASE_URL}/ItemLotNos"
  headers = {"Authorization": token}
  try:
    response = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
    if response.status_code == 200:
      return True, response.json()
    return False, f"[{response.status_code}] {response.text}"
  except requests.exceptions.RequestException as e:
    return False, str(e)


def get_mock_data():
  """提供本地測試用的防呆 DataFrame（A1 連線失敗時的備援顯示資料）"""
  return pd.DataFrame([
      {
          "倉庫名稱": "食品廠鳳仁倉",
          "商品分類": "(海濤客)_成品11",
          "品號": "011101180001",
          "品名": "醬料_烏金干貝醬",
          "單位": "罐",
          "庫存數量": 262.00,
          "平均成本": 214.41,
          "圖片": None,
          "商品型態": "1",
      },
      {
          "倉庫名稱": "食品廠鳳仁倉",
          "商品分類": "(海濤客)_成品11",
          "品號": "011101180002",
          "品名": "醬料_飛魚卵XO醬",
          "單位": "罐",
          "庫存數量": 258.00,
          "平均成本": 90.50,
          "圖片": None,
          "商品型態": "1",
      },
      {
          "倉庫名稱": "永福倉",
          "商品分類": "(海濤客)_原料21",
          "品號": "011102000001",
          "品名": "原料_干貝散裝",
          "單位": "公斤",
          "庫存數量": 500.00,
          "平均成本": 150.00,
          "圖片": None,
          "商品型態": "1",
      },
      {
          "倉庫名稱": "小琉球現場",
          "商品分類": "(海濤客)_限定組合99",
          "品號": "011109900001",
          "品名": "現場限定澎湃禮盒",
          "單位": "組",
          "庫存數量": 45.00,
          "平均成本": 680.00,
          "圖片": None,
          "商品型態": "2",
      },
  ])


ITEM_TYPE_LABELS = {
    "1": "一般商品",
    "2": "組合品-先組合再銷售",
    "3": "組合品-先銷售自動組合",
}

# 對應「商品組合明細」Excel／Google Sheet 的欄位名稱
# （Excel 版沿用之前範本的「（選填，供參考）」字樣；Google Sheets 版本
# 建議直接用同一組欄位名稱建表頭，程式才讀得到）
BOM_COL_PARENT_ID = "主件品號"
BOM_COL_PARENT_NAME = "主件品名（選填，供參考）"
BOM_COL_CHILD_ID = "子件品號"
BOM_COL_CHILD_NAME = "子件品名（選填，供參考）"
BOM_COL_QTY = "用量"
BOM_COL_UNIT = "單位（選填）"
BOM_COL_LOSS_RATE = "損耗率(%)"          # 新增：子件用量的耗損比例
BOM_COL_LEAD_TIME = "子件採購前置天數"     # 新增：下單後供應商需要幾天到貨
BOM_COL_WORK_DAYS = "主件生產/組裝工時(天)"  # 新增：備齊原料後還要幾天才能組裝完成
BOM_COL_SUPPLIER = "主要供應商（選填）"     # 新增
BOM_COL_MEMO = "備註"


def _parse_bom_records(records):
  """把「一列＝一筆主件+子件關係」的原始資料（不管是從 Excel 或 Google
  Sheet 讀來的）統一轉成 {主件品號: [子件明細...]} 的結構。

  records: list[dict]，每個 dict 的 key 是欄位名稱（BOM_COL_*）
  """
  bom_map = {}
  for row in records:
    parent_id = str(row.get(BOM_COL_PARENT_ID, "") or "").strip()
    child_id = str(row.get(BOM_COL_CHILD_ID, "") or "").strip()
    if not parent_id or not child_id:
      continue  # 略過空白列或只填了一半的列

    def _to_float(raw, default=0.0):
      raw = str(raw).strip() if raw not in (None, "") else ""
      if not raw:
        return default
      try:
        return float(raw)
      except ValueError:
        return raw  # 萬一填了非數字，原樣顯示，不擋住整批匯入

    bom_map.setdefault(parent_id, []).append({
        "子件品號": child_id,
        "子件品名": str(row.get(BOM_COL_CHILD_NAME, "") or "").strip(),
        "用量": _to_float(row.get(BOM_COL_QTY)),
        "單位": str(row.get(BOM_COL_UNIT, "") or "").strip(),
        "損耗率": _to_float(row.get(BOM_COL_LOSS_RATE), default=0.0),
        "採購前置天數": _to_float(row.get(BOM_COL_LEAD_TIME), default=0.0),
        "生產工時天數": _to_float(row.get(BOM_COL_WORK_DAYS), default=0.0),
        "供應商": str(row.get(BOM_COL_SUPPLIER, "") or "").strip(),
        "備註": str(row.get(BOM_COL_MEMO, "") or "").strip(),
    })
  return bom_map


def load_bom_from_excel(path):
  """讀取人工維護的「商品組合明細」Excel（過渡期方案／Google Sheets 未設
  定時的備援），因 A1 API 沒有組合明細查詢端點，只能這樣手動補資料。
  """
  if not path or not os.path.exists(path):
    print(f"未找到商品組合明細 Excel（{path}），商品組合資訊頁籤將只顯示型態、無子件明細")
    return {}

  try:
    # dtype=str 是關鍵：避免品號（例如 011109900001）被 pandas 當數字，
    # 把前導 0 吃掉
    df = pd.read_excel(path, sheet_name=BOM_EXCEL_SHEET_NAME, dtype=str)
  except Exception as e:
    print(f"讀取商品組合明細 Excel 失敗：{e}")
    return {}

  bom_map = _parse_bom_records(df.fillna("").to_dict("records"))
  print(
      f"商品組合明細 Excel 讀取完成：共 {len(bom_map)} 個主件品號、"
      f"{sum(len(v) for v in bom_map.values())} 筆子件關係"
  )
  return bom_map


def load_bom_from_google_sheet():
  """讀取 Google Sheet 版的「商品組合明細」（正式串接方式）。

  未設定 GOOGLE_SHEETS_CREDENTIALS_JSON / BOM_GOOGLE_SHEET_ID 任一環境變數
  時，回傳 None（代表「沒有要用 Google Sheets」，呼叫端會自動退回 Excel），
  跟「有設定但讀取失敗」（回傳空 dict）要分開，才不會誤判成「這個品項真
  的沒有子件」。

  設定步驟：
  1. Google Cloud Console 建立服務帳號，下載 JSON 金鑰
  2. 把金鑰 JSON 的完整內容存進環境變數 GOOGLE_SHEETS_CREDENTIALS_JSON
  3. 把該服務帳號的 email（金鑰 JSON 裡的 client_email）加入 Google
     Sheet 的「共用」名單，權限「檢視者」即可
  4. 設定 BOM_GOOGLE_SHEET_ID（Sheet 網址中 /d/ 與 /edit 中間那串）
  5. 分頁名稱要跟 BOM_GOOGLE_SHEET_TAB（預設「組合明細」）一致，
     欄位標題要跟本檔案裡的 BOM_COL_* 常數完全一致
  6. requirements.txt 需要加上 gspread、google-auth
  """
  if not GOOGLE_SHEETS_CREDENTIALS_JSON or not BOM_GOOGLE_SHEET_ID:
    return None

  try:
    import gspread
    from google.oauth2.service_account import Credentials

    creds_info = json.loads(GOOGLE_SHEETS_CREDENTIALS_JSON)
    creds = Credentials.from_service_account_info(
        creds_info,
        scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
    )
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(BOM_GOOGLE_SHEET_ID)
    ws = sh.worksheet(BOM_GOOGLE_SHEET_TAB)
    records = ws.get_all_records()
  except Exception as e:
    print(f"讀取 Google Sheet 商品組合明細失敗，暫時改用本機 Excel 備援：{e}")
    return {}

  bom_map = _parse_bom_records(records)
  print(
      f"Google Sheet 商品組合明細讀取完成：共 {len(bom_map)} 個主件品號、"
      f"{sum(len(v) for v in bom_map.values())} 筆子件關係"
  )
  return bom_map


def load_bom_data():
  """統一入口：優先讀 Google Sheets，沒設定或設定但失敗時自動退回本機
  Excel（過渡期相容），回傳 (bom_map, 資料來源標籤)。
  """
  sheet_result = load_bom_from_google_sheet()
  if sheet_result is not None:
    return sheet_result, "Google Sheets"
  return load_bom_from_excel(BOM_EXCEL_PATH), "本機 Excel（尚未設定 Google Sheets）"


# -------------------------------------------------------------------------
# 興聖集團旗下分公司清單（右上角切換用）
# 目前僅「海濤客食品工業(股)公司」已完成 A1 API 串接，其餘分公司頁面預留、
# 之後陸續串接時只要比照 render_hai_tao_ke_page() 的寫法為它們各自建立
# render_xxx_page() 即可。
# -------------------------------------------------------------------------
COMPANIES = ["興聖(股)公司", "海濤客食品工業(股)公司", "容鴻(股)公司", "芙萊柏(股)公司"]
ACTIVE_COMPANY_LABEL = "海濤客食品工業(股)公司"


# 初始化全域狀態
initial_df, initial_whs, initial_cats, initial_items_map = fetch_all_a1_inventory()
initial_bom_map, initial_bom_source = load_bom_data()
app_state = {
    "df": initial_df,
    "items_map": initial_items_map,
    "bom_map": initial_bom_map,
    "bom_source": initial_bom_source,
    "lot_nos": [],  # 批號資料改成頁籤點開時才抓（避免每次啟動都多打一支 API）
    # 6.1 同步狀態／錯誤日誌：每次 handle_sync 執行後會 append 一筆，
    # 只保留最新 20 筆，供「系統設定」頁籤顯示
    "sync_log": [],
    "last_sync_time": None,
    "last_sync_status": "尚未手動同步過（顯示的是啟動時自動抓取的資料）",
    # 6.3 參數設定：目前只存在記憶體，服務重啟就會回到預設值；
    # 之後若要長期保存，建議也存到 Google Sheets 的一個「系統參數」分頁
    "settings": {
        "low_stock_alert_ratio": 1.0,  # 庫存 <= 安全庫存 * 此比例 視為風險
        "default_lead_time_days": 7,   # 沒在 BOM 填前置天數時的預設值
    },
    "warehouses": (
        initial_whs
        if initial_whs
        else [
            "食品廠鳳仁倉",
            "即期品/報廢倉",
            "供應商-原料倉",
            "永福倉",
            "北仁街辦公室",
            "小琉球現場",
            "供應商-耗材倉",
        ]
    ),
    "categories": (
        initial_cats
        if initial_cats
        else [
            "(海濤客)_成品11",
            "(海濤客)_原料21",
            "(海濤客)_物料31",
            "(海濤客)_組合品61",
            "(海濤客)_費用71",
            "(海濤客)_代工含料81",
            "(海濤客)_限定組合99",
        ]
    ),
}

# -------------------------------------------------------------------------
# 2. NiceGUI 網頁介面設計
# -------------------------------------------------------------------------


@ui.page("/")
def inventory_dashboard():
  ui.add_head_html("""
        <style>
            body { background-color: #f7f6f2; color: #1a1a1a; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }
            .q-table__container { background-color: #ffffff !important; border: 1px solid #e2e1dc; border-radius: 0px; box-shadow: none !important; }
            .q-table th { color: #555555 !important; font-weight: 700 !important; font-size: 13px; border-bottom: 2px solid #1a1a1a !important; }
            .q-table td { color: #1a1a1a !important; border-bottom: 1px solid #eeede8 !important; }
            .sync-btn { background-color: #5bc0be !important; color: #ffffff !important; font-weight: 700; }
            .q-tabs { border-bottom: 1px solid #e2e1dc; }
            .q-tab { color: #777777 !important; font-weight: 700; text-transform: none; }
            .q-tab--active { color: #1a1a1a !important; }
            .q-tab-indicator { background: #5bc0be !important; }
        </style>
    """)

  # -----------------------------------------------------------------------
  # 右上角分公司切換
  # -----------------------------------------------------------------------
  content_container = ui.column().classes("w-full")

  def render_placeholder_company(name):
    content_container.clear()
    with content_container:
      with ui.column().classes(
          "w-full p-8 max-w-[1600px] mx-auto items-center justify-center"
      ):
        with ui.card().classes(
            "w-full p-16 bg-white border border-[#e2e1dc] shadow-none"
            " rounded-none text-center"
        ):
          ui.label(name).classes("text-lg font-bold text-zinc-900 mb-2")
          ui.label("此分公司尚未串接 A1 API，敬請期待").classes(
              "text-sm text-zinc-500"
          )

  def render_hai_tao_ke_page():
    content_container.clear()
    with content_container:
      # refs 用來讓 handle_sync 在同步完資料後，能回頭更新各頁籤內的
      # 篩選選單與表格／卡片內容（各頁籤的 UI 元件是在下面才建立的，
      # 但 Python closure 只在「實際呼叫時」才查找變數，所以這裡先寫
      # handle_sync 沒問題，執行當下 refs 早已被填好）。
      refs = {}

      def handle_sync():
        sync_time = datetime.now()
        try:
          df, whs, cats, items_map = fetch_all_a1_inventory()
          app_state["df"] = df
          app_state["items_map"] = items_map
          if whs:
            app_state["warehouses"] = whs
          if cats:
            app_state["categories"] = cats

          is_mock = API_KEY == "" or API_PASSWORD == ""
          status = "成功（防呆資料，未設定 A1 憑證）" if is_mock else "成功"
        except Exception as e:  # 保底：同步過程任何未預期錯誤都不讓頁面掛掉
          status = f"失敗：{e}"

        app_state["last_sync_time"] = sync_time
        app_state["last_sync_status"] = status
        app_state["sync_log"].insert(
            0,
            {
                "時間": sync_time.strftime("%Y-%m-%d %H:%M:%S"),
                "結果": status,
                "庫存筆數": len(app_state["df"]),
            },
        )
        app_state["sync_log"] = app_state["sync_log"][:20]  # 只留最新 20 筆

        if "wh_select" in refs:
          refs["wh_select"].options = ["全部倉庫"] + app_state["warehouses"]
        if "cat_select" in refs:
          refs["cat_select"].options = ["全部分類"] + app_state["categories"]
        if "cat_select_p" in refs:
          refs["cat_select_p"].options = ["全部分類"] + app_state["categories"]

        ui.notify(
            f"已成功從鼎新 A1 API 同步 {COMPANY_NAME} 最新庫存資料！"
            f"（{sync_time.strftime('%H:%M:%S')}）",
            color="positive",
        )
        for ref_key in (
            "update_inventory_table",
            "update_products_grid",
            "update_combo_list",
            "update_dashboard",
            "update_sync_log",
            "update_procurement_list",
        ):
          if ref_key in refs:
            refs[ref_key]()

      with ui.column().classes("w-full p-8 max-w-[1600px] mx-auto"):
        with ui.row().classes("w-full items-center justify-between mb-4"):
          ui.label(ACTIVE_COMPANY_LABEL).classes(
              "text-lg font-bold text-zinc-900 tracking-wide"
          )
          ui.button("同步 A1 最新庫存", on_click=handle_sync).classes(
              "sync-btn px-4 py-2 text-xs rounded-none"
          )

        with ui.tabs().classes("w-full") as page_tabs:
          tab_dashboard = ui.tab("1. 儀表板與即時預警")
          tab_products_group = ui.tab("2. 商品與組合管理")
          tab_orders = ui.tab("3. 訂單與出貨管理")
          tab_production = ui.tab("4. 生產與包裝排程")
          tab_procurement = ui.tab("5. 採購分析與決策支援")
          tab_settings = ui.tab("6. 系統設定與同步管理")

        with ui.tab_panels(page_tabs, value=tab_dashboard).classes(
            "w-full bg-transparent"
        ):
          # ==================================================
          # 1. 儀表板與即時預警中心
          # ==================================================
          with ui.tab_panel(tab_dashboard):
            with ui.card().classes(
                "w-full p-6 bg-white border border-[#e2e1dc] shadow-none"
                " rounded-none"
            ):
              with ui.row().classes(
                  "w-full p-3 mb-4 bg-[#e8f6f5] border border-[#bfe6e3]"
              ):
                ui.label(
                    "目前可算的是「現有庫存 vs A1 商品主檔的安全存量"
                    "(SafetyStock)」——這是 A1 本身就有、但原本沒人在用"
                    "的欄位。「未來 7/14/30 天出貨預估」「今日預計出貨"
                    "訂單數」需要訂單/通路需求資料，A1 目前的 API 只能"
                    "上傳訂單、無法查詢，這塊等通路需求的 Google Sheets "
                    "串接好之後再補上，屆時這裡會自動更新，不用改版面。"
                ).classes("text-xs text-teal-800")

              dashboard_kpi_row = ui.row().classes("w-full gap-4 mb-6 flex-wrap")
              dashboard_alert_label = ui.label().classes(
                  "text-sm font-bold text-zinc-700 mb-2"
              )
              dashboard_alert_container = ui.column().classes("w-full")

              def update_dashboard():
                dashboard_kpi_row.clear()
                dashboard_alert_container.clear()

                df = app_state["df"].copy()
                items_map = app_state.get("items_map", {})
                settings = app_state["settings"]

                # 依品號彙總庫存（同一品號在不同倉庫的加總）
                if not df.empty:
                  stock_by_item = df.groupby("品號", as_index=False)[
                      "庫存數量"
                  ].sum()
                  stock_lookup = dict(
                      zip(stock_by_item["品號"], stock_by_item["庫存數量"])
                  )
                else:
                  stock_lookup = {}

                total_value = 0.0
                risk_rows = []
                for item_id, info in items_map.items():
                  safety_stock = info.get("SafetyStock")
                  try:
                    safety_stock = float(safety_stock)
                  except (TypeError, ValueError):
                    safety_stock = 0.0
                  current_stock = stock_lookup.get(item_id, 0.0)
                  unit_cost = info.get("StdPurPrice") or 0
                  try:
                    total_value += float(current_stock) * float(unit_cost)
                  except (TypeError, ValueError):
                    pass

                  threshold = safety_stock * settings["low_stock_alert_ratio"]
                  if safety_stock > 0 and current_stock <= threshold:
                    risk_rows.append({
                        "品號": item_id,
                        "品名": info.get("Name"),
                        "目前庫存": current_stock,
                        "安全存量": safety_stock,
                        "缺口": round(safety_stock - current_stock, 2),
                    })

                risk_rows.sort(key=lambda r: r["缺口"], reverse=True)

                with dashboard_kpi_row:
                  kpi_cards = [
                      ("集團／本公司庫存總值（估）", f"NT$ {total_value:,.0f}"),
                      ("低於安全庫存品項數", f"{len(risk_rows)} 項"),
                      (
                          "建議立即採購品項",
                          f"{sum(1 for r in risk_rows if r['缺口'] > 0)} 項",
                      ),
                      (
                          "今日預計出貨訂單數",
                          "－（待訂單資料源）",
                      ),
                  ]
                  for label, value in kpi_cards:
                    with ui.column().classes(
                        "bg-[#f7f6f2] border border-[#e2e1dc] p-4 min-w-[200px]"
                        " flex-1"
                    ):
                      ui.label(label).classes("text-xs text-zinc-500 mb-1")
                      ui.label(value).classes(
                          "text-xl font-black text-zinc-900"
                      )

                dashboard_alert_label.text = (
                    f"缺貨/採購風險清單（依「缺口」由大到小排序，"
                    f"共 {len(risk_rows)} 項）"
                )
                with dashboard_alert_container:
                  if not risk_rows:
                    ui.label(
                        "目前沒有品項低於安全庫存，或商品主檔尚未設定"
                        "安全存量／尚未同步資料"
                    ).classes("text-xs text-zinc-400")
                  else:
                    ui.table(
                        columns=[
                            {"name": "品號", "label": "品號", "field": "品號", "align": "left"},
                            {"name": "品名", "label": "品名", "field": "品名", "align": "left"},
                            {"name": "目前庫存", "label": "目前庫存", "field": "目前庫存"},
                            {"name": "安全存量", "label": "安全存量", "field": "安全存量"},
                            {"name": "缺口", "label": "缺口", "field": "缺口"},
                        ],
                        rows=risk_rows,
                    ).classes("w-full")

              update_dashboard()
              refs["update_dashboard"] = update_dashboard

          # ==================================================
          # 2. 商品與組合管理（原本的 3 個頁籤 + 新增批號/效期追蹤）
          # ==================================================
          with ui.tab_panel(tab_products_group):
            with ui.tabs().classes("w-full mb-2") as sub_tabs:
              tab_products = ui.tab("2.1 商品資料")
              tab_inventory = ui.tab("2.3 庫存即時查詢")
              tab_bom = ui.tab("2.2 商品組合資訊（BOM）")
              tab_lotno = ui.tab("2.3 批號／效期追蹤")

            with ui.tab_panels(sub_tabs, value=tab_products).classes(
                "w-full bg-transparent"
            ):
              # ---------------- 2.1：商品資料 ----------------
              with ui.tab_panel(tab_products):
                with ui.card().classes(
                    "w-full p-6 bg-white border border-[#e2e1dc] shadow-none"
                    " rounded-none"
                ):
                  with ui.row().classes("items-center gap-3 flex-wrap mb-2"):
                    cat_options_p = ["全部分類"] + app_state["categories"]
                    cat_select_p = ui.select(
                        options=cat_options_p, value="全部分類"
                    ).classes(
                        "bg-[#f7f6f2] text-zinc-900 rounded-none px-3 py-1"
                        " text-xs font-bold border border-[#e2e1dc]"
                    )
                    search_input_p = ui.input(
                        placeholder="輸入品號或品名關鍵字..."
                    ).classes("w-64 text-xs")

                  products_stats_label = ui.label().classes(
                      "text-xs text-zinc-500 mb-3"
                  )
                  products_container = ui.row().classes("w-full flex-wrap gap-4")

                  def update_products_grid():
                    products_container.clear()
                    df = app_state["df"].copy()

                    if df.empty:
                      products_stats_label.text = "目前尚無商品資料"
                      with products_container:
                        ui.label("請先按右上角「同步 A1 最新庫存」").classes(
                            "text-xs text-zinc-400"
                        )
                      return

                    # 舊資料（或防呆資料）可能沒有「圖片」欄位，這裡先補上避免
                    # groupby 時噴錯
                    if "圖片" not in df.columns:
                      df["圖片"] = None

                    # 依品號彙整（同一品號在不同倉庫的庫存加總，
                    # 品名/分類/單位/平均成本/圖片取第一筆即可，均來自商品主檔）
                    catalog = df.groupby("品號", as_index=False).agg({
                        "品名": "first",
                        "商品分類": "first",
                        "單位": "first",
                        "平均成本": "first",
                        "庫存數量": "sum",
                        "圖片": "first",
                    })

                    if (
                        cat_select_p.value
                        and cat_select_p.value != "全部分類"
                    ):
                      catalog = catalog[
                          catalog["商品分類"] == cat_select_p.value
                      ]

                    keyword = (search_input_p.value or "").strip()
                    if keyword:
                      mask = catalog["品號"].astype(str).str.contains(
                          keyword, case=False, na=False
                      ) | catalog["品名"].astype(str).str.contains(
                          keyword, case=False, na=False
                      )
                      catalog = catalog[mask]

                    products_stats_label.text = f"共 {len(catalog)} 項商品"

                    with products_container:
                      for _, row in catalog.iterrows():
                        item_name = safe_text(row["品名"], "(未命名商品)")
                        initial = item_name[0] if item_name else "?"
                        with ui.card().classes(
                            "w-56 p-4 bg-white border border-[#e2e1dc]"
                            " shadow-none rounded-none"
                        ):
                          # 依手冊 ItemImage[Get] 抓取的商品圖片（已轉成 base64
                          # data URI）；沒有上傳過圖片的商品，改用品名首字當
                          # 預留圖示。
                          # 注意：若整欄「圖片」都是 None，pandas 會把該欄自動
                          # 轉成 float64 的 NaN，而 Python 的 bool(nan) 是
                          # True，用 `if image_uri:` 判斷會誤把 NaN 當成「有
                          # 圖片」，把 NaN 傳進 ui.image() 造成 TypeError。
                          # 因此這裡明確要求是非空字串才視為有圖片。
                          image_uri = row.get("圖片")
                          has_image = (
                              isinstance(image_uri, str) and bool(image_uri)
                          )
                          with ui.row().classes(
                              "w-full items-center justify-center mb-2"
                          ):
                            if has_image:
                              ui.image(image_uri).classes(
                                  "w-14 h-14 rounded-full object-cover"
                              )
                            else:
                              ui.label(initial).classes(
                                  "w-14 h-14 flex items-center justify-center"
                                  " rounded-full bg-[#5bc0be] text-white text-xl"
                                  " font-black"
                              )
                          ui.label(item_name).classes(
                              "text-sm font-bold text-zinc-900 text-center"
                              " w-full truncate"
                          )
                          ui.label(f"品號：{row['品號']}").classes(
                              "text-xs text-zinc-500 text-center w-full"
                          )
                          ui.label(safe_text(row["商品分類"], "未分類")).classes(
                              "text-xs text-center w-full text-[#5bc0be]"
                              " font-bold mt-1"
                          )
                          ui.separator().classes("my-2")
                          with ui.row().classes(
                              "w-full justify-between text-xs text-zinc-700"
                          ):
                            ui.label(
                                f"庫存：{row['庫存數量']:g} {safe_text(row['單位'])}"
                            )
                            ui.label(f"成本：{row['平均成本']:.2f}")

                  cat_select_p.on_value_change(lambda e: update_products_grid())
                  search_input_p.on_value_change(
                      lambda e: update_products_grid()
                  )
                  update_products_grid()

                  refs["cat_select_p"] = cat_select_p
                  refs["update_products_grid"] = update_products_grid

              # ---------------- 2.3：庫存即時查詢 ----------------
              with ui.tab_panel(tab_inventory):
                with ui.card().classes(
                    "w-full p-6 bg-white border border-[#e2e1dc] shadow-none"
                    " rounded-none"
                ):
                  with ui.row().classes(
                      "w-full items-center justify-between mb-6 gap-4 flex-wrap"
                  ):
                    ui.label("倉庫即時庫存總表").classes(
                        "text-lg font-bold text-zinc-900 tracking-wide"
                    )

                    with ui.row().classes("items-center gap-3 flex-wrap"):
                      wh_options = ["全部倉庫"] + app_state["warehouses"]
                      wh_select = ui.select(
                          options=wh_options, value="全部倉庫"
                      ).classes(
                          "bg-[#f7f6f2] text-zinc-900 rounded-none px-3 py-1"
                          " text-xs font-bold border border-[#e2e1dc]"
                      )

                      cat_options = ["全部分類"] + app_state["categories"]
                      cat_select = ui.select(
                          options=cat_options, value="全部分類"
                      ).classes(
                          "bg-[#f7f6f2] text-zinc-900 rounded-none px-3 py-1"
                          " text-xs font-bold border border-[#e2e1dc]"
                      )

                      search_input = ui.input(
                          placeholder="輸入品號或品名關鍵字..."
                      ).classes("w-64 text-xs")

                  stats_label = ui.label().classes("text-xs text-zinc-500 mb-3")

                  with ui.row().classes(
                      "w-full items-center gap-3 mb-4 p-3 bg-[#f7f6f2] border"
                      " border-[#e2e1dc] flex-wrap"
                  ):
                    ui.label(
                        "單一品號覆核（直查 A1，不受表格篩選/分頁影響）"
                    ).classes("text-xs font-bold text-zinc-700")
                    verify_input = ui.input(placeholder="輸入完整品號...").classes(
                        "w-56 text-xs"
                    )
                    verify_result = ui.label().classes("text-xs text-zinc-700")

                    def handle_verify():
                      item_id = verify_input.value.strip()
                      if not item_id:
                        ui.notify("請輸入品號", color="warning")
                        return
                      token = get_a1_token()
                      if not token:
                        verify_result.text = "無法登入 A1，請確認 API 憑證"
                        return
                      ok, msg, rows = fetch_stock_single_item(token, item_id)
                      if not ok:
                        verify_result.text = f"查詢失敗：{msg}"
                        return
                      if not rows:
                        verify_result.text = (
                            "A1 回傳：此品號目前在任何倉庫都沒有庫存資料"
                            "（Stock 查無資料）"
                        )
                        return
                      parts = [
                          f"{r.get('WarehouseName')}: {r.get('Qty')}"
                          for r in rows
                          if r.get("WarehouseName")
                      ]
                      verify_result.text = "A1 即時庫存 → " + "，".join(parts)

                    ui.button("查詢", on_click=handle_verify).classes(
                        "sync-btn px-3 py-1 text-xs rounded-none"
                    )

                  table_container = ui.column().classes("w-full")

                  def update_inventory_table():
                    table_container.clear()
                    df = app_state["df"].copy()
                    total_before_filters = len(df)

                    # 1. 倉庫篩選
                    if wh_select.value and wh_select.value != "全部倉庫":
                      df = df[df["倉庫名稱"] == wh_select.value]

                    # 2. 分類篩選
                    if cat_select.value and cat_select.value != "全部分類":
                      df = df[df["商品分類"] == cat_select.value]

                    # 3. 關鍵字搜尋
                    keyword = (search_input.value or "").strip()
                    if keyword:
                      mask = df["品號"].astype(str).str.contains(
                          keyword, case=False, na=False
                      ) | df["品名"].astype(str).str.contains(
                          keyword, case=False, na=False
                      )
                      df = df[mask]

                    # 4. 庫存數量篩選：0 不顯示，負數（超賣/盤差）要顯示
                    before_qty_filter = len(df)
                    df = df[df["庫存數量"] != 0]
                    hidden_zero_count = before_qty_filter - len(df)

                    stats_label.text = (
                        f"同步資料共 {total_before_filters} 列（品號 x 倉庫）｜"
                        f"符合篩選條件 {before_qty_filter} 列｜"
                        f"其中庫存為 0 已隱藏 {hidden_zero_count} 列｜"
                        f"目前顯示 {len(df)} 列"
                    )

                    with table_container:
                      ui.table(
                          columns=[
                              {
                                  "name": "倉庫名稱",
                                  "label": "倉庫",
                                  "field": "倉庫名稱",
                                  "align": "left",
                              },
                              {
                                  "name": "商品分類",
                                  "label": "商品分類",
                                  "field": "商品分類",
                                  "align": "left",
                              },
                              {
                                  "name": "品號",
                                  "label": "品號",
                                  "field": "品號",
                                  "align": "left",
                              },
                              {
                                  "name": "品名",
                                  "label": "品名",
                                  "field": "品名",
                                  "align": "left",
                              },
                              {"name": "單位", "label": "單位", "field": "單位"},
                              {
                                  "name": "庫存數量",
                                  "label": "庫存數量",
                                  "field": "庫存數量",
                              },
                              {
                                  "name": "平均成本",
                                  "label": "平均成本",
                                  "field": "平均成本",
                              },
                          ],
                          rows=df.to_dict("records"),
                      ).classes("w-full")

                  wh_select.on_value_change(lambda e: update_inventory_table())
                  cat_select.on_value_change(lambda e: update_inventory_table())
                  search_input.on_value_change(
                      lambda e: update_inventory_table()
                  )

                  update_inventory_table()

                  refs["wh_select"] = wh_select
                  refs["cat_select"] = cat_select
                  refs["update_inventory_table"] = update_inventory_table

              # ---------------- 2.2：商品組合資訊（BOM） ----------------
              # 手冊 1.0.35 全文查過一遍，Items[Get] 只回傳「商品型態」
              # （1.一般商品 2.組合品-先組合再銷售 3.組合品-先銷售自動組合），
              # A1 本身的匯出功能也只有主件、沒有子件。因此子件/用量/損耗率
              # /前置天數/工時明細改由「商品組合明細」Google Sheet 補齊
              # （見 load_bom_data()；未設定 Google Sheets 時自動退回本機
              # Excel 上傳），這裡讀取後與 A1 商品主檔的組合品清單合併顯示。
              with ui.tab_panel(tab_bom):
                with ui.card().classes(
                    "w-full p-6 bg-white border border-[#e2e1dc] shadow-none"
                    " rounded-none"
                ):
                  ui.label("商品組合資訊（BOM）").classes(
                      "text-lg font-bold text-zinc-900 tracking-wide mb-2"
                  )
                  with ui.row().classes(
                      "w-full p-3 mb-4 bg-[#fff8e6] border border-[#f0dca0]"
                  ):
                    ui.label(
                        "⚠ 鼎新 A1 目前的 API／後台匯出都只有組合品「主件」，"
                        "沒有「子件＋用量＋損耗率＋前置天數」明細，這些資料"
                        "改由「商品組合明細」Google Sheet 維護（尚未設定"
                        "Google Sheets 時，暫時退回本機 Excel 上傳作為過渡）。"
                    ).classes("text-xs text-amber-800")

                  bom_source_badge = ui.label().classes(
                      "text-xs font-bold mb-2"
                  )

                  with ui.row().classes(
                      "w-full items-center gap-3 flex-wrap mb-4 p-3"
                      " bg-[#f7f6f2] border border-[#e2e1dc]"
                  ):

                    def handle_reload_bom():
                      app_state["bom_map"], app_state["bom_source"] = (
                          load_bom_data()
                      )
                      ui.notify(
                          f"已重新讀取商品組合明細（來源："
                          f"{app_state['bom_source']}）",
                          color="positive",
                      )
                      update_combo_list()

                    ui.button(
                        "重新讀取商品組合明細", on_click=handle_reload_bom
                    ).classes("sync-btn px-3 py-1 text-xs rounded-none")

                    def handle_bom_upload(e):
                      os.makedirs(
                          os.path.dirname(BOM_EXCEL_PATH), exist_ok=True
                      )
                      with open(BOM_EXCEL_PATH, "wb") as f:
                        f.write(e.content.read())
                      # 手動上傳 Excel 一律視為使用者想切回 Excel 來源，
                      # 直接讀 Excel（不受 Google Sheets 是否有設定影響），
                      # 方便在 Google Sheets 串接完成前先用這個過渡方案
                      app_state["bom_map"] = load_bom_from_excel(BOM_EXCEL_PATH)
                      app_state["bom_source"] = "本機 Excel（手動上傳）"
                      ui.notify(
                          f"已上傳並套用最新的商品組合明細（{e.name}）",
                          color="positive",
                      )
                      update_combo_list()

                    ui.upload(
                        label="或上傳商品組合明細 Excel（.xlsx，過渡用）",
                        on_upload=handle_bom_upload,
                        auto_upload=True,
                    ).props('accept=".xlsx"').classes("max-w-sm text-xs")

                  with ui.row().classes(
                      "items-center gap-3 flex-wrap mb-2 justify-between w-full"
                  ):
                    combo_search_input = ui.input(
                        placeholder="輸入品號或品名關鍵字..."
                    ).classes("w-64 text-xs")

                  combo_stats_label = ui.label().classes(
                      "text-xs text-zinc-500 mb-3"
                  )
                  combo_list_container = ui.column().classes("w-full gap-2")

                  def update_combo_list():
                    combo_list_container.clear()
                    items_map = app_state.get("items_map", {})
                    bom_map = app_state.get("bom_map", {})

                    bom_source_badge.text = (
                        f"目前資料來源：{app_state.get('bom_source', '未知')}"
                    )

                    combo_entries = [
                        (item_id, info)
                        for item_id, info in items_map.items()
                        if str(info.get("Type")) in ("2", "3")
                    ]

                    keyword = (combo_search_input.value or "").strip().lower()
                    if keyword:
                      combo_entries = [
                          (item_id, info)
                          for item_id, info in combo_entries
                          if keyword in str(item_id).lower()
                          or keyword in str(info.get("Name", "")).lower()
                      ]

                    filled_count = sum(
                        1 for item_id, _ in combo_entries if bom_map.get(item_id)
                    )
                    combo_stats_label.text = (
                        f"共 {len(combo_entries)} 項組合品｜其中 {filled_count} 項"
                        f"已補齊子件明細"
                    )

                    with combo_list_container:
                      if not combo_entries:
                        ui.label(
                            "目前沒有標記為組合品的商品，或尚未同步商品資料"
                        ).classes("text-xs text-zinc-400")

                      for item_id, info in combo_entries:
                        components = bom_map.get(item_id, [])
                        type_label = ITEM_TYPE_LABELS.get(
                            str(info.get("Type")), str(info.get("Type"))
                        )
                        header_suffix = (
                            f"（{len(components)} 項子件）"
                            if components
                            else "（尚未補齊子件明細）"
                        )
                        with ui.expansion(
                            f"{item_id}｜{info.get('Name', '')}｜{type_label}"
                            f"{header_suffix}",
                            icon="inventory_2",
                        ).classes(
                            "w-full border border-[#e2e1dc] text-sm"
                        ):
                          if components:
                            ui.table(
                                columns=[
                                    {
                                        "name": "子件品號",
                                        "label": "子件品號",
                                        "field": "子件品號",
                                        "align": "left",
                                    },
                                    {
                                        "name": "子件品名",
                                        "label": "子件品名",
                                        "field": "子件品名",
                                        "align": "left",
                                    },
                                    {
                                        "name": "用量",
                                        "label": "用量",
                                        "field": "用量",
                                    },
                                    {
                                        "name": "單位",
                                        "label": "單位",
                                        "field": "單位",
                                    },
                                    {
                                        "name": "損耗率",
                                        "label": "損耗率(%)",
                                        "field": "損耗率",
                                    },
                                    {
                                        "name": "採購前置天數",
                                        "label": "採購前置天數",
                                        "field": "採購前置天數",
                                    },
                                    {
                                        "name": "生產工時天數",
                                        "label": "生產工時(天)",
                                        "field": "生產工時天數",
                                    },
                                    {
                                        "name": "供應商",
                                        "label": "供應商",
                                        "field": "供應商",
                                        "align": "left",
                                    },
                                    {
                                        "name": "備註",
                                        "label": "備註",
                                        "field": "備註",
                                        "align": "left",
                                    },
                                ],
                                rows=components,
                            ).classes("w-full")
                          else:
                            ui.label(
                                "尚未在「商品組合明細」資料中新增此品號的子件"
                                "資料，請於 Google Sheets（或過渡用 Excel）"
                                "新增一列，主件品號填此品號。"
                            ).classes("text-xs text-zinc-400 p-2")

                      # 提醒：Excel 裡有填，但目前 A1 商品主檔查無此品號的主件
                      # （可能是品號打錯，或該商品已停售）
                      orphan_ids = [
                          pid for pid in bom_map if pid not in items_map
                      ]
                      if orphan_ids:
                        shown = "、".join(orphan_ids[:10])
                        more = "…" if len(orphan_ids) > 10 else ""
                        with ui.row().classes(
                            "w-full p-3 mt-2 bg-[#fdecea] border"
                            " border-[#f5c2c0]"
                        ):
                          ui.label(
                              f"⚠ 商品組合明細中有 {len(orphan_ids)} 個主件品號，在"
                              f"目前 A1 商品主檔中查無此品號（可能是品號填錯，"
                              f"或該商品已停售）：{shown}{more}"
                          ).classes("text-xs text-red-700")

                  combo_search_input.on_value_change(lambda e: update_combo_list())
                  update_combo_list()

                  refs["update_combo_list"] = update_combo_list

              # ---------------- 2.3：批號／效期追蹤 ----------------
              with ui.tab_panel(tab_lotno):
                with ui.card().classes(
                    "w-full p-6 bg-white border border-[#e2e1dc] shadow-none"
                    " rounded-none"
                ):
                  ui.label("批號／效期追蹤").classes(
                      "text-lg font-bold text-zinc-900 tracking-wide mb-2"
                  )
                  with ui.row().classes(
                      "w-full p-3 mb-4 bg-[#e8f6f5] border border-[#bfe6e3]"
                  ):
                    ui.label(
                        "來自 A1 的 ItemLotNos[Get]（取得所有商品批號資料）。"
                        "A1 商品主檔本身沒有「保存效期」欄位，但若貴公司有"
                        "租用批號模組、且商品有啟用批號管理，每一批號會各自"
                        "記錄有效日期，可以用來做效期預警。若公司沒有租用"
                        "批號模組，這裡會是空的，屬於正常情況。"
                    ).classes("text-xs text-teal-800")

                  with ui.row().classes("items-center gap-3 mb-3"):
                    lotno_search_input = ui.input(
                        placeholder="輸入品號或品名關鍵字..."
                    ).classes("w-64 text-xs")

                    def handle_load_lot_nos():
                      token = get_a1_token()
                      if not token:
                        ui.notify("無法登入 A1，請確認 API 憑證", color="warning")
                        return
                      ok, data = fetch_all_lot_nos(token)
                      if not ok:
                        ui.notify(f"查詢批號資料失敗：{data}", color="negative")
                        return
                      app_state["lot_nos"] = data if isinstance(data, list) else []
                      ui.notify(
                          f"已載入 {len(app_state['lot_nos'])} 筆批號資料",
                          color="positive",
                      )
                      update_lotno_table()

                    ui.button(
                        "查詢批號資料", on_click=handle_load_lot_nos
                    ).classes("sync-btn px-3 py-1 text-xs rounded-none")

                  lotno_stats_label = ui.label().classes(
                      "text-xs text-zinc-500 mb-3"
                  )
                  lotno_table_container = ui.column().classes("w-full")

                  def update_lotno_table():
                    lotno_table_container.clear()
                    rows = app_state.get("lot_nos", [])

                    keyword = (lotno_search_input.value or "").strip().lower()
                    if keyword:
                      rows = [
                          r
                          for r in rows
                          if keyword in str(r.get("ItemDetailID", "")).lower()
                          or keyword in str(r.get("ItemName", "")).lower()
                          or keyword in str(r.get("LotNo", "")).lower()
                      ]

                    lotno_stats_label.text = f"共 {len(rows)} 筆批號資料"

                    with lotno_table_container:
                      if not rows:
                        ui.label(
                            "尚未查詢，或此公司未啟用批號管理／沒有批號資料。"
                            "按上方「查詢批號資料」試試看。"
                        ).classes("text-xs text-zinc-400")
                      else:
                        ui.table(
                            columns=[
                                {"name": "ItemDetailID", "label": "品號", "field": "ItemDetailID", "align": "left"},
                                {"name": "ItemName", "label": "品名", "field": "ItemName", "align": "left"},
                                {"name": "LotNo", "label": "批號", "field": "LotNo", "align": "left"},
                                {"name": "ManufacturedDate", "label": "製造日期", "field": "ManufacturedDate"},
                                {"name": "ExpiryDate", "label": "有效日期", "field": "ExpiryDate"},
                                {"name": "IsDisable", "label": "停用否", "field": "IsDisable"},
                            ],
                            rows=rows,
                        ).classes("w-full")

                  lotno_search_input.on_value_change(lambda e: update_lotno_table())
                  update_lotno_table()

          # ==================================================
          # 3. 訂單與出貨管理
          # ==================================================
          with ui.tab_panel(tab_orders):
            with ui.card().classes(
                "w-full p-6 bg-white border border-[#e2e1dc] shadow-none"
                " rounded-none"
            ):
              ui.label("訂單與出貨管理").classes(
                  "text-lg font-bold text-zinc-900 tracking-wide mb-3"
              )
              with ui.row().classes(
                  "w-full p-4 bg-[#fdecea] border border-[#f5c2c0]"
              ):
                ui.label(
                    "⚠ 規劃中，卡在 A1 API 本身的限制：手冊裡 Orders 只有"
                    "「[Post] 上傳訂單」，沒有對應的查詢/取得端點，訂單資料"
                    "進得去、拉不出來，所以無法從 A1 直接產生「訂單需求清單」"
                    "或「出貨量統計」。這塊需要另一個資料來源——例如之前"
                    "討論的「通路填報」Google Sheet（各通路每月填報的品名/"
                    "數量），等那份 Sheet 串接完成，這裡就能顯示訂單需求清單"
                    "與出貨量趨勢，頁面架構已經留好，屆時直接接上就好。"
                ).classes("text-xs text-red-700")

          # ==================================================
          # 4. 生產與包裝排程
          # ==================================================
          with ui.tab_panel(tab_production):
            with ui.card().classes(
                "w-full p-6 bg-white border border-[#e2e1dc] shadow-none"
                " rounded-none"
            ):
              ui.label("生產與包裝排程").classes(
                  "text-lg font-bold text-zinc-900 tracking-wide mb-3"
              )
              with ui.row().classes(
                  "w-full p-4 bg-[#fdecea] border border-[#f5c2c0]"
              ):
                ui.label(
                    "⚠ 規劃中，依賴第 3 點的訂單/出貨需求資料，還沒有資料"
                    "來源。等訂單需求資料到位後，這裡會結合「商品組合資訊"
                    "(BOM)」的用量與生產工時，自動算出包裝作業的先後順序，"
                    "以及開工前需要備齊的原物料清單。"
                ).classes("text-xs text-red-700")

          # ==================================================
          # 5. 採購分析與決策支援
          # ==================================================
          with ui.tab_panel(tab_procurement):
            with ui.card().classes(
                "w-full p-6 bg-white border border-[#e2e1dc] shadow-none"
                " rounded-none"
            ):
              ui.label("採購分析與決策支援").classes(
                  "text-lg font-bold text-zinc-900 tracking-wide mb-2"
              )
              with ui.row().classes(
                  "w-full p-3 mb-4 bg-[#fff8e6] border border-[#f0dca0]"
              ):
                ui.label(
                    "目前是簡化版：淨需求只用「安全庫存 − 現有庫存」計算，"
                    "沒有把未來訂單需求（BOM 展開）算進去，因為訂單資料還沒"
                    "有來源（見第 3 點）。「供應商歷史採購單價走勢」「庫存"
                    "週轉率／滯銷品分析」需要銷售歷史資料，目前只有設計、"
                    "還沒寫進程式碼，之後接上通路填報/銷售歷史 Google Sheet "
                    "後會自動補上，不用改版面。"
                ).classes("text-xs text-amber-800")

              procurement_stats_label = ui.label().classes(
                  "text-xs text-zinc-500 mb-3"
              )
              procurement_list_container = ui.column().classes("w-full")

              def update_procurement_list():
                procurement_list_container.clear()
                df = app_state["df"].copy()
                items_map = app_state.get("items_map", {})
                bom_map = app_state.get("bom_map", {})
                default_lead_time = app_state["settings"][
                    "default_lead_time_days"
                ]

                if not df.empty:
                  stock_by_item = df.groupby("品號", as_index=False)[
                      "庫存數量"
                  ].sum()
                  stock_lookup = dict(
                      zip(stock_by_item["品號"], stock_by_item["庫存數量"])
                  )
                else:
                  stock_lookup = {}

                # 找出每個品項在 BOM 裡「作為子件」時登記的採購前置天數，
                # 沒有的話用系統預設值
                lead_time_by_child = {}
                for components in bom_map.values():
                  for comp in components:
                    lt = comp.get("採購前置天數")
                    if isinstance(lt, (int, float)) and lt > 0:
                      lead_time_by_child[comp["子件品號"]] = lt

                rows = []
                for item_id, info in items_map.items():
                  safety_stock = info.get("SafetyStock")
                  try:
                    safety_stock = float(safety_stock)
                  except (TypeError, ValueError):
                    safety_stock = 0.0
                  if safety_stock <= 0:
                    continue
                  current_stock = stock_lookup.get(item_id, 0.0)
                  net_need = round(safety_stock - current_stock, 2)
                  if net_need <= 0:
                    continue
                  rows.append({
                      "品號": item_id,
                      "品名": info.get("Name"),
                      "現有庫存": current_stock,
                      "安全庫存": safety_stock,
                      "建議採購量（簡化版）": net_need,
                      "參考前置天數": lead_time_by_child.get(
                          item_id, default_lead_time
                      ),
                  })

                rows.sort(key=lambda r: r["建議採購量（簡化版）"], reverse=True)
                procurement_stats_label.text = (
                    f"共 {len(rows)} 項建議採購品項（安全庫存 > 現有庫存）"
                )

                with procurement_list_container:
                  if not rows:
                    ui.label(
                        "目前沒有品項需要採購，或商品主檔尚未設定安全庫存"
                    ).classes("text-xs text-zinc-400")
                  else:
                    ui.table(
                        columns=[
                            {"name": c, "label": c, "field": c, "align": "left" if c in ("品號", "品名") else "right"}
                            for c in rows[0].keys()
                        ],
                        rows=rows,
                    ).classes("w-full")

              update_procurement_list()
              refs["update_procurement_list"] = update_procurement_list

          # ==================================================
          # 6. 系統設定與同步管理
          # ==================================================
          with ui.tab_panel(tab_settings):
            with ui.column().classes("w-full gap-4"):
              # ---------------- 6.1：同步狀態與日誌 ----------------
              with ui.card().classes(
                  "w-full p-6 bg-white border border-[#e2e1dc] shadow-none"
                  " rounded-none"
              ):
                ui.label("鼎新 A1 連線狀態與同步日誌").classes(
                    "text-lg font-bold text-zinc-900 tracking-wide mb-3"
                )
                sync_status_label = ui.label().classes(
                    "text-sm text-zinc-700 mb-3"
                )
                sync_log_container = ui.column().classes("w-full")

                def update_sync_log():
                  sync_status_label.text = (
                      f"最後同步時間："
                      f"{app_state['last_sync_time'].strftime('%Y-%m-%d %H:%M:%S') if app_state['last_sync_time'] else '尚未手動同步'}"
                      f"｜最後結果：{app_state['last_sync_status']}"
                  )
                  sync_log_container.clear()
                  with sync_log_container:
                    if not app_state["sync_log"]:
                      ui.label("尚無同步紀錄").classes("text-xs text-zinc-400")
                    else:
                      ui.table(
                          columns=[
                              {"name": "時間", "label": "時間", "field": "時間", "align": "left"},
                              {"name": "結果", "label": "結果", "field": "結果", "align": "left"},
                              {"name": "庫存筆數", "label": "庫存筆數", "field": "庫存筆數"},
                          ],
                          rows=app_state["sync_log"],
                      ).classes("w-full")

                update_sync_log()
                refs["update_sync_log"] = update_sync_log

              # ---------------- 6.2：Google Sheets 設定檢視 ----------------
              with ui.card().classes(
                  "w-full p-6 bg-white border border-[#e2e1dc] shadow-none"
                  " rounded-none"
              ):
                ui.label("Google Sheets 資料對應設定").classes(
                    "text-lg font-bold text-zinc-900 tracking-wide mb-3"
                )
                sheets_configured = bool(
                    GOOGLE_SHEETS_CREDENTIALS_JSON and BOM_GOOGLE_SHEET_ID
                )
                ui.label(
                    "狀態："
                    + (
                        "已設定，正在使用 Google Sheets 作為商品組合明細來源"
                        if sheets_configured
                        else "尚未設定，目前使用本機 Excel 過渡方案"
                    )
                ).classes(
                    "text-sm mb-2 "
                    + ("text-teal-700" if sheets_configured else "text-amber-700")
                )
                with ui.column().classes("gap-1"):
                  ui.label(
                      f"Sheet ID：{BOM_GOOGLE_SHEET_ID or '（未設定）'}"
                  ).classes("text-xs text-zinc-600")
                  ui.label(
                      f"工作表分頁名稱：{BOM_GOOGLE_SHEET_TAB}"
                  ).classes("text-xs text-zinc-600")
                  ui.label(
                      "服務帳號金鑰："
                      + ("已設定" if GOOGLE_SHEETS_CREDENTIALS_JSON else "（未設定）")
                  ).classes("text-xs text-zinc-600")
                ui.label(
                    "設定方式：環境變數 GOOGLE_SHEETS_CREDENTIALS_JSON（服務"
                    "帳號金鑰 JSON 內容）、BOM_GOOGLE_SHEET_ID、"
                    "BOM_GOOGLE_SHEET_TAB。三者皆設定齊全才會啟用，任一缺漏"
                    "會自動退回本機 Excel，不會讓頁面壞掉。"
                ).classes("text-xs text-zinc-500 mt-2")

              # ---------------- 6.3：參數設定 ----------------
              with ui.card().classes(
                  "w-full p-6 bg-white border border-[#e2e1dc] shadow-none"
                  " rounded-none"
              ):
                ui.label("參數設定").classes(
                    "text-lg font-bold text-zinc-900 tracking-wide mb-3"
                )
                ui.label(
                    "⚠ 這裡的設定目前只存在記憶體中，服務重啟（含 Render "
                    "重新部署或睡眠喚醒）就會回到預設值。之後若要長期保存，"
                    "建議也存進 Google Sheets 的一個「系統參數」分頁。"
                ).classes("text-xs text-amber-700 mb-3")

                with ui.row().classes("items-center gap-3 mb-3"):
                  ui.label("低庫存警戒比例（庫存 ≤ 安全庫存 × 此比例視為風險）").classes(
                      "text-sm text-zinc-700"
                  )
                  ratio_input = ui.number(
                      value=app_state["settings"]["low_stock_alert_ratio"],
                      min=0.1,
                      max=3.0,
                      step=0.1,
                  ).classes("w-24")

                with ui.row().classes("items-center gap-3 mb-3"):
                  ui.label("預設採購前置天數（BOM 未填時使用）").classes(
                      "text-sm text-zinc-700"
                  )
                  lead_time_input = ui.number(
                      value=app_state["settings"]["default_lead_time_days"],
                      min=0,
                      max=90,
                      step=1,
                  ).classes("w-24")

                def handle_save_settings():
                  app_state["settings"]["low_stock_alert_ratio"] = (
                      ratio_input.value or 1.0
                  )
                  app_state["settings"]["default_lead_time_days"] = (
                      lead_time_input.value or 0
                  )
                  ui.notify("參數已更新（僅本次服務執行期間有效）", color="positive")
                  if "update_dashboard" in refs:
                    refs["update_dashboard"]()
                  if "update_procurement_list" in refs:
                    refs["update_procurement_list"]()

                ui.button("儲存參數", on_click=handle_save_settings).classes(
                    "sync-btn px-4 py-2 text-xs rounded-none"
                )

  def handle_company_change(e):
    selected = e.value
    if selected == ACTIVE_COMPANY_LABEL:
      render_hai_tao_ke_page()
    else:
      render_placeholder_company(selected)

  with ui.row().classes(
      "w-full items-center justify-between bg-white border-b border-[#e2e1dc]"
      " px-8 py-4 sticky top-0 z-50"
  ):
    ui.label("興聖集團｜A1 智慧進銷存總管理系統").classes(
        "text-base font-black tracking-wider"
    )
    with ui.tabs(on_change=handle_company_change).props(
        "dense no-caps"
    ) as company_tabs:
      for c in COMPANIES:
        ui.tab(c)
    company_tabs.set_value(ACTIVE_COMPANY_LABEL)

  render_hai_tao_ke_page()


ui.run(port=8080, title="興聖集團 A1 智慧進銷存總管理系統", host="0.0.0.0")
