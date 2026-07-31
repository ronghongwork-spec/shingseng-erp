import base64
import concurrent.futures
import os
import sys
from datetime import datetime

from nicegui import ui
import pandas as pd
import requests

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
# 預設放在 app.py 同層的 data/ 資料夾，也可用環境變數 A1_BOM_EXCEL_PATH
# 指到其他路徑（例如掛載的網路磁碟、共用資料夾）。
BOM_EXCEL_PATH = os.environ.get(
    "A1_BOM_EXCEL_PATH",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "商品組合明細.xlsx"),
)
BOM_EXCEL_SHEET_NAME = "組合明細"


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

# 對應 build_template.py 產生的「商品組合明細_填寫範本.xlsx」欄位名稱
BOM_COL_PARENT_ID = "主件品號"
BOM_COL_PARENT_NAME = "主件品名（選填，供參考）"
BOM_COL_CHILD_ID = "子件品號"
BOM_COL_CHILD_NAME = "子件品名（選填，供參考）"
BOM_COL_QTY = "用量"
BOM_COL_UNIT = "單位（選填）"
BOM_COL_MEMO = "備註"


def load_bom_from_excel(path):
  """讀取人工維護的「商品組合明細」Excel（因 A1 API 無此明細查詢端點）

  回傳 {主件品號: [{子件品號, 子件品名, 用量, 單位, 備註}, ...]}
  找不到檔案或讀取失敗時回傳空 dict，不會讓頁面掛掉——只是「商品組合
  資訊」頁籤會顯示「尚未補齊子件明細」而已。
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

  df = df.fillna("")
  bom_map = {}
  for _, row in df.iterrows():
    parent_id = str(row.get(BOM_COL_PARENT_ID, "")).strip()
    child_id = str(row.get(BOM_COL_CHILD_ID, "")).strip()
    if not parent_id or not child_id:
      continue  # 略過空白列或只填了一半的列

    qty_raw = str(row.get(BOM_COL_QTY, "")).strip()
    try:
      qty = float(qty_raw) if qty_raw else 0.0
    except ValueError:
      qty = qty_raw  # 萬一填了非數字，原樣顯示，不擋住整批匯入

    bom_map.setdefault(parent_id, []).append({
        "子件品號": child_id,
        "子件品名": str(row.get(BOM_COL_CHILD_NAME, "")).strip(),
        "用量": qty,
        "單位": str(row.get(BOM_COL_UNIT, "")).strip(),
        "備註": str(row.get(BOM_COL_MEMO, "")).strip(),
    })

  print(
      f"商品組合明細 Excel 讀取完成：共 {len(bom_map)} 個主件品號、"
      f"{sum(len(v) for v in bom_map.values())} 筆子件關係"
  )
  return bom_map


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
initial_bom_map = load_bom_from_excel(BOM_EXCEL_PATH)
app_state = {
    "df": initial_df,
    "items_map": initial_items_map,
    "bom_map": initial_bom_map,
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
        df, whs, cats, items_map = fetch_all_a1_inventory()
        app_state["df"] = df
        app_state["items_map"] = items_map
        if whs:
          app_state["warehouses"] = whs
        if cats:
          app_state["categories"] = cats

        if "wh_select" in refs:
          refs["wh_select"].options = ["全部倉庫"] + app_state["warehouses"]
        if "cat_select" in refs:
          refs["cat_select"].options = ["全部分類"] + app_state["categories"]
        if "cat_select_p" in refs:
          refs["cat_select_p"].options = ["全部分類"] + app_state["categories"]

        ui.notify(
            f"已成功從鼎新 A1 API 同步 {COMPANY_NAME} 最新庫存資料！"
            f"（{datetime.now().strftime('%H:%M:%S')}）",
            color="positive",
        )
        if "update_inventory_table" in refs:
          refs["update_inventory_table"]()
        if "update_products_grid" in refs:
          refs["update_products_grid"]()
        if "update_combo_list" in refs:
          refs["update_combo_list"]()

      with ui.column().classes("w-full p-8 max-w-[1600px] mx-auto"):
        with ui.row().classes("w-full items-center justify-between mb-4"):
          ui.label(ACTIVE_COMPANY_LABEL).classes(
              "text-lg font-bold text-zinc-900 tracking-wide"
          )
          ui.button("同步 A1 最新庫存", on_click=handle_sync).classes(
              "sync-btn px-4 py-2 text-xs rounded-none"
          )

        with ui.tabs().classes("w-full") as page_tabs:
          tab_products = ui.tab("商品資料")
          tab_inventory = ui.tab("倉庫即時庫存總表")
          tab_bom = ui.tab("商品組合資訊")

        with ui.tab_panels(page_tabs, value=tab_products).classes(
            "w-full bg-transparent"
        ):
          # ---------------- 頁籤 1：商品資料 ----------------
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
                    item_name = row["品名"] or "(未命名商品)"
                    initial = item_name[0] if item_name else "?"
                    with ui.card().classes(
                        "w-56 p-4 bg-white border border-[#e2e1dc]"
                        " shadow-none rounded-none"
                    ):
                      # 依手冊 ItemImage[Get] 抓取的商品圖片（已轉成 base64
                      # data URI）；沒有上傳過圖片的商品，改用品名首字當
                      # 預留圖示。
                      image_uri = row.get("圖片")
                      with ui.row().classes(
                          "w-full items-center justify-center mb-2"
                      ):
                        if image_uri:
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
                      ui.label(row["商品分類"] or "未分類").classes(
                          "text-xs text-center w-full text-[#5bc0be]"
                          " font-bold mt-1"
                      )
                      ui.separator().classes("my-2")
                      with ui.row().classes(
                          "w-full justify-between text-xs text-zinc-700"
                      ):
                        ui.label(
                            f"庫存：{row['庫存數量']:g} {row['單位'] or ''}"
                        )
                        ui.label(f"成本：{row['平均成本']:.2f}")

              cat_select_p.on_value_change(lambda e: update_products_grid())
              search_input_p.on_value_change(
                  lambda e: update_products_grid()
              )
              update_products_grid()

              refs["cat_select_p"] = cat_select_p
              refs["update_products_grid"] = update_products_grid

          # ---------------- 頁籤 2：倉庫即時庫存總表 ----------------
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

          # ---------------- 頁籤 3：商品組合資訊 ----------------
          # 手冊 1.0.35 全文查過一遍，Items[Get] 只回傳「商品型態」
          # （1.一般商品 2.組合品-先組合再銷售 3.組合品-先銷售自動組合），
          # A1 本身的匯出功能也只有主件、沒有子件。因此子件/用量明細改由
          # 人工維護的「商品組合明細」Excel 補齊（見 BOM_EXCEL_PATH），
          # 這裡讀取後與 A1 商品主檔的組合品清單合併顯示。
          with ui.tab_panel(tab_bom):
            with ui.card().classes(
                "w-full p-6 bg-white border border-[#e2e1dc] shadow-none"
                " rounded-none"
            ):
              ui.label("商品組合資訊").classes(
                  "text-lg font-bold text-zinc-900 tracking-wide mb-2"
              )
              with ui.row().classes(
                  "w-full p-3 mb-4 bg-[#fff8e6] border border-[#f0dca0]"
              ):
                ui.label(
                    "⚠ 鼎新 A1 目前的 API／後台匯出都只有組合品「主件」，"
                    "沒有「子件＋用量」明細，因此子件資訊改由人工維護的"
                    "「商品組合明細」Excel 補齊，系統會自動讀取合併顯示。"
                    "填寫範本請洽系統管理員索取，或參考先前提供的"
                    "「商品組合明細_填寫範本.xlsx」。"
                ).classes("text-xs text-amber-800")

              with ui.row().classes(
                  "items-center gap-3 flex-wrap mb-2 justify-between w-full"
              ):
                combo_search_input = ui.input(
                    placeholder="輸入品號或品名關鍵字..."
                ).classes("w-64 text-xs")

                def handle_reload_bom_excel():
                  app_state["bom_map"] = load_bom_from_excel(BOM_EXCEL_PATH)
                  ui.notify(
                      f"已重新讀取商品組合明細 Excel（{BOM_EXCEL_PATH}）",
                      color="positive",
                  )
                  update_combo_list()

                ui.button(
                    "重新載入 Excel", on_click=handle_reload_bom_excel
                ).classes("sync-btn px-3 py-1 text-xs rounded-none")

              combo_stats_label = ui.label().classes(
                  "text-xs text-zinc-500 mb-3"
              )
              combo_list_container = ui.column().classes("w-full gap-2")

              def update_combo_list():
                combo_list_container.clear()
                items_map = app_state.get("items_map", {})
                bom_map = app_state.get("bom_map", {})

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
                    f"已在 Excel 補齊子件明細"
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
                        else "（尚未在 Excel 補齊子件明細）"
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
                            "尚未在「商品組合明細」Excel 中新增此品號的子件"
                            "資料，請於範本新增一列（主件品號填此品號）。"
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
                          f"⚠ Excel 中有 {len(orphan_ids)} 個主件品號，在"
                          f"目前 A1 商品主檔中查無此品號（可能是品號填錯，"
                          f"或該商品已停售）：{shown}{more}"
                      ).classes("text-xs text-red-700")

              combo_search_input.on_value_change(lambda e: update_combo_list())
              update_combo_list()

              refs["update_combo_list"] = update_combo_list

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
