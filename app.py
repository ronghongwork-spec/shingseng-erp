import base64
import concurrent.futures
import io
import json
import math
import os
import secrets
import sys
from collections import defaultdict
from datetime import datetime, timedelta

from nicegui import app, ui
import pandas as pd
import requests
from starlette.responses import Response


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


def ceil_qty(value, default=0):
  """庫存/銷量這類「數量」欄位統一無條件進位成整數（例如 12.3 顯示成 13），
  避免小數點讓人誤以為可以訂購零點幾件、或看起來像是算錯了。
  用 math.ceil 而不是四捨五入，是刻意的——採購/庫存寧可估多一點，不要
  估少導致真的缺貨。
  """
  if value is None:
    return default
  try:
    if pd.isna(value):
      return default
  except (TypeError, ValueError):
    pass
  try:
    return int(math.ceil(float(value)))
  except (TypeError, ValueError):
    return default


def _safe_sheet_name(name):
  """Excel 工作表名稱不能包含 \\ / ? * [ ] : ，超過 31 字元也會報錯。
  之前有個 KPI 彈窗標題含有「/」（成品＋原料/子件），沒清過就直接拿去
  當 sheet_name，會讓 openpyxl 丟例外、匯出整個失敗——這裡統一清乾淨，
  避免同類問題在其他地方重演。
  """
  for ch in '\\/?*[]:':
    name = name.replace(ch, "_")
  return name[:31] or "工作表1"


def rows_to_xlsx_bytes(rows, sheet_name="工作表1"):
  """把 list[dict] 轉成 xlsx 檔案的 bytes，供 ui.download() 直接下載，
  不落地寫檔（Render 的檔案系統是暫時的，用記憶體 buffer 比較乾淨）。
  """
  df = pd.DataFrame(rows)
  buffer = io.BytesIO()
  with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name=_safe_sheet_name(sheet_name))
  return buffer.getvalue()


def multi_sheet_xlsx_bytes(sheets):
  """sheets: {分頁名稱: rows(list[dict])}，打包成一個多分頁 xlsx 的
  bytes，一次匯出「總表」＋「分通路表」這種需要好幾個分頁的情境。
  """
  buffer = io.BytesIO()
  with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
    for name, rows in sheets.items():
      df = pd.DataFrame(rows)
      df.to_excel(writer, index=False, sheet_name=_safe_sheet_name(name))
  return buffer.getvalue()

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

# -------------------------------------------------------------------------
# 興聖(股)公司｜SHOPLINE 官網訂單
# 文件：https://open-api.docs.shoplineapp.com/docs/search-orders
# .env 需設定：
#   SHOPLINE_XINGSHENG_ACCESS_TOKEN=（SHOPLINE後台管理員設定產生的token）
#   SHOPLINE_XINGSHENG_USER_AGENT=（識別用字串，無嚴格規定，已實測 "Xingsheng-ERP" 可用）
# -------------------------------------------------------------------------
SHOPLINE_API_DOMAIN = "https://open.shopline.io"
SHOPLINE_XINGSHENG_ACCESS_TOKEN = os.environ.get("SHOPLINE_XINGSHENG_ACCESS_TOKEN", "")
SHOPLINE_XINGSHENG_USER_AGENT = os.environ.get("SHOPLINE_XINGSHENG_USER_AGENT", "")


def fetch_shopline_pending_orders(access_token, user_agent):
  """興聖SHOPLINE官網：抓全部「待處理」(status=pending)訂單，自動翻頁，
  回傳簡化過、給表格顯示用的欄位（不是原始API的完整資料）。
  回傳 (rows, error_message)；成功時 error_message 是 None。
  """
  if not access_token or not user_agent:
    return None, "尚未設定 SHOPLINE_XINGSHENG_ACCESS_TOKEN / SHOPLINE_XINGSHENG_USER_AGENT"
  try:
    all_orders = []
    page = 1
    while True:
      resp = requests.get(
          f"{SHOPLINE_API_DOMAIN}/v1/orders/search",
          params={"status": "pending", "per_page": 50, "page": page},
          headers={
              "accept": "application/json",
              "authorization": f"Bearer {access_token}",
              "User-Agent": user_agent,
          },
          timeout=REQUEST_TIMEOUT,
      )
      resp.raise_for_status()
      body = resp.json()
      all_orders.extend(body.get("items", []) or [])
      pagination = body.get("pagination", {}) or {}
      total_pages = pagination.get("total_pages", 1) or 1
      if page >= total_pages:
        break
      page += 1

    rows = []
    for o in all_orders:
      created_raw = o.get("created_at", "") or ""
      try:
        dt_utc = datetime.fromisoformat(created_raw.replace("Z", "+00:00"))
        created_display = (dt_utc + timedelta(hours=8)).strftime("%Y-%m-%d %H:%M")
      except (ValueError, TypeError):
        created_display = created_raw
      delivery = o.get("order_delivery") or {}
      rows.append({
          "訂單編號": o.get("order_number") or o.get("id"),
          "客戶": o.get("customer_name") or "",
          "金額": (o.get("total") or {}).get("label") or "",
          "建立時間": created_display,
          "物流狀態": delivery.get("delivery_status") or "",
      })
    return rows, None
  except Exception as e:
    return None, str(e)

# -------------------------------------------------------------------------
# 1.5. 內部員工登入保護（HTTP Basic Auth）
#    整個網站（含所有頁籤、API、靜態檔案）都會被擋住，瀏覽器打開時會跳出
#    原生的「輸入帳號密碼」對話框，帳密正確才放行。適合「全公司共用一組
#    帳密」這種內部系統，不需要個別員工帳號。
#
#    部署到 Render 時請在後台「Environment」設定這兩個環境變數：
#      BASIC_AUTH_USERNAME=你要用的帳號
#      BASIC_AUTH_PASSWORD=你要用的密碼
#    兩個都設定了才會啟用保護；只要有一個沒設，網站會維持開放不擋（本機
#    開發測試時通常不會設這兩個變數，才不會每次都要輸入密碼）。
# -------------------------------------------------------------------------
BASIC_AUTH_USERNAME = os.environ.get("BASIC_AUTH_USERNAME", "")
BASIC_AUTH_PASSWORD = os.environ.get("BASIC_AUTH_PASSWORD", "")

if not BASIC_AUTH_USERNAME or not BASIC_AUTH_PASSWORD:
  print(
      "警告：尚未設定環境變數 BASIC_AUTH_USERNAME / BASIC_AUTH_PASSWORD，"
      "網站目前沒有密碼保護，任何人有網址都能看到。",
      file=sys.stderr,
  )


class BasicAuthMiddleware:
  """檢查每一個一般網頁請求的 Authorization header，帳密不符就回401要求重新輸入。

  刻意寫成最底層的 ASGI middleware（而不是 Starlette 的
  BaseHTTPMiddleware），是因為 BaseHTTPMiddleware 會把每個請求包裝成
  request/response 物件再轉發，這個包裝過程跟 NiceGUI 用來即時更新畫面
  的 WebSocket 連線不相容，實測會直接讓連線壞掉、整頁噴 500，而不是照
  我們要的邏輯回401。這裡改成只在 scope["type"] == "http"（一般網頁請求）
  時才檢查帳密，WebSocket 跟其他類型的連線完全不經過這段邏輯、直接放行
  給 NiceGUI 自己處理。
  """

  def __init__(self, asgi_app):
    self.asgi_app = asgi_app

  async def __call__(self, scope, receive, send):
    if scope["type"] != "http" or not BASIC_AUTH_USERNAME or not BASIC_AUTH_PASSWORD:
      await self.asgi_app(scope, receive, send)
      return

    headers = dict(scope.get("headers") or [])
    auth_header = headers.get(b"authorization", b"").decode("latin-1")

    authorized = False
    if auth_header.startswith("Basic "):
      try:
        decoded = base64.b64decode(auth_header[6:]).decode("utf-8")
        username, _, password = decoded.partition(":")
        # 用 secrets.compare_digest 而不是 == ，避免帳密比對時間差被用來
        # 猜測密碼內容（timing attack），這裡資安等級不需要到這麼高，但
        # 反正代價很低，順手做好。
        authorized = secrets.compare_digest(
            username, BASIC_AUTH_USERNAME
        ) and secrets.compare_digest(password, BASIC_AUTH_PASSWORD)
      except Exception:
        authorized = False

    if authorized:
      await self.asgi_app(scope, receive, send)
      return

    response = Response(
        content="請輸入帳號密碼才能瀏覽此系統",
        status_code=401,
        headers={"WWW-Authenticate": 'Basic realm="Internal System"'},
    )
    await response(scope, receive, send)


app.add_middleware(BasicAuthMiddleware)

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

# ---- Google Sheets 串接（正式方案，逐步取代本機 Excel 上傳） ----
# 需要在 Google Cloud 建立一個服務帳號(Service Account)，下載其 JSON 金鑰，
# 並把該服務帳號的 email 加入 Google Sheet 的「共用」名單（唯讀權限即可）。
# 三份資料（BOM／訂單資訊／銷售歷史）建議放在「同一份」Google Sheet 的
# 三個不同分頁，只設定一組 GOOGLE_SHEET_ID 即可，分頁名稱各自對應。
GOOGLE_SHEETS_CREDENTIALS_JSON = os.environ.get(
    "GOOGLE_SHEETS_CREDENTIALS_JSON", ""
)  # 服務帳號 JSON 金鑰的「內容」（整包貼進環境變數），不是檔案路徑
GOOGLE_SHEET_ID = os.environ.get(
    "GOOGLE_SHEET_ID", os.environ.get("BOM_GOOGLE_SHEET_ID", "")
)  # 相容舊的 BOM_GOOGLE_SHEET_ID 命名，新專案請直接用 GOOGLE_SHEET_ID
BOM_GOOGLE_SHEET_TAB = os.environ.get("BOM_GOOGLE_SHEET_TAB", "BOM表")
ORDERS_GOOGLE_SHEET_TAB = os.environ.get("ORDERS_GOOGLE_SHEET_TAB", "訂單資訊")
SALES_HISTORY_GOOGLE_SHEET_TAB = os.environ.get(
    "SALES_HISTORY_GOOGLE_SHEET_TAB", "銷售歷史"
)
RECEIVING_GOOGLE_SHEET_TAB = os.environ.get(
    "RECEIVING_GOOGLE_SHEET_TAB", "進貨明細"
)
CHANNEL_SALES_GOOGLE_SHEET_TAB = os.environ.get(
    "CHANNEL_SALES_GOOGLE_SHEET_TAB", "通路銷售明細"
)


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
    return get_mock_data(), [], [], {}, {}, {}

  headers = {
      "Content-Type": "application/json",
      "Authorization": token,
  }

  warehouses = fetch_warehouses(token)
  categories_map = fetch_categories(token)
  items_map = fetch_items_map(token)
  customers_map = fetch_customers(token)
  suppliers_map = fetch_suppliers(token)

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
    return (
        get_mock_data(), warehouses, list(categories_map.values()),
        items_map, customers_map, suppliers_map,
    )

  return (
      pd.DataFrame(all_stock_data),
      warehouses,
      list(categories_map.values()),
      items_map,
      customers_map,
      suppliers_map,
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


def fetch_sales_details_range(token, start_date, end_date, endpoint_path, detail_key):
  """GetSales／GetSaleReturns 共用的抓取邏輯：手冊限制每次查詢區間最長
  7 天、每頁 50 筆，這裡處理單一區間（呼叫端要自己切 7 天視窗）內的
  完整分頁抓取，回傳所有單頭資料（每筆單頭底下的 detail_key 陣列
  就是明細，例如 SaleDetails／SaleReturnDetails）。
  """
  url = f"{A1_BASE_URL}{endpoint_path}"
  headers = {"Content-Type": "application/json", "Authorization": token}
  all_data = []
  pagination = 1
  more = True
  MAX_PAGES = 200  # 安全上限，避免 More 一直是 true 造成無窮迴圈
  while more and pagination <= MAX_PAGES:
    payload = {
        "StartDate": start_date.strftime("%Y-%m-%d"),
        "EndDate": end_date.strftime("%Y-%m-%d"),
        "Pagination": pagination,
    }
    try:
      response = requests.post(
          url, json=payload, headers=headers, timeout=REQUEST_TIMEOUT
      )
      if response.status_code != 200:
        print(f"{endpoint_path} 抓取失敗 [{response.status_code}]: {response.text}")
        break
      res_json = response.json()
      data = res_json.get("Data", []) if isinstance(res_json, dict) else []
      all_data.extend(data)
      more = res_json.get("More", False) if isinstance(res_json, dict) else False
      pagination += 1
    except requests.exceptions.RequestException as e:
      print(f"{endpoint_path} 請求異常: {e}")
      break
  return all_data


def _aggregate_sales_history_range(token, start_date, end_date):
  """核心彙整邏輯：抓指定日期區間（會自動切成 7 天視窗）的銷貨/銷退，
  換算淨銷量。回傳 (qty_by_month_item, name_by_item)，方便呼叫端把多段
  不連續的區間（例如「近3個月」+「去年同月」）合併，不用整段硬抓。
  """
  qty_by_month_item = defaultdict(float)
  name_by_item = {}

  window_start = start_date
  while window_start <= end_date:
    window_end = min(window_start + timedelta(days=6), end_date)

    for sale in fetch_sales_details_range(
        token, window_start, window_end, "/Sales/PaginationQuery", "SaleDetails"
    ):
      year_month = str(sale.get("TradeDate", ""))[:7].replace("/", "-")
      for detail in sale.get("SaleDetails", []) or []:
        item_id = detail.get("ItemDetailID")
        if not item_id:
          continue
        try:
          qty = float(detail.get("Qty") or 0)
        except (TypeError, ValueError):
          qty = 0.0
        qty_by_month_item[(year_month, item_id)] += qty
        if detail.get("ItemName"):
          name_by_item[item_id] = detail["ItemName"]

    for ret in fetch_sales_details_range(
        token, window_start, window_end, "/SaleReturns/PaginationQuery",
        "SaleReturnDetails",
    ):
      year_month = str(ret.get("TradeDate", ""))[:7].replace("/", "-")
      for detail in ret.get("SaleReturnDetails", []) or []:
        item_id = detail.get("ItemDetailID")
        if not item_id:
          continue
        try:
          qty = float(detail.get("Qty") or 0)
        except (TypeError, ValueError):
          qty = 0.0
        qty_by_month_item[(year_month, item_id)] -= qty  # 淨銷量扣掉銷退
        if detail.get("ItemName"):
          name_by_item[item_id] = detail["ItemName"]

    window_start = window_end + timedelta(days=1)

  return qty_by_month_item, name_by_item


def _sales_history_rows_from_aggregate(qty_by_month_item, name_by_item):
  return [
      {
          "年月": year_month,
          "品號": item_id,
          "品名": name_by_item.get(item_id, ""),
          "銷售數量": round(qty, 2),
      }
      for (year_month, item_id), qty in qty_by_month_item.items()
  ]


def fetch_sales_history_from_a1(token, months_back=3):
  """用手冊記載的正式端點 GetSales + GetSaleReturns，直接向 A1 抓近
  N 個月的銷貨／銷退明細，換算成淨銷量（銷貨 − 銷退），彙整成跟
  Google Sheet「銷售歷史」分頁一樣的格式：
  [{"年月","品號","品名","銷售數量"}, ...]

  手冊限制查詢區間最長 7 天，所以要切成很多個 7 天視窗逐一呼叫；抓近
  3 個月大約要跑 13 個視窗 x 2 個端點 ≈ 26 次主要請求（單量大的話還會
  多幾次分頁），會需要幾秒到十幾秒，所以刻意不放進「同步 A1 最新庫存」
  或程式啟動流程裡，避免拖慢一般操作——改成 5.3 頁籤裡一個獨立的手動
  按鈕，需要的時候自己按。
  """
  end_date = datetime.now().date()
  start_date = end_date - timedelta(days=months_back * 30)
  qty, names = _aggregate_sales_history_range(token, start_date, end_date)
  return _sales_history_rows_from_aggregate(qty, names)


def fetch_sales_history_for_forecast(token, target_year_month):
  """5.5 專用：只抓「近3個月」＋「去年同月」這兩段（都是 5.5 算法會用到
  的），不是整年硬抓，可以省下大半的 API 呼叫次數跟等待時間——如果抓
  整年大概要 50+ 個視窗，這樣做大概只要近3個月(~13個視窗) + 去年那個
  月(~5個視窗) ≈ 18 個視窗 x 2 端點，明顯快很多。
  """
  import calendar

  target_year, target_month_num = (int(x) for x in target_year_month.split("-"))
  today = datetime.now().date()

  recent_start = today - timedelta(days=95)
  recent_end = today

  last_year = target_year - 1
  last_day = calendar.monthrange(last_year, target_month_num)[1]
  last_year_start = datetime(last_year, target_month_num, 1).date()
  last_year_end = datetime(last_year, target_month_num, last_day).date()

  qty_recent, names_recent = _aggregate_sales_history_range(
      token, recent_start, recent_end
  )
  qty_last_year, names_last_year = _aggregate_sales_history_range(
      token, last_year_start, last_year_end
  )

  combined_qty = defaultdict(float)
  for d in (qty_recent, qty_last_year):
    for k, v in d.items():
      combined_qty[k] += v
  combined_names = {**names_recent, **names_last_year}

  return _sales_history_rows_from_aggregate(combined_qty, combined_names)


def fetch_customers(token):
  """Customers[Get]（不傳客戶代號）：取得所有客戶列表（ID/Name/ShortName）

  目前系統還沒有畫面直接用到這個，先做成基礎功能備用——之後如果
  「訂單資訊」Sheet 要加客戶欄位，就能直接對照客戶全名，不用另外維護
  一份客戶對照表。
  """
  url = f"{A1_BASE_URL}/Customers"
  headers = {"Authorization": token}
  try:
    response = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
    if response.status_code == 200:
      return {c["ID"]: c.get("Name") or c.get("ShortName") for c in response.json()}
    print(f"取得客戶列表失敗 [{response.status_code}]: {response.text}")
  except requests.exceptions.RequestException as e:
    print(f"取得客戶列表失敗: {e}")
  return {}


def fetch_suppliers(token):
  """Suppliers[Get]（不傳廠商代號）：取得所有廠商列表（ID/Name/ShortName）

  同上，先做成基礎功能備用，之後如果 BOM 表要把「主要供應商」從自由
  文字改成對照 A1 廠商代號，就能直接用這份資料驗證。
  """
  url = f"{A1_BASE_URL}/Suppliers"
  headers = {"Authorization": token}
  try:
    response = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
    if response.status_code == 200:
      return {s["ID"]: s.get("Name") or s.get("ShortName") for s in response.json()}
    print(f"取得廠商列表失敗 [{response.status_code}]: {response.text}")
  except requests.exceptions.RequestException as e:
    print(f"取得廠商列表失敗: {e}")
  return {}


# -------------------------------------------------------------------------
# Orders[Post] 反向同步：把 Google Sheets「訂單資訊」的內容寫回 A1
# 這是「寫入」正式 A1 系統的動作，刻意做成手動觸發＋確認彈窗，不會自動
# 執行。手冊要求 CustomerID（客戶代號）、TotalSaleAmount（總金額）、
# 每個品項的 Amount 都必填，所以「訂單資訊」Sheet 新增了「客戶代號」跟
# 「金額」兩欄——只有要同步到 A1 才需要填，純粹讀取分析（缺貨預警等）
# 不需要這兩欄也能正常運作。
# -------------------------------------------------------------------------
def build_order_upload_payloads(orders):
  """把 orders（Sheet 讀到的列，每列＝一個品項）依「訂單編號」分組，組成
  Orders[Post] 需要的完整格式（一張訂單可以有多個品項）。沒填訂單編號的
  列，各自當成獨立的單品項訂單處理。

  回傳 (payloads, skipped)：
    payloads: list[(order_id, payload_dict)]，可以直接上傳的訂單
    skipped: list[str]，因缺客戶代號或金額而無法組成訂單的說明文字
  """
  grouped = defaultdict(list)
  for idx, o in enumerate(orders):
    key = o.get("訂單編號") or f"__single__{o['品號']}_{o['預計出貨日'].isoformat()}_{idx}"
    grouped[key].append(o)

  payloads = []
  skipped = []
  today_str = datetime.now().date().strftime("%Y/%m/%d")

  for order_id, lines in grouped.items():
    customer_id = next(
        (l["客戶代號"] for l in lines if l.get("客戶代號")), ""
    )
    if not customer_id:
      skipped.append(f"{order_id}：缺「客戶代號」，無法上傳")
      continue

    missing_amount = [l for l in lines if l.get("金額") is None]
    if missing_amount:
      skipped.append(f"{order_id}：有品項缺「金額」，無法上傳")
      continue

    pre_delivery_date = max(l["預計出貨日"] for l in lines)
    details = []
    total_amount = 0.0
    for i, l in enumerate(lines, start=1):
      amount = l["金額"]
      total_amount += amount
      details.append({
          "ID": i,
          "ItemID": l["品號"],
          "Qty": l["預計出貨數量"],
          "Amount": amount,
          "PreDeliveryDate": l["預計出貨日"].strftime("%Y/%m/%d"),
      })

    display_id = order_id if not order_id.startswith("__single__") else lines[0]["品號"]
    payload = {
        "ID": order_id,
        "TradeDate": today_str,  # Sheet 沒有「下單日」欄位，用上傳當天當交易日期
        "CustomerID": customer_id,
        "TotalSaleAmount": round(total_amount, 2),
        "PreDeliveryDate": pre_delivery_date.strftime("%Y/%m/%d"),
        "Details": details,
    }
    payloads.append((display_id, payload))

  return payloads, skipped


def upload_order_to_a1(token, payload):
  """Orders[Post]：上傳單張訂單。回傳 (成功與否, 訊息)。

  409（唯一辨識碼重複）視為「這張訂單先前已經上傳過」，不當成錯誤，讓
  呼叫端可以正常統計、不用擔心重複按會出亂子。
  """
  url = f"{A1_BASE_URL}/Orders"
  headers = {"Content-Type": "application/json", "Authorization": token}
  try:
    response = requests.post(
        url, json=payload, headers=headers, timeout=REQUEST_TIMEOUT
    )
    if response.status_code == 200:
      return True, "上傳成功"
    if response.status_code == 409:
      return False, "訂單編號重複（先前應該已經上傳過，正常現象）"
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


def _fetch_google_sheet_records(tab_name):
  """共用的 Google Sheet 讀取邏輯（BOM／訂單資訊／銷售歷史都靠這支）。

  回傳 None 代表「沒有設定 Google Sheets」（呼叫端應退回其他備援來源）；
  回傳空 list 代表「有設定，但讀取失敗，或該分頁本來就沒有資料」。
  這兩種情況要分開，才不會把「沒設定」誤判成「設定了但是空的」。

  設定步驟：
  1. Google Cloud Console 建立服務帳號，下載 JSON 金鑰
  2. 把金鑰 JSON 的完整內容存進環境變數 GOOGLE_SHEETS_CREDENTIALS_JSON
  3. 把該服務帳號的 email（金鑰 JSON 裡的 client_email）加入 Google
     Sheet 的「共用」名單，權限「檢視者」即可
  4. 設定 GOOGLE_SHEET_ID（Sheet 網址中 /d/ 與 /edit 中間那串）
  5. 分頁名稱要跟 BOM_GOOGLE_SHEET_TAB／ORDERS_GOOGLE_SHEET_TAB／
     SALES_HISTORY_GOOGLE_SHEET_TAB 一致，欄位標題要跟本檔案裡的
     BOM_COL_* / ORDER_COL_* / SALES_COL_* 常數完全一致
  6. requirements.txt 需要加上 gspread、google-auth
  """
  if not GOOGLE_SHEETS_CREDENTIALS_JSON or not GOOGLE_SHEET_ID:
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
    sh = gc.open_by_key(GOOGLE_SHEET_ID)
    ws = sh.worksheet(tab_name)
    return ws.get_all_records()
  except Exception as e:
    print(f"讀取 Google Sheet「{tab_name}」分頁失敗：{e}")
    return []


def load_bom_from_google_sheet():
  """讀取 Google Sheet 版的「BOM表」分頁；None＝未設定，[]/dict＝已設定"""
  records = _fetch_google_sheet_records(BOM_GOOGLE_SHEET_TAB)
  if records is None:
    return None
  bom_map = _parse_bom_records(records)
  print(
      f"Google Sheet BOM表讀取完成：共 {len(bom_map)} 個主件品號、"
      f"{sum(len(v) for v in bom_map.values())} 筆子件關係"
  )
  return bom_map


def load_bom_data():
  """統一入口：優先讀 Google Sheets，沒設定時自動退回本機 Excel（過渡期
  相容），回傳 (bom_map, 資料來源標籤)。
  """
  sheet_result = load_bom_from_google_sheet()
  if sheet_result is not None:
    return sheet_result, "Google Sheets"
  return load_bom_from_excel(BOM_EXCEL_PATH), "本機 Excel（尚未設定 Google Sheets）"


# ---- 訂單資訊（Google Sheet，手動覆蓋更新／後續可改自動抓取） ----
# 因 A1 API 只能上傳訂單、無法查詢訂單（手冊只有 Orders[Post]，沒有對應
# 的取得端點），出貨排程/缺貨預警需要的「未來訂單」資料改成人工／其他
# 系統匯出後貼進這份 Google Sheet，程式只負責讀取跟計算。
ORDER_COL_NO = "訂單編號（選填）"
ORDER_COL_ITEM_ID = "品號"
ORDER_COL_ITEM_NAME = "品名（選填，供參考）"
ORDER_COL_DUE_DATE = "預計出貨日"
ORDER_COL_QTY = "預計出貨數量"
ORDER_COL_STATUS = "狀態（選填：未出貨/備貨中/已出貨）"
ORDER_COL_CUSTOMER_ID = "客戶代號（同步到A1才需要）"
ORDER_COL_AMOUNT = "金額（同步到A1才需要）"
ORDER_COL_MEMO = "備註"


def _parse_flexible_date(raw):
  """把 Google Sheet 裡各種日期寫法（2026/8/10、2026-08-10...）轉成
  Python date；轉不出來就回傳 None，呼叫端會跳過該列，不會讓整批資料
  因為一列日期寫錯就整個讀取失敗。
  """
  if raw in (None, ""):
    return None
  try:
    ts = pd.to_datetime(raw)
    if pd.isna(ts):
      return None
    return ts.date()
  except (ValueError, TypeError):
    return None


def _parse_order_records(records):
  orders = []
  for row in records:
    item_id = str(row.get(ORDER_COL_ITEM_ID, "") or "").strip()
    due_date = _parse_flexible_date(row.get(ORDER_COL_DUE_DATE))
    if not item_id or due_date is None:
      continue  # 沒填品號或日期格式看不懂的列，直接跳過，不擋整批匯入

    qty_raw = str(row.get(ORDER_COL_QTY, "") or "").strip()
    try:
      qty = float(qty_raw) if qty_raw else 0.0
    except ValueError:
      qty = 0.0

    status = str(row.get(ORDER_COL_STATUS, "") or "").strip() or "未出貨"

    amount_raw = str(row.get(ORDER_COL_AMOUNT, "") or "").strip()
    try:
      amount = float(amount_raw) if amount_raw else None
    except ValueError:
      amount = None

    orders.append({
        "訂單編號": str(row.get(ORDER_COL_NO, "") or "").strip(),
        "品號": item_id,
        "品名": str(row.get(ORDER_COL_ITEM_NAME, "") or "").strip(),
        "預計出貨日": due_date,
        "預計出貨數量": qty,
        "狀態": status,
        "客戶代號": str(row.get(ORDER_COL_CUSTOMER_ID, "") or "").strip(),
        "金額": amount,
        "備註": str(row.get(ORDER_COL_MEMO, "") or "").strip(),
    })
  return orders


def load_orders_from_google_sheet():
  """讀取「訂單資訊」分頁。回傳 (orders, configured)。

  configured=False 代表 Google Sheets 根本沒設定；這種情況下畫面要顯示
  「請先設定 Google Sheets」而不是「目前沒有訂單」，兩種情況給使用者的
  訊息應該不一樣。
  """
  records = _fetch_google_sheet_records(ORDERS_GOOGLE_SHEET_TAB)
  if records is None:
    return [], False
  orders = _parse_order_records(records)
  print(f"訂單資訊讀取完成：共 {len(orders)} 筆有效訂單列")
  return orders, True


# ---- 銷售歷史（Google Sheet，用於 5.3 庫存週轉率/滯銷品分析） ----
SALES_COL_YM = "年月"
SALES_COL_ITEM_ID = "品號"
SALES_COL_ITEM_NAME = "品名（選填，供參考）"
SALES_COL_QTY = "銷售數量"


def _parse_sales_history_records(records):
  rows = []
  for row in records:
    item_id = str(row.get(SALES_COL_ITEM_ID, "") or "").strip()
    year_month = str(row.get(SALES_COL_YM, "") or "").strip()
    if not item_id or not year_month:
      continue
    qty_raw = str(row.get(SALES_COL_QTY, "") or "").strip()
    try:
      qty = float(qty_raw) if qty_raw else 0.0
    except ValueError:
      qty = 0.0
    rows.append({
        "年月": year_month,
        "品號": item_id,
        "品名": str(row.get(SALES_COL_ITEM_NAME, "") or "").strip(),
        "銷售數量": qty,
    })
  return rows


def load_sales_history_from_google_sheet():
  """讀取「銷售歷史」分頁。回傳 (sales_rows, configured)，意義同上。"""
  records = _fetch_google_sheet_records(SALES_HISTORY_GOOGLE_SHEET_TAB)
  if records is None:
    return [], False
  rows = _parse_sales_history_records(records)
  print(f"銷售歷史讀取完成：共 {len(rows)} 筆有效紀錄")
  return rows, True


# ---- 進貨明細（Google Sheet；A1 只有 Receives[Post] 上傳、沒有查詢端點，
#      所以「進貨明細」跟訂單資訊/銷售歷史一樣改用 Sheet 維護。詳見本次
#      回覆中的 A1 報表匯出可行性說明） ----
RECEIVING_COL_DATE = "進貨日期"
RECEIVING_COL_ITEM_ID = "品號"
RECEIVING_COL_ITEM_NAME = "品名（選填，供參考）"
RECEIVING_COL_QTY = "進貨數量"
RECEIVING_COL_UNIT_PRICE = "單價（選填）"
RECEIVING_COL_SUPPLIER = "供應商（選填）"
RECEIVING_COL_MEMO = "備註"


def _parse_receiving_records(records):
  rows = []
  for row in records:
    item_id = str(row.get(RECEIVING_COL_ITEM_ID, "") or "").strip()
    receiving_date = _parse_flexible_date(row.get(RECEIVING_COL_DATE))
    if not item_id or receiving_date is None:
      continue

    qty_raw = str(row.get(RECEIVING_COL_QTY, "") or "").strip()
    try:
      qty = float(qty_raw) if qty_raw else 0.0
    except ValueError:
      qty = 0.0

    price_raw = str(row.get(RECEIVING_COL_UNIT_PRICE, "") or "").strip()
    try:
      unit_price = float(price_raw) if price_raw else None
    except ValueError:
      unit_price = None

    rows.append({
        "進貨日期": receiving_date,
        "品號": item_id,
        "品名": str(row.get(RECEIVING_COL_ITEM_NAME, "") or "").strip(),
        "進貨數量": qty,
        "單價": unit_price,
        "供應商": str(row.get(RECEIVING_COL_SUPPLIER, "") or "").strip(),
        "備註": str(row.get(RECEIVING_COL_MEMO, "") or "").strip(),
    })
  return rows


def load_receivings_from_google_sheet():
  """讀取「進貨明細」分頁。回傳 (receivings, configured)，意義同訂單資訊。"""
  records = _fetch_google_sheet_records(RECEIVING_GOOGLE_SHEET_TAB)
  if records is None:
    return [], False
  rows = _parse_receiving_records(records)
  print(f"進貨明細讀取完成：共 {len(rows)} 筆有效紀錄")
  return rows, True


# ---- 通路銷售明細（Google Sheet；供未來「通路別分析」使用） ----
# 這份跟「銷售歷史」不同：銷售歷史是單純的「年月/品號/數量」彙總，
# 給 5.3 週轉率、5.5 月產銷分析用；這份多了通路分類、客戶、成本，是
# 更細的原始明細，供之後要做「哪個通路賺最多」「哪個客戶貢獻最大」這類
# 分析時使用。目前系統還沒有畫面直接呈現這份資料，先把讀取功能建好。
CHANNEL_COL_YM = "年月"
CHANNEL_COL_CATEGORY = "通路分類（官網/蝦皮/門市/經銷(團購)/KOL）"
CHANNEL_COL_CUSTOMER = "客戶（選填）"
CHANNEL_COL_ITEM_ID = "品號（選填，供對照）"
CHANNEL_COL_ITEM_NAME = "品名"
CHANNEL_COL_QTY = "數量"
CHANNEL_COL_COST = "成本"


def _parse_channel_sales_records(records):
  rows = []
  for row in records:
    item_name = str(row.get(CHANNEL_COL_ITEM_NAME, "") or "").strip()
    year_month = str(row.get(CHANNEL_COL_YM, "") or "").strip()
    if not item_name or not year_month:
      continue

    qty_raw = str(row.get(CHANNEL_COL_QTY, "") or "").strip()
    try:
      qty = float(qty_raw) if qty_raw else 0.0
    except ValueError:
      qty = 0.0

    cost_raw = str(row.get(CHANNEL_COL_COST, "") or "").strip()
    try:
      cost = float(cost_raw) if cost_raw else 0.0
    except ValueError:
      cost = 0.0

    rows.append({
        "年月": year_month,
        "通路分類": str(row.get(CHANNEL_COL_CATEGORY, "") or "").strip(),
        "客戶": str(row.get(CHANNEL_COL_CUSTOMER, "") or "").strip(),
        "品號": str(row.get(CHANNEL_COL_ITEM_ID, "") or "").strip(),
        "品名": item_name,
        "數量": qty,
        "成本": cost,
    })
  return rows


def load_channel_sales_from_google_sheet():
  """讀取「通路銷售明細」分頁。回傳 (rows, configured)。"""
  records = _fetch_google_sheet_records(CHANNEL_SALES_GOOGLE_SHEET_TAB)
  if records is None:
    return [], False
  rows = _parse_channel_sales_records(records)
  print(f"通路銷售明細讀取完成：共 {len(rows)} 筆有效紀錄")
  return rows, True


# -------------------------------------------------------------------------
# 興聖集團旗下分公司清單（右上角切換用）
# 目前「海濤客食品工業(股)公司」已完成完整 A1 API 串接（render_hai_tao_ke_page）。
# 「興聖(股)公司」「容鴻(股)公司」「芙萊柏(股)公司」三間已建好共用的頁面骨架
# （render_channel_company_page：儀表板／訂單出貨4通路／每日出貨／調撥紀錄／
# 退換貨記錄），但各區塊資料都還沒串接，畫面上顯示佔位提示。之後陸續拿到
# 各分公司/各通路的 API 時，把 render_section_placeholder(...) 換成真的資料
# 表格即可，不用重搭分頁結構。
# -------------------------------------------------------------------------
# -------------------------------------------------------------------------
# 顏色標籤：儀表板／提醒中心統一用這組顏色分辨「嚴重程度」
# danger=紅（已逾期/缺口大，立即處理）、warning=黃（提前準備）、
# info=藍綠（一般提醒）、success=綠（狀況良好）
# -------------------------------------------------------------------------
SEVERITY_STYLES = {
    "danger": {
        "box": "bg-[#fdecea] border-[#f5c2c0]",
        "text": "text-red-700",
        "badge": "bg-red-700 text-white",
        "label": "緊急",
    },
    "warning": {
        "box": "bg-[#fff8e6] border-[#f0dca0]",
        "text": "text-amber-800",
        "badge": "bg-amber-500 text-white",
        "label": "注意",
    },
    "info": {
        "box": "bg-[#e8f6f5] border-[#bfe6e3]",
        "text": "text-teal-800",
        "badge": "bg-teal-600 text-white",
        "label": "提醒",
    },
    "success": {
        "box": "bg-[#eaf6ec] border-[#bfe3c5]",
        "text": "text-green-800",
        "badge": "bg-green-700 text-white",
        "label": "正常",
    },
}


def compute_order_demand_alerts(orders, items_map, bom_map, stock_lookup, settings, horizon_days=30):
  """核心運算：把「訂單資訊」Google Sheet 的未來出貨需求，展開 BOM 子件，
  跟目前庫存比對，算出：
    1. 未來 horizon_days 天內要出貨、但成品庫存不夠的品項
    2. 因此連帶需要補的原料/子件缺口，以及建議下單日
       （建議下單日 = 最早相關出貨日 － 採購前置天數 － 生產工時）
  這支函式同時餵給「儀表板」「訂單出貨」「生產排程」，
  確保三個頁面看到的是同一套邏輯算出來的數字，不會互相矛盾。

  簡化假設（先講清楚，之後有更多資料再精進）：
  - 沒有追蹤「在途採購量」，缺口＝需求－現有庫存，沒有扣掉已下單未到貨的量
  - 建議下單日抓「最早相關出貨日」回推，同一原料被多張訂單用到時，用最早
    那張抓緊
  """
  today = datetime.now().date()
  horizon_end = today + timedelta(days=horizon_days)
  default_lead_time = settings.get("default_lead_time_days", 7)

  orders_in_horizon = [
      o
      for o in orders
      if o["狀態"] != "已出貨" and today <= o["預計出貨日"] <= horizon_end
  ]

  demand_by_item = defaultdict(float)
  earliest_due_by_item = {}
  for o in orders_in_horizon:
    demand_by_item[o["品號"]] += o["預計出貨數量"]
    if o["品號"] not in earliest_due_by_item or o["預計出貨日"] < earliest_due_by_item[o["品號"]]:
      earliest_due_by_item[o["品號"]] = o["預計出貨日"]

  finished_goods_shortfall = []
  raw_material_demand = defaultdict(float)
  raw_material_earliest_due = {}
  raw_material_work_days = defaultdict(float)

  for item_id, demand_qty in demand_by_item.items():
    current_stock = stock_lookup.get(item_id, 0.0)
    shortage = demand_qty - current_stock
    info = items_map.get(item_id, {})
    if shortage <= 0:
      continue

    finished_goods_shortfall.append({
        "品號": item_id,
        "品名": info.get("Name"),
        "未來需求量": demand_qty,
        "現有庫存": current_stock,
        "缺口": round(shortage, 2),
        "最早出貨日": earliest_due_by_item.get(item_id, today).isoformat(),
    })

    for comp in bom_map.get(item_id, []):
      loss_rate = comp.get("損耗率") or 0
      try:
        loss_rate = float(loss_rate)
      except (TypeError, ValueError):
        loss_rate = 0.0
      try:
        unit_qty = float(comp.get("用量") or 0)
      except (TypeError, ValueError):
        unit_qty = 0.0
      multiplier = 1 + (loss_rate / 100 if loss_rate else 0)
      child_id = comp["子件品號"]
      raw_material_demand[child_id] += shortage * unit_qty * multiplier

      due = earliest_due_by_item.get(item_id, today)
      if child_id not in raw_material_earliest_due or due < raw_material_earliest_due[child_id]:
        raw_material_earliest_due[child_id] = due

      work_days = comp.get("生產工時天數") or 0
      try:
        work_days = float(work_days)
      except (TypeError, ValueError):
        work_days = 0.0
      raw_material_work_days[child_id] = max(raw_material_work_days[child_id], work_days)

  raw_material_shortfall = []
  for child_id, need_qty in raw_material_demand.items():
    current_stock = stock_lookup.get(child_id, 0.0)
    gap = need_qty - current_stock
    if gap <= 0:
      continue
    info = items_map.get(child_id, {})

    lead_time = None
    for components in bom_map.values():
      for comp in components:
        if comp["子件品號"] == child_id:
          lt = comp.get("採購前置天數")
          if isinstance(lt, (int, float)) and lt > 0:
            lead_time = lt
      if lead_time is not None:
        break
    lead_time = lead_time if lead_time is not None else default_lead_time

    due_date = raw_material_earliest_due.get(child_id, today)
    work_days = raw_material_work_days.get(child_id, 0)
    suggested_order_date = due_date - timedelta(days=lead_time + work_days)

    days_until_order = (suggested_order_date - today).days
    if days_until_order < 0:
      severity = "danger"
    elif days_until_order <= 3:
      severity = "warning"
    else:
      severity = "info"

    raw_material_shortfall.append({
        "品號": child_id,
        "品名": info.get("Name"),
        "未來需求量(含損耗)": round(need_qty, 2),
        "現有庫存": current_stock,
        "缺口": round(gap, 2),
        "採購前置天數": lead_time,
        "建議下單日": suggested_order_date.isoformat(),
        "severity": severity,
    })

  raw_material_shortfall.sort(key=lambda r: r["建議下單日"])
  finished_goods_shortfall.sort(key=lambda r: r["缺口"], reverse=True)

  return {
      "orders_in_horizon": orders_in_horizon,
      "finished_goods_shortfall": finished_goods_shortfall,
      "raw_material_shortfall": raw_material_shortfall,
  }


def compute_dashboard_announcements(orders, items_map, bom_map, stock_lookup, settings, horizon_days=30):
  """把 compute_order_demand_alerts() 的結果，整理成 4 種提醒公告：
    - shipping　　訂單出貨提醒：未來要出貨的訂單
    - production　生產組裝確認：成品庫存不夠，但原料已經備妥，可以安排組裝/生產
    - procurement　採購提醒：原料/子件不夠，需要下單
    - incoming　　進貨提醒：依建議下單日＋前置天數推算的預計到貨日快到了，
                  提醒去確認廠商到貨狀況（這是推算出來的日期，不是真的追蹤
                  在途訂單，之後有進貨單資料可以取代這個推算）

  每種都回傳 list[{"text":..., "severity":...}]，供儀表板或個別頁面頂端
  的提醒橫幅使用；儀表板顯示全部 4 種，頁面各自只顯示跟自己相關的那種，
  但共用同一套運算結果，數字不會兜不起來。
  """
  today = datetime.now().date()
  result = compute_order_demand_alerts(
      orders, items_map, bom_map, stock_lookup, settings, horizon_days
  )

  shipping = []
  for o in sorted(result["orders_in_horizon"], key=lambda x: x["預計出貨日"]):
    days_left = (o["預計出貨日"] - today).days
    severity = "danger" if days_left <= 1 else "warning" if days_left <= 3 else "info"
    order_label = o["訂單編號"] or o["品號"]
    shipping.append({
        "text": (
            f"訂單 {order_label}（{o.get('品名') or o['品號']}）需於"
            f" {o['預計出貨日'].isoformat()} 出貨（數量 {o['預計出貨數量']:g}）"
        ),
        "severity": severity,
    })

  # 生產組裝確認：成品有缺口，但它用到的原料都不在「原料缺口清單」裡，
  # 代表原料已經夠了，可以排組裝/生產把成品補齊
  raw_shortfall_ids = {r["品號"] for r in result["raw_material_shortfall"]}
  production = []
  for r in result["finished_goods_shortfall"]:
    item_id = r["品號"]
    components = bom_map.get(item_id, [])
    blocked = any(c["子件品號"] in raw_shortfall_ids for c in components)
    if blocked:
      continue
    try:
      due_date = datetime.fromisoformat(r["最早出貨日"]).date()
      days_left = (due_date - today).days
    except (ValueError, TypeError):
      days_left = None
    severity = "danger" if (days_left is not None and days_left <= 2) else "warning"
    production.append({
        "text": (
            f"生產組裝確認：{r.get('品名') or item_id} 缺口 {r['缺口']:g}，"
            f"原料已備妥，建議安排組裝／生產（最早出貨日 {r['最早出貨日']}）"
        ),
        "severity": severity,
    })

  # 建議採購成品（只看母件）：未來需求 > 現有庫存的成品，不論原料夠不夠，
  # 都先讓管理者看到「這個成品接下來會不夠」，可以決定要生產還是外購
  finished_goods = []
  for r in result["finished_goods_shortfall"]:
    try:
      due_date = datetime.fromisoformat(r["最早出貨日"]).date()
      days_left = (due_date - today).days
    except (ValueError, TypeError):
      days_left = None
    severity = (
        "danger" if (days_left is not None and days_left <= 3)
        else "warning"
    )
    finished_goods.append({
        "text": (
            f"建議採購成品：{r.get('品名') or r['品號']} 未來需求"
            f" {r['未來需求量']:g}，現有庫存 {r['現有庫存']:g}，"
            f"缺口 {r['缺口']:g}（最早出貨日 {r['最早出貨日']}）"
        ),
        "severity": severity,
    })

  procurement = []
  incoming = []
  for m in result["raw_material_shortfall"]:
    procurement.append({
        "text": (
            f"採購提醒：{m.get('品名') or m['品號']}（缺口 {m['缺口']:g}），"
            f"建議下單日 {m['建議下單日']}"
        ),
        "severity": m["severity"],
    })
    try:
      order_date = datetime.fromisoformat(m["建議下單日"]).date()
      lead_time = m.get("採購前置天數") or 0
      arrival_date = order_date + timedelta(days=lead_time)
      days_to_arrival = (arrival_date - today).days
      if days_to_arrival <= 5:
        arrival_severity = (
            "danger" if days_to_arrival < 0
            else "warning" if days_to_arrival <= 2
            else "info"
        )
        incoming.append({
            "text": (
                f"進貨提醒：{m.get('品名') or m['品號']} 依前置天數推算"
                f"預計到貨日 {arrival_date.isoformat()}，請確認廠商到貨狀況"
            ),
            "severity": arrival_severity,
        })
    except (ValueError, TypeError):
      pass

  return {
      "shipping": shipping,
      "production": production,
      "finished_goods": finished_goods,
      "procurement": procurement,
      "incoming": incoming,
  }


def compute_turnover_metrics(sales_history, stock_lookup, items_map, slow_moving_days=90):
  """5.3 庫存週轉率／滯銷品分析：用「銷售歷史」Google Sheet 抓每個品號最近
  3 個月的月銷量，算出週轉天數 = 現有庫存 ÷ 日均銷量。週轉天數異常長（或
  完全沒賣出過、但還有庫存）就標記為滯銷。
  """
  by_item = defaultdict(list)
  for row in sales_history:
    by_item[row["品號"]].append(row)

  results = []
  for item_id, records in by_item.items():
    records_sorted = sorted(records, key=lambda r: r["年月"])
    last3 = records_sorted[-3:]
    avg_monthly = (
        sum(r["銷售數量"] for r in last3) / len(last3) if last3 else 0.0
    )
    current_stock = stock_lookup.get(item_id, 0.0)
    daily_avg = avg_monthly / 30 if avg_monthly else 0.0
    turnover_days = (current_stock / daily_avg) if daily_avg > 0 else None

    info = items_map.get(item_id, {})
    item_name = info.get("Name") or (records[-1]["品名"] if records else "")

    is_slow_moving = current_stock > 0 and (
        turnover_days is None or turnover_days > slow_moving_days
    )

    results.append({
        "品號": item_id,
        "品名": item_name,
        "近3月平均月銷": ceil_qty(avg_monthly),
        "現有庫存": ceil_qty(current_stock),
        "庫存週轉天數": round(turnover_days, 1) if turnover_days is not None else "從未銷售",
        "滯銷": "是" if is_slow_moving else "否",
    })

  results.sort(
      key=lambda r: (r["庫存週轉天數"] if isinstance(r["庫存週轉天數"], (int, float)) else 999999),
      reverse=True,
  )
  return results


def generate_month_options(n=7):
  """產生「從本月開始」的 N 個月份選項（YYYY-MM），供 5.5 月份選單用。
  第一個選項是當月（方便月中還想針對當月剩餘天數抓貨），後面接著未來
  幾個月。
  """
  today = datetime.now().date()
  options = []
  y, m = today.year, today.month
  for _ in range(n):
    options.append(f"{y}-{m:02d}")
    m += 1
    if m > 12:
      m = 1
      y += 1
  return options


def is_finished_or_combo_category(category_name):
  """5.5 只分析「成品」跟「組合品」這兩類——原料/物料/費用/代工含料這些
  不是賣給客戶的品項，不該算進銷售預測。用類別名稱是否包含「成品」或
  「組合品」字樣判斷（對照 A1 目前的分類命名慣例，例如"
  "「(海濤客)_成品11」「(海濤客)_組合品61」）。如果貴公司分類命名方式
  不同（例如用「半成品」也會被誤判成「成品」），需要再調整這裡的規則。
  """
  category_name = category_name or ""
  return "成品" in category_name or "組合品" in category_name


def compute_monthly_production_sales_forecast(
    sales_history, items_map, bom_map, target_year_month, target_revenue, settings,
    last_year_target_revenue=0,
):
  """5.5 月產銷分析：預估目標月份（例如 "2026-08"）的採購量／成本／
  建議採購時間。

  範圍：只分析「成品」跟「組合品」兩個分類的商品（見
  is_finished_or_combo_category），原料/物料/費用類不列入計算——不只是
  畫面上濾掉，是連「近3個月平均營收」「去年比例回推」這些基準計算都只
  看這個範圍，確保分析的樣本跟結果是一致的，不會發生「篩選後看到的」
  跟「實際拿去算比例的」是兩組不同資料的情況。

  參考依據（依需求指定）：
  - 去年同期：目標月份的去年同月（8月 → 去年8月）。如果去年的品號
    編碼跟今年不一樣（常見情況：換過品號規則），會完全比對不到，這時
    改用「去年目標營業額」等比例回推——用「近3個月」的營收佔比當作
    權重，把去年目標營業額依同樣的權重分攤回各品項，再除以單價換算
    回數量。這是估算，不是實際數字，畫面會用「去年銷量來源」欄位
    標明是「實際」還是「依比例推算」。
  - 近3個月平均：目標月份往前推 3 個「完整月」（8月 → 5、6、7月平均）
  - 基準預估銷量 = 兩者平均
  - 若有填「目標營業額」，用「目標營業額 ÷ 基準預估總營收」的比例，
    等比例縮放每個品項的預估銷量，讓由下而上算出來的總營收貼近業務
    設定的目標（由上而下校正），沒填就直接用基準預估量
  - 目標採購量＝校正後的預估銷量（先不額外疊加安全庫存，這是最基本
    版本，之後可以再疊加安全庫存邏輯）
  - 預估總成本＝目標採購量 × 單位成本（StdPurPrice，A1 商品主檔既有
    欄位）
  - 建議採購時間＝目標月份第一天 − 採購前置天數（優先抓 BOM 表裡這個
    品項「作為子件」登記的前置天數，沒有就用系統預設值）
  """
  target_year, target_month_num = (int(x) for x in target_year_month.split("-"))
  last_year_ym = f"{target_year - 1}-{target_month_num:02d}"

  recent_months = []
  y, m = target_year, target_month_num
  for _ in range(3):
    m -= 1
    if m == 0:
      m = 12
      y -= 1
    recent_months.append(f"{y}-{m:02d}")

  qty_last_year = defaultdict(float)
  qty_recent = defaultdict(lambda: defaultdict(float))
  name_by_item = {}

  for row in sales_history:
    ym = row.get("年月")
    item_id = row.get("品號")
    if not item_id:
      continue
    qty = row.get("銷售數量") or 0
    if row.get("品名"):
      name_by_item[item_id] = row["品名"]
    if ym == last_year_ym:
      qty_last_year[item_id] += qty
    if ym in recent_months:
      qty_recent[item_id][ym] += qty

  all_item_ids = set(qty_last_year) | set(qty_recent)
  all_item_ids = {
      iid for iid in all_item_ids
      if is_finished_or_combo_category(items_map.get(iid, {}).get("CategoryName"))
  }

  # 先算好每個品項的「近3月平均」跟「單價」，供「依去年目標營業額比例
  # 回推」使用（權重 = 這個品項近3月營收 ÷ 全部品項近3月營收）
  recent_avg_by_item = {}
  unit_price_by_item = {}
  for item_id in all_item_ids:
    recent_qty_dict = qty_recent.get(item_id, {})
    recent_avg_by_item[item_id] = (
        sum(recent_qty_dict.values()) / 3 if recent_qty_dict else 0.0
    )
    info = items_map.get(item_id, {})
    unit_price_by_item[item_id] = info.get("SalesPrice") or 0

  total_recent_revenue = sum(
      recent_avg_by_item[i] * unit_price_by_item[i] for i in all_item_ids
  )

  lead_time_by_child = {}
  for components in bom_map.values():
    for comp in components:
      lt = comp.get("採購前置天數")
      if isinstance(lt, (int, float)) and lt > 0:
        lead_time_by_child[comp["子件品號"]] = lt
  default_lead_time = settings.get("default_lead_time_days", 7)
  target_month_start = datetime(target_year, target_month_num, 1).date()

  rows = []
  for item_id in all_item_ids:
    last_year_qty_actual = qty_last_year.get(item_id, 0.0)
    recent_avg = recent_avg_by_item[item_id]
    unit_price = unit_price_by_item[item_id]

    last_year_qty = last_year_qty_actual
    last_year_source = "實際"
    if last_year_qty_actual <= 0:
      if last_year_target_revenue and total_recent_revenue > 0 and unit_price > 0:
        revenue_share = (recent_avg * unit_price) / total_recent_revenue
        last_year_qty = (last_year_target_revenue * revenue_share) / unit_price
        last_year_source = "依去年目標營業額比例推算"
      else:
        last_year_source = "無資料"

    baseline_qty = (last_year_qty + recent_avg) / 2
    if baseline_qty <= 0:
      continue

    info = items_map.get(item_id, {})
    item_name = info.get("Name") or name_by_item.get(item_id, "")
    unit_cost = info.get("StdPurPrice") or 0
    lead_time = lead_time_by_child.get(item_id, default_lead_time)
    suggested_order_date = target_month_start - timedelta(days=lead_time)

    rows.append({
        "品號": item_id,
        "品名": item_name,
        "商品分類": info.get("CategoryName") or "未分類",
        "去年同期銷量": ceil_qty(last_year_qty),
        "去年銷量來源": last_year_source,
        "近3月平均銷量": ceil_qty(recent_avg),
        "基準預估銷量": round(baseline_qty, 1),  # 內部用，未四捨五入避免縮放誤差累積
        "單價": unit_price,
        "單位成本": unit_cost,
        "建議採購時間": suggested_order_date.isoformat(),
    })

  baseline_total_revenue = sum(r["基準預估銷量"] * r["單價"] for r in rows)
  if target_revenue and baseline_total_revenue > 0:
    scale_factor = target_revenue / baseline_total_revenue
  else:
    scale_factor = 1.0

  for r in rows:
    est_qty = ceil_qty(r["基準預估銷量"] * scale_factor)
    r["目標採購量"] = est_qty
    r["預估總成本"] = round(est_qty * r["單位成本"], 0)
    del r["基準預估銷量"]  # 只是計算用的中間值，不用顯示給使用者

  rows.sort(key=lambda r: r["預估總成本"], reverse=True)

  return {
      "rows": rows,
      "scale_factor": round(scale_factor, 3),
      "total_est_qty": sum(r["目標採購量"] for r in rows),
      "total_est_cost": sum(r["預估總成本"] for r in rows),
      "total_est_revenue": sum(r["目標採購量"] * r["單價"] for r in rows),
      "earliest_order_date": min((r["建議採購時間"] for r in rows), default=None),
      "last_year_ym": last_year_ym,
      "recent_months": recent_months,
  }


PRODUCTION_SALES_CHANNELS = [
    "官網", "蝦皮", "經銷", "團購", "門市", "KOL", "快閃", "其它通路",
]


def compute_channel_breakdown(forecast_result, channel_percentages, target_revenue):
  """5.5 通路占比拆分：不對每個品項拆通路（現有資料沒有細到「這個品項
  在這個通路賣多少」），而是把整體預估結果，依人工填的占比（例如
  官網50%、蝦皮25%...）等比例分攤，簡單、好懂、好維護——這正是使用者
  要的「簡化」版本，不是更複雜的逐品項通路分析。

  channel_percentages: {通路名稱: 佔比(0-100)}
  回傳 (rows, total_pct)
  """
  rows = []
  total_pct = sum(v for v in channel_percentages.values() if v)
  for channel in PRODUCTION_SALES_CHANNELS:
    pct = channel_percentages.get(channel) or 0
    if pct <= 0:
      continue
    ratio = pct / 100
    rows.append({
        "通路": channel,
        "佔比(%)": pct,
        "目標營業額": round(target_revenue * ratio, 0) if target_revenue else None,
        "目標採購量": ceil_qty(forecast_result["total_est_qty"] * ratio),
        "預估總成本": round(forecast_result["total_est_cost"] * ratio, 0),
        "預估總營收": round(forecast_result["total_est_revenue"] * ratio, 0),
    })
  return rows, total_pct


COMPANIES = ["興聖(股)公司", "海濤客食品工業(股)公司", "容鴻(股)公司", "芙萊柏(股)公司"]
ACTIVE_COMPANY_LABEL = "海濤客食品工業(股)公司"

# -------------------------------------------------------------------------
# 靜態檔案（Logo 圖片）
# 把公司 Logo 檔案放在 app.py 同層的 static/ 資料夾裡（例如
# static/logo.png），系統會自動在網址 /static/logo.png 提供這個檔案，
# 標題列右側就會顯示出來。沒有放檔案時，那個位置只會是空白，不會報錯。
# -------------------------------------------------------------------------
STATIC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")
os.makedirs(STATIC_DIR, exist_ok=True)
app.add_static_files("/static", STATIC_DIR)
LOGO_PATH = os.path.join(STATIC_DIR, "logo.png")


# 初始化全域狀態
initial_df, initial_whs, initial_cats, initial_items_map, initial_customers_map, initial_suppliers_map = fetch_all_a1_inventory()
initial_bom_map, initial_bom_source = load_bom_data()
initial_orders, initial_orders_configured = load_orders_from_google_sheet()
initial_sales_history, initial_sales_configured = load_sales_history_from_google_sheet()
initial_receivings, initial_receivings_configured = load_receivings_from_google_sheet()
initial_channel_sales, initial_channel_sales_configured = load_channel_sales_from_google_sheet()
app_state = {
    "df": initial_df,
    "items_map": initial_items_map,
    "customers_map": initial_customers_map,
    "suppliers_map": initial_suppliers_map,
    "bom_map": initial_bom_map,
    "bom_source": initial_bom_source,
    "orders": initial_orders,
    "orders_configured": initial_orders_configured,
    "sales_history": initial_sales_history,
    "sales_history_configured": initial_sales_configured,
    "sales_history_source": (
        "Google Sheets" if initial_sales_configured else "尚未設定"
    ),
    "channel_sales": initial_channel_sales,
    "channel_sales_configured": initial_channel_sales_configured,
    "receivings": initial_receivings,
    "receivings_configured": initial_receivings_configured,
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
        "slow_moving_days": 90,        # 週轉天數超過此值視為滯銷
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
  # 這個 header（含公司切換分頁）故意寫在 content_container 建立「之前」，
  # 讓它在畫面元素的排列順序上排在所有頁面內容前面，天生就位在最上方。
  # 如果順序反過來（content_container先建立），header雖然有sticky
  # top-0，但因為它在畫面排列順序上落在一大串內容「後面」，一開始會被
  # 排到頁面最下方，資料一多，使用者要滑到最底部才看得到，sticky也救
  # 不了這個順序問題。
  #
  # 注意：這裡只建立 header 本身（含空的 company_tabs），還不能馬上呼叫
  # company_tabs.set_value(...) 來觸發第一次畫面內容，因為 handle_company_change
  # 裡面用到的 render_hai_tao_ke_page() 等函式，要等下面 content_container
  # 建立完、這些函式都 def 好之後才存在。實際觸發放在函式最後面。
  # -----------------------------------------------------------------------
  COMPANY_TAB_COLORS = {
      "興聖(股)公司": {"text": "#5bc0be", "active_bg": "#5bc0be"},
      "海濤客食品工業(股)公司": {"text": "#e0824a", "active_bg": "#e0824a"},
      "容鴻(股)公司": {"text": "#8e7cc3", "active_bg": "#8e7cc3"},
      "芙萊柏(股)公司": {"text": "#5b8fc0", "active_bg": "#5b8fc0"},
  }
  ui.add_head_html(
      "<style>"
      + "".join(
          f".company-tab-{i} {{ color: {c['text']} !important; "
          f"font-weight: 700 !important; }}"
          f".company-tab-{i}.q-tab--active {{ background: {c['active_bg']}22 !important; "
          f"border-bottom: 3px solid {c['active_bg']} !important; }}"
          for i, c in enumerate(COMPANY_TAB_COLORS.values())
      )
      + "</style>"
  )

  def handle_company_change(e):
    selected = e.value
    if selected == ACTIVE_COMPANY_LABEL:
      render_hai_tao_ke_page()
    elif selected in ("興聖(股)公司", "容鴻(股)公司", "芙萊柏(股)公司"):
      render_channel_company_page(selected)
    else:
      render_placeholder_company(selected)

  with ui.row().classes(
      "w-full flex flex-nowrap items-center justify-between bg-white"
      " border-b border-[#e2e1dc] px-8 py-4 sticky top-0 z-50"
  ):
    with ui.row().classes("items-center gap-3 flex-shrink-0"):
      ui.label("興聖集團｜A1 智慧進銷存總管理系統").classes(
          "text-base font-black tracking-wider"
      )
      # Logo 放置位置：把檔案存成 static/logo.png（跟 app.py 同層的
      # static 資料夾）就會自動顯示在這裡，不用改程式碼。沒有檔案時
      # 這裡會是空白，不影響頁面運作。
      if os.path.exists(LOGO_PATH):
        ui.image("/static/logo.png").classes("h-8 w-auto")
    with ui.tabs(on_change=handle_company_change).props(
        "dense no-caps"
    ).classes("flex-shrink-0 ml-auto") as company_tabs:
      for i, c in enumerate(COMPANIES):
        ui.tab(c).classes(f"company-tab-{i}")

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

  def render_section_placeholder(title, hint="此區尚未串接資料來源，敬請期待"):
    """訂單出貨／每日出貨／調撥紀錄／退換貨記錄 共用的「還沒串API」佔位畫面。
    等之後陸續串接各分公司/各通路的 API 時，把對應區塊換成真的資料表格即可，
    版面（分頁結構）不用重搭。
    """
    with ui.card().classes(
        "w-full p-10 bg-white border border-[#e2e1dc] shadow-none"
        " rounded-none text-center"
    ):
      ui.label(title).classes("text-sm font-bold text-zinc-700 mb-2")
      ui.label(hint).classes("text-xs text-zinc-500")

  def render_shopline_pending_orders():
    """興聖(股)公司／訂單出貨／SHOPLINE官網：抓待處理訂單畫成表格。
    每次切換到這個分頁都會重新打一次API（目前沒有做快取／同步按鈕，
    先求資料正確；之後如果需要跟海濤客頁面一樣的「手動刷新」機制，
    再仿照 handle_sync 的寫法加上去）。
    """
    rows, error = fetch_shopline_pending_orders(
        SHOPLINE_XINGSHENG_ACCESS_TOKEN, SHOPLINE_XINGSHENG_USER_AGENT
    )
    if error:
      render_section_placeholder("訂單出貨－SHOPLINE官網", f"抓取失敗：{error}")
      return
    if not rows:
      render_section_placeholder("訂單出貨－SHOPLINE官網", "目前沒有待處理訂單")
      return
    with ui.card().classes(
        "w-full p-6 bg-white border border-[#e2e1dc] shadow-none rounded-none"
    ):
      ui.label(f"待處理訂單（共 {len(rows)} 筆）").classes(
          "text-sm font-bold text-zinc-700 mb-3"
      )
      ui.table(
          columns=[
              {"name": "訂單編號", "label": "訂單編號", "field": "訂單編號", "align": "left"},
              {"name": "客戶", "label": "客戶", "field": "客戶", "align": "left"},
              {"name": "金額", "label": "金額", "field": "金額", "align": "right"},
              {"name": "建立時間", "label": "建立時間（台灣時間）", "field": "建立時間", "align": "left"},
              {"name": "物流狀態", "label": "物流狀態", "field": "物流狀態", "align": "left"},
          ],
          rows=rows,
          row_key="訂單編號",
      ).classes("w-full")

  # 訂單出貨底下的4個通路子分頁，之後每個通路會各自串不同的訂單來源
  # API（SHOPLINE官網／蝦皮／經銷／其它），目前先放佔位畫面
  ORDER_CHANNELS = ["SHOPLINE官網", "蝦皮", "經銷", "其它"]

  # 公司名稱轉成安全的英文代碼，用來組CSS class名稱（中文當class名稱在
  # 部分瀏覽器/選擇器語法下容易出錯，改用英文代碼比較保險）
  COMPANY_SLUGS = {
      "興聖(股)公司": "xingsheng",
      "容鴻(股)公司": "ronghong",
      "芙萊柏(股)公司": "fulaibo",
  }

  def render_channel_company_page(company_name):
    """興聖(股)公司／容鴻(股)公司／芙萊柏(股)公司 共用的頁面骨架：
    儀表板／訂單出貨(4通路)／每日出貨／調撥紀錄／退換貨記錄，共5個分頁。
    目前資料都還沒串接，先讓分頁結構跟導覽長出來，之後每個區塊陸續串上
    真的 API 時，只要把 render_section_placeholder(...) 換成真的內容即可，
    不用動到分頁結構本身。
    """
    content_container.clear()
    accent = COMPANY_TAB_COLORS.get(company_name, {}).get("active_bg", "#5bc0be")
    slug = COMPANY_SLUGS.get(company_name, "default")
    tabs_class = f"section-tabs-{slug}"
    with content_container:
      with ui.column().classes("w-full p-8 max-w-[1600px] mx-auto gap-4"):
        ui.label(company_name).classes(
            "text-lg font-bold text-zinc-900"
        )
        with ui.tabs().props("dense no-caps").classes(
            f"w-full {tabs_class}"
        ) as section_tabs:
          tab_ch_dashboard = ui.tab("儀表板")
          tab_ch_orders = ui.tab("訂單出貨")
          tab_ch_daily_shipping = ui.tab("每日出貨")
          tab_ch_transfer = ui.tab("調撥紀錄")
          tab_ch_returns = ui.tab("退換貨記錄")
        # 只套用在這組分頁自己身上（用.section-tabs-xxx限定範圍），
        # 不會影響到最上面的公司切換列或其他分公司頁面的分頁顏色
        ui.add_head_html(
            f"<style>.{tabs_class}.q-tabs .q-tab--active {{ color: {accent} !important; }}"
            f" .{tabs_class}.q-tabs .q-tab-indicator {{ background: {accent} !important; }}</style>"
        )
        with ui.tab_panels(section_tabs, value=tab_ch_dashboard).classes(
            "w-full bg-transparent"
        ):
          with ui.tab_panel(tab_ch_dashboard):
            render_section_placeholder(
                "儀表板", "尚未串接此分公司的庫存／訂單資料，敬請期待"
            )

          with ui.tab_panel(tab_ch_orders):
            with ui.tabs().props("dense no-caps").classes("w-full") as channel_tabs:
              channel_tab_objs = [ui.tab(ch) for ch in ORDER_CHANNELS]
            with ui.tab_panels(
                channel_tabs, value=channel_tab_objs[0]
            ).classes("w-full bg-transparent"):
              for ch, ch_tab in zip(ORDER_CHANNELS, channel_tab_objs):
                with ui.tab_panel(ch_tab):
                  if company_name == "興聖(股)公司" and ch == "SHOPLINE官網":
                    render_shopline_pending_orders()
                  else:
                    render_section_placeholder(
                        f"訂單出貨－{ch}",
                        f"「{ch}」通路的訂單 API 尚未串接，敬請期待",
                    )

          with ui.tab_panel(tab_ch_daily_shipping):
            render_section_placeholder("每日出貨")

          with ui.tab_panel(tab_ch_transfer):
            render_section_placeholder("調撥紀錄")

          with ui.tab_panel(tab_ch_returns):
            render_section_placeholder("退換貨記錄")

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
          df, whs, cats, items_map, customers_map, suppliers_map = fetch_all_a1_inventory()
          app_state["df"] = df
          app_state["items_map"] = items_map
          app_state["customers_map"] = customers_map
          app_state["suppliers_map"] = suppliers_map
          if whs:
            app_state["warehouses"] = whs
          if cats:
            app_state["categories"] = cats

          # 同步時順便重新讀取三份 Google Sheet 資料，這樣按一次「同步」
          # 就能拿到最新的庫存 + BOM + 訂單 + 銷售歷史，不用分開點好幾個
          # 「重新載入」按鈕（若銷售歷史目前是用「5.3 手動從 A1 抓取」的
          # 結果，這裡就不覆蓋回 Sheets，避免白抓一次又被蓋掉）
          app_state["bom_map"], app_state["bom_source"] = load_bom_data()
          app_state["orders"], app_state["orders_configured"] = (
              load_orders_from_google_sheet()
          )
          if not str(app_state.get("sales_history_source", "")).startswith("鼎新 A1"):
            sheet_rows, sheet_configured = load_sales_history_from_google_sheet()
            app_state["sales_history"] = sheet_rows
            app_state["sales_history_configured"] = sheet_configured
            app_state["sales_history_source"] = (
                "Google Sheets" if sheet_configured else "尚未設定"
            )
          app_state["receivings"], app_state["receivings_configured"] = (
              load_receivings_from_google_sheet()
          )
          app_state["channel_sales"], app_state["channel_sales_configured"] = (
              load_channel_sales_from_google_sheet()
          )

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
        if "procurement_cat_select" in refs:
          refs["procurement_cat_select"].options = ["全部分類"] + app_state["categories"]

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
            "update_orders_list",
            "update_packing_schedule",
            "update_turnover_list",
            "update_receivings_list",
        ):
          if ref_key in refs:
            refs[ref_key]()

      with ui.column().classes("w-full p-8 max-w-[1600px] mx-auto"):
        with ui.row().classes("w-full items-center justify-between mb-4"):
          with ui.row().classes("items-center gap-3"):
            ui.label().classes(
                "w-1.5 h-8 bg-[#5bc0be]"
            )  # 左側色塊，強化視覺焦點
            ui.label(ACTIVE_COMPANY_LABEL).classes(
                "text-2xl font-black text-zinc-900 tracking-wide"
            )
          ui.button("同步 A1 最新庫存", on_click=handle_sync).classes(
              "sync-btn px-4 py-2 text-xs rounded-none"
          )

        with ui.tabs().classes("w-full") as page_tabs:
          tab_dashboard = ui.tab("儀表板")
          tab_products_group = ui.tab("商品資訊")
          tab_orders = ui.tab("訂單出貨")
          tab_production = ui.tab("生產排程")
          tab_procurement = ui.tab("採購分析")
          tab_settings = ui.tab("系統設定")

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
              dashboard_source_label = ui.label().classes(
                  "text-xs text-zinc-500 mb-4"
              )
              dashboard_kpi_row = ui.row().classes("w-full gap-4 mb-6 flex-wrap")

              ui.label("提醒／公告中心").classes(
                  "text-sm font-bold text-zinc-700 mb-2"
              )
              dashboard_announce_container = ui.column().classes("w-full gap-2 mb-6")

              def _severity_box(severity, text, accent_color=None):
                style = SEVERITY_STYLES[severity]
                accent_class = (
                    f"border-l-4 border-l-[{accent_color}]" if accent_color else ""
                )
                with ui.row().classes(
                    f"w-full items-center gap-2 p-2 border {style['box']} {accent_class}"
                ):
                  ui.label(style["label"]).classes(
                      f"text-[11px] px-2 py-0.5 rounded-none font-bold"
                      f" {style['badge']}"
                  )
                  ui.label(text).classes(f"text-xs {style['text']} flex-1")

              def _render_announcements(container, announcements, category_keys, max_each=6):
                """把 compute_dashboard_announcements() 的結果畫成一疊
                _severity_box。category_keys 是要顯示哪幾種公告（例如頁面
                各自只顯示跟自己相關的那種），max_each 限制每種最多顯示
                幾則，避免公告洗版。
                """
                container.clear()
                any_item = False
                with container:
                  for key in category_keys:
                    for item in announcements.get(key, [])[:max_each]:
                      _severity_box(item["severity"], item["text"])
                      any_item = True
                  if not any_item:
                    _severity_box("success", "目前沒有相關提醒")

              # KPI 卡片明細彈窗：所有卡片共用同一個 dialog，點擊時換內容
              with ui.dialog() as kpi_dialog, ui.card().classes(
                  "min-w-[320px] max-w-[90vw] p-5"
              ):
                kpi_dialog_title = ui.label().classes(
                    "text-base font-bold text-zinc-900 mb-3"
                )
                kpi_dialog_body = ui.column().classes("w-full")
                ui.button("關閉", on_click=kpi_dialog.close).classes(
                    "sync-btn px-4 py-1 text-xs rounded-none mt-3 self-end"
                )

              def open_kpi_dialog(title, rows, columns):
                kpi_dialog_title.text = title
                kpi_dialog_body.clear()
                with kpi_dialog_body:
                  if not rows:
                    ui.label("目前沒有符合條件的品項").classes(
                        "text-xs text-zinc-400"
                    )
                  else:
                    ui.table(columns=columns, rows=rows).classes("w-full")

                    def handle_export():
                      try:
                        xlsx_bytes = rows_to_xlsx_bytes(rows, sheet_name=title)
                        safe_filename = "".join(
                            c for c in title if c not in '\\/:*?"<>|'
                        ) or "匯出資料"
                        ui.download(
                            xlsx_bytes,
                            f"{safe_filename}.xlsx",
                            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                      except Exception as e:
                        ui.notify(f"匯出失敗：{e}", color="negative")

                    ui.button("匯出 xlsx", on_click=handle_export).classes(
                        "sync-btn px-3 py-1 text-xs rounded-none mt-3"
                    )
                kpi_dialog.open()

              def _kpi_card(label, value, severity, on_click=None):
                style = SEVERITY_STYLES.get(severity, SEVERITY_STYLES["info"])
                classes = (
                    f"{style['box']} border p-4 min-w-[200px] flex-1"
                )
                if on_click:
                  classes += " cursor-pointer hover:brightness-95"
                card = ui.column().classes(classes)
                with card:
                  ui.label(label).classes(f"text-xs {style['text']} mb-1")
                  ui.label(value).classes(
                      f"text-xl font-black {style['text']}"
                  )
                  if on_click:
                    ui.label("點擊查看明細 →").classes(
                        f"text-[10px] {style['text']} opacity-70 mt-1"
                    )
                if on_click:
                  card.on("click", on_click)
                return card

              def update_dashboard():
                dashboard_kpi_row.clear()
                dashboard_announce_container.clear()

                df = app_state["df"].copy()
                items_map = app_state.get("items_map", {})
                bom_map = app_state.get("bom_map", {})
                orders = app_state.get("orders", [])
                orders_configured = app_state.get("orders_configured", False)
                settings = app_state["settings"]

                dashboard_source_label.text = (
                    f"庫存來源：鼎新 A1｜BOM 來源：{app_state.get('bom_source', '未知')}"
                    f"｜訂單來源：{'Google Sheets' if orders_configured else '尚未設定 Google Sheets'}"
                )

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

                # ---- 低於安全庫存清單（沿用原本邏輯）----
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
                        "目前庫存": ceil_qty(current_stock),
                        "安全存量": ceil_qty(safety_stock),
                        "缺口": ceil_qty(max(safety_stock - current_stock, 0)),
                    })
                risk_rows.sort(key=lambda r: r["缺口"], reverse=True)

                # ---- 1.1/1.2/1.3：訂單需求 + BOM 展開缺貨預警 ----
                today = datetime.now().date()
                result_30 = compute_order_demand_alerts(
                    orders, items_map, bom_map, stock_lookup, settings,
                    horizon_days=30,
                )
                orders_today = [
                    o for o in orders
                    if o["狀態"] != "已出貨" and o["預計出貨日"] == today
                ]
                today_qty = sum(o["預計出貨數量"] for o in orders_today)

                with dashboard_kpi_row:
                  # ---- 卡片1：庫存總值，點擊看價值最高的品項 ----
                  value_detail_rows = []
                  for item_id, info in items_map.items():
                    current_stock = stock_lookup.get(item_id, 0.0)
                    unit_cost = info.get("StdPurPrice") or 0
                    try:
                      item_value = float(current_stock) * float(unit_cost)
                    except (TypeError, ValueError):
                      item_value = 0.0
                    if item_value > 0:
                      value_detail_rows.append({
                          "品號": item_id,
                          "品名": info.get("Name"),
                          "庫存量": ceil_qty(current_stock),
                          "單價": unit_cost,
                          "庫存價值": round(item_value, 0),
                      })
                  value_detail_rows.sort(key=lambda r: r["庫存價值"], reverse=True)
                  value_detail_rows = value_detail_rows[:15]

                  _kpi_card(
                      "集團／本公司庫存總值（估）",
                      f"NT$ {total_value:,.0f}",
                      "info",
                      on_click=lambda e=None: open_kpi_dialog(
                          "庫存價值最高的品項（前 15 名）",
                          value_detail_rows,
                          [
                              {"name": "品號", "label": "品號", "field": "品號", "align": "left"},
                              {"name": "品名", "label": "品名", "field": "品名", "align": "left"},
                              {"name": "庫存量", "label": "庫存量", "field": "庫存量"},
                              {"name": "單價", "label": "單價", "field": "單價"},
                              {"name": "庫存價值", "label": "庫存價值", "field": "庫存價值"},
                          ],
                      ),
                  )

                  # ---- 卡片2：低於安全庫存，點擊看清單 ----
                  _kpi_card(
                      "低於安全庫存品項數",
                      f"{len(risk_rows)} 項",
                      "danger" if risk_rows else "success",
                      on_click=lambda e=None: open_kpi_dialog(
                          "低於安全庫存品項清單",
                          risk_rows,
                          [
                              {"name": "品號", "label": "品號", "field": "品號", "align": "left"},
                              {"name": "品名", "label": "品名", "field": "品名", "align": "left"},
                              {"name": "目前庫存", "label": "目前庫存", "field": "目前庫存"},
                              {"name": "安全存量", "label": "安全存量", "field": "安全存量"},
                              {"name": "缺口", "label": "缺口", "field": "缺口"},
                          ],
                      ),
                  )

                  # ---- 卡片3：今日預計出貨 ----
                  today_display_rows = [
                      {**o, "預計出貨日": o["預計出貨日"].isoformat()}
                      for o in orders_today
                  ]
                  if orders_configured:
                    _kpi_card(
                        "今日預計出貨訂單數／總量",
                        f"{len(orders_today)} 張／{ceil_qty(today_qty)}",
                        "warning" if orders_today else "success",
                        on_click=lambda e=None: open_kpi_dialog(
                            "今日預計出貨訂單",
                            today_display_rows,
                            [
                                {"name": "訂單編號", "label": "訂單編號", "field": "訂單編號", "align": "left"},
                                {"name": "品號", "label": "品號", "field": "品號", "align": "left"},
                                {"name": "品名", "label": "品名", "field": "品名", "align": "left"},
                                {"name": "預計出貨數量", "label": "預計出貨數量", "field": "預計出貨數量"},
                                {"name": "狀態", "label": "狀態", "field": "狀態"},
                            ],
                        ),
                    )
                  else:
                    _kpi_card(
                        "今日預計出貨訂單數／總量",
                        "－（待設定 Google Sheets）",
                        "info",
                    )

                  # ---- 卡片4：未來30天缺貨風險，點擊看成品+原料缺口 ----
                  shortage_count = (
                      len(result_30["finished_goods_shortfall"])
                      + len(result_30["raw_material_shortfall"])
                  )
                  if orders_configured:
                    combined_shortage_rows = [
                        {
                            "類型": "成品",
                            "品號": r["品號"],
                            "品名": r["品名"],
                            "缺口": r["缺口"],
                            "說明": f"最早出貨日 {r['最早出貨日']}",
                        }
                        for r in result_30["finished_goods_shortfall"]
                    ] + [
                        {
                            "類型": "原料/子件",
                            "品號": r["品號"],
                            "品名": r["品名"],
                            "缺口": r["缺口"],
                            "說明": f"建議下單日 {r['建議下單日']}",
                        }
                        for r in result_30["raw_material_shortfall"]
                    ]
                    _kpi_card(
                        "未來30天缺貨風險品項",
                        f"{shortage_count} 項",
                        "danger" if shortage_count else "success",
                        on_click=lambda e=None: open_kpi_dialog(
                            "未來30天缺貨風險品項（成品＋原料/子件）",
                            combined_shortage_rows,
                            [
                                {"name": "類型", "label": "類型", "field": "類型", "align": "left"},
                                {"name": "品號", "label": "品號", "field": "品號", "align": "left"},
                                {"name": "品名", "label": "品名", "field": "品名", "align": "left"},
                                {"name": "缺口", "label": "缺口", "field": "缺口"},
                                {"name": "說明", "label": "說明", "field": "說明", "align": "left"},
                            ],
                        ),
                    )
                  else:
                    _kpi_card(
                        "未來30天缺貨風險品項",
                        "－（待設定 Google Sheets）",
                        "info",
                    )

                # ---- 提醒／公告中心：4 種提醒統一用顏色分級 ----
                with dashboard_announce_container:
                  if not orders_configured:
                    ui.label(
                        "尚未設定 Google Sheets「訂單資訊」分頁，暫時無法"
                        "顯示出貨與補貨提醒，設定方式見「系統設定」。"
                    ).classes("text-xs text-zinc-400")
                  else:
                    announcements = compute_dashboard_announcements(
                        orders, items_map, bom_map, stock_lookup, settings,
                        horizon_days=30,
                    )
                    category_labels = [
                        ("shipping", "🚚 訂單出貨提醒", "#2563eb"),
                        ("production", "🏭 生產組裝確認", "#9333ea"),
                        ("finished_goods", "🧾 建議採購成品（母件）", "#db2777"),
                    ]
                    any_announcement = False
                    for key, label, accent_color in category_labels:
                      items = announcements.get(key, [])[:6]
                      if not items:
                        continue
                      ui.label(label).classes(
                          "text-xs font-bold text-zinc-600 mt-1"
                      )
                      for item in items:
                        _severity_box(item["severity"], item["text"], accent_color)
                      any_announcement = True

                    if not any_announcement:
                      _severity_box(
                          "success", "未來 30 天內沒有已知的出貨、生產、"
                          "採購或進貨提醒"
                      )

              update_dashboard()
              refs["update_dashboard"] = update_dashboard

          # ==================================================
          # 2. 商品與組合管理（原本的 3 個頁籤 + 新增批號/效期追蹤）
          # ==================================================
          with ui.tab_panel(tab_products_group):
            with ui.tabs().classes("w-full mb-2") as sub_tabs:
              tab_products = ui.tab("商品資料")
              tab_inventory = ui.tab("庫存查詢")
              tab_bom = ui.tab("商品組合(BOM)")
              tab_lotno = ui.tab("批號效期")

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
                                f"庫存：{ceil_qty(row['庫存數量'])} {safe_text(row['單位'])}"
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
                      display_rows = df.to_dict("records")
                      for r in display_rows:
                        r["庫存數量"] = ceil_qty(r.get("庫存數量"))
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
                          rows=display_rows,
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

                    keyword = (combo_search_input.value or "").strip().lower()

                    # 主清單：只列出「商品組合明細」Sheet 裡實際有填資料的主件。
                    # 之前的版本會把 A1 商品主檔裡所有「商品型態=組合品」的
                    # 品項都列出來——但 A1 裡光是型態標記就有 100 多筆，跟
                    # 貴公司實際維護的 BOM 資料量完全對不上，變成滿頁都是
                    # 「尚未補齊子件明細」的空殼，沒有意義。改成只顯示已經
                    # 在 Sheet 裡設定過的品項，畫面才會跟你實際填的資料一致。
                    configured_entries = [
                        (item_id, components)
                        for item_id, components in bom_map.items()
                        if item_id in items_map
                    ]
                    if keyword:
                      configured_entries = [
                          (item_id, components)
                          for item_id, components in configured_entries
                          if keyword in str(item_id).lower()
                          or keyword in str(items_map[item_id].get("Name", "")).lower()
                      ]

                    # 次清單：A1 裡標記為組合品、但還沒在 Sheet 建立子件資料的
                    # 品項——只顯示「數量」跟一個可展開的清單，不要每個都自動
                    # 展開撐爆頁面
                    configured_ids = {pid for pid, _ in configured_entries}
                    unconfigured_entries = [
                        (item_id, info)
                        for item_id, info in items_map.items()
                        if str(info.get("Type")) in ("2", "3")
                        and item_id not in bom_map
                    ]
                    if keyword:
                      unconfigured_entries = [
                          (item_id, info)
                          for item_id, info in unconfigured_entries
                          if keyword in str(item_id).lower()
                          or keyword in str(info.get("Name", "")).lower()
                      ]

                    combo_stats_label.text = (
                        f"已設定子件明細的組合品：{len(configured_entries)} 項"
                        f"｜A1 中標記為組合品但尚未設定：{len(unconfigured_entries)} 項"
                    )

                    with combo_list_container:
                      if not configured_entries and not unconfigured_entries:
                        ui.label(
                            "目前沒有標記為組合品的商品，或尚未同步商品資料"
                        ).classes("text-xs text-zinc-400")

                      for item_id, components in configured_entries:
                        info = items_map[item_id]
                        type_label = ITEM_TYPE_LABELS.get(
                            str(info.get("Type")), str(info.get("Type"))
                        )
                        with ui.expansion(
                            f"{item_id}｜{info.get('Name', '')}｜{type_label}"
                            f"（{len(components)} 項子件）",
                            icon="inventory_2",
                            value=True,
                        ).classes(
                            "w-full border border-[#e2e1dc] text-sm"
                        ):
                          ui.table(
                              columns=[
                                  {"name": "子件品號", "label": "子件品號", "field": "子件品號", "align": "left"},
                                  {"name": "子件品名", "label": "子件品名", "field": "子件品名", "align": "left"},
                                  {"name": "用量", "label": "用量", "field": "用量"},
                                  {"name": "單位", "label": "單位", "field": "單位"},
                                  {"name": "損耗率", "label": "損耗率(%)", "field": "損耗率"},
                                  {"name": "採購前置天數", "label": "採購前置天數", "field": "採購前置天數"},
                                  {"name": "生產工時天數", "label": "生產工時(天)", "field": "生產工時天數"},
                                  {"name": "供應商", "label": "供應商", "field": "供應商", "align": "left"},
                                  {"name": "備註", "label": "備註", "field": "備註", "align": "left"},
                              ],
                              rows=components,
                          ).classes("w-full")

                      if unconfigured_entries:
                        with ui.expansion(
                            f"A1 中另有 {len(unconfigured_entries)} 個品項標記為"
                            f"組合品，但尚未在商品組合明細中設定子件（點擊展開"
                            f"清單）",
                            icon="help_outline",
                            value=False,
                        ).classes(
                            "w-full border border-[#e2e1dc] text-sm text-zinc-500"
                        ):
                          ui.table(
                              columns=[
                                  {"name": "品號", "label": "品號", "field": "品號", "align": "left"},
                                  {"name": "品名", "label": "品名", "field": "品名", "align": "left"},
                                  {"name": "商品型態", "label": "商品型態", "field": "商品型態", "align": "left"},
                              ],
                              rows=[
                                  {
                                      "品號": item_id,
                                      "品名": info.get("Name"),
                                      "商品型態": ITEM_TYPE_LABELS.get(
                                          str(info.get("Type")), str(info.get("Type"))
                                      ),
                                  }
                                  for item_id, info in unconfigured_entries
                              ],
                              pagination=10,
                          ).classes("w-full")

                      # 提醒：Sheet 裡有填，但目前 A1 商品主檔查無此品號的主件
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
                  "w-full p-3 mb-4 bg-[#e8f6f5] border border-[#bfe6e3]"
              ):
                ui.label(
                    "資料來源：Google Sheets「訂單資訊」分頁（手動覆蓋更新，"
                    "非即時自動抓取；A1 本身沒有查詢訂單的 API，只能上傳，"
                    "所以這裡改讀 Sheet）。只需要維護「品號、預計出貨日、"
                    "預計出貨數量」，後面的 BOM 表會自動判斷是否需要補貨；"
                    "「客戶代號」「金額」只有要同步到 A1（下方按鈕）才需要"
                    "填寫。"
                ).classes("text-xs text-teal-800")

              # 反向同步到 A1：這是「寫入」正式系統的動作，故意做成手動
              # 觸發＋確認彈窗，不會自動執行，避免誤觸
              with ui.dialog() as order_sync_dialog, ui.card().classes(
                  "min-w-[360px] max-w-[90vw] p-5"
              ):
                order_sync_dialog_body = ui.column().classes("w-full gap-2")
                with ui.row().classes("w-full justify-end gap-2 mt-4"):
                  ui.button(
                      "取消", on_click=order_sync_dialog.close
                  ).classes("px-4 py-1 text-xs rounded-none")
                  order_sync_confirm_button = ui.button("確認上傳").classes(
                      "sync-btn px-4 py-1 text-xs rounded-none"
                  )

              def handle_order_sync_click():
                orders = app_state.get("orders", [])
                payloads, skipped = build_order_upload_payloads(orders)
                order_sync_dialog_body.clear()
                with order_sync_dialog_body:
                  ui.label("確認上傳訂單到 A1").classes(
                      "text-base font-bold text-zinc-900"
                  )
                  ui.label(
                      f"準備上傳 {len(payloads)} 張訂單到正式 A1 系統"
                      f"（Orders[Post]）。這個動作會寫入你們正式的進銷存"
                      f"資料，請確認客戶代號、金額都正確再繼續。"
                  ).classes("text-xs text-zinc-700")
                  if skipped:
                    ui.label(
                        f"另有 {len(skipped)} 張因缺客戶代號或金額，"
                        f"這次不會上傳（可在下方查看明細）。"
                    ).classes("text-xs text-amber-700")
                    with ui.expansion("查看略過明細").classes("w-full text-xs"):
                      for s in skipped[:20]:
                        ui.label(s).classes("text-xs text-zinc-500")

                def handle_confirm_upload():
                  token = get_a1_token()
                  if not token:
                    ui.notify("無法登入 A1，請確認 API 憑證", color="warning")
                    return
                  success, duplicate, failed = 0, 0, []
                  for display_id, payload in payloads:
                    ok, msg = upload_order_to_a1(token, payload)
                    if ok:
                      success += 1
                    elif "重複" in msg:
                      duplicate += 1
                    else:
                      failed.append(f"{display_id}: {msg}")
                  order_sync_dialog.close()
                  summary = (
                      f"上傳完成：成功 {success} 張／已存在略過 {duplicate}"
                      f" 張／失敗 {len(failed)} 張"
                  )
                  ui.notify(
                      summary, color="positive" if not failed else "warning"
                  )
                  if failed:
                    print("Orders[Post] 上傳失敗明細:\n" + "\n".join(failed))

                order_sync_confirm_button.on_click(handle_confirm_upload)
                order_sync_dialog.open()

              with ui.row().classes("items-center gap-3 flex-wrap mb-4"):
                ui.button(
                    "同步訂單到 A1（寫入正式系統）",
                    on_click=handle_order_sync_click,
                ).classes(
                    "px-3 py-1 text-xs rounded-none bg-amber-600 text-white"
                    " font-bold"
                )
                ui.label(
                    "⚠ 會實際寫入 A1，上傳前請先確認 Sheet 裡的客戶代號／"
                    "金額都填對"
                ).classes("text-xs text-amber-700")

              orders_reminder_container = ui.column().classes("w-full gap-2 mb-4")

              with ui.row().classes("items-center gap-3 flex-wrap mb-3"):
                orders_search_input = ui.input(
                    placeholder="輸入品號、品名或訂單編號..."
                ).classes("w-64 text-xs")
                orders_status_select = ui.select(
                    options=["全部狀態", "未出貨", "備貨中", "已出貨"],
                    value="全部狀態",
                ).classes(
                    "bg-[#f7f6f2] text-zinc-900 rounded-none px-3 py-1"
                    " text-xs font-bold border border-[#e2e1dc]"
                )

              orders_stats_label = ui.label().classes(
                  "text-xs text-zinc-500 mb-3"
              )
              orders_table_container = ui.column().classes("w-full")

              def update_orders_list():
                orders_table_container.clear()
                orders = app_state.get("orders", [])
                configured = app_state.get("orders_configured", False)

                if not configured:
                  orders_reminder_container.clear()
                  orders_stats_label.text = ""
                  with orders_table_container:
                    ui.label(
                        "尚未設定 Google Sheets，請見「系統設定」的設定"
                        "說明。"
                    ).classes("text-xs text-zinc-400")
                  return

                items_map = app_state.get("items_map", {})
                bom_map = app_state.get("bom_map", {})
                df = app_state["df"].copy()
                settings = app_state["settings"]
                if not df.empty:
                  stock_by_item = df.groupby("品號", as_index=False)[
                      "庫存數量"
                  ].sum()
                  stock_lookup = dict(
                      zip(stock_by_item["品號"], stock_by_item["庫存數量"])
                  )
                else:
                  stock_lookup = {}
                announcements = compute_dashboard_announcements(
                    orders, items_map, bom_map, stock_lookup, settings,
                    horizon_days=14,
                )
                _render_announcements(
                    orders_reminder_container, announcements, ["shipping"]
                )

                rows = list(orders)
                keyword = (orders_search_input.value or "").strip().lower()
                if keyword:
                  rows = [
                      r for r in rows
                      if keyword in str(r["品號"]).lower()
                      or keyword in str(r.get("品名", "")).lower()
                      or keyword in str(r.get("訂單編號", "")).lower()
                  ]
                if orders_status_select.value != "全部狀態":
                  rows = [r for r in rows if r["狀態"] == orders_status_select.value]

                rows = sorted(rows, key=lambda r: r["預計出貨日"])
                display_rows = [
                    {**r, "預計出貨日": r["預計出貨日"].isoformat()}
                    for r in rows
                ]

                total_qty = sum(r["預計出貨數量"] for r in rows)
                orders_stats_label.text = (
                    f"共 {len(rows)} 筆訂單｜預計出貨總量 {total_qty:g}"
                )

                with orders_table_container:
                  if not rows:
                    ui.label("目前沒有符合條件的訂單資料").classes(
                        "text-xs text-zinc-400"
                    )
                  else:
                    ui.table(
                        columns=[
                            {"name": "訂單編號", "label": "訂單編號", "field": "訂單編號", "align": "left"},
                            {"name": "品號", "label": "品號", "field": "品號", "align": "left"},
                            {"name": "品名", "label": "品名", "field": "品名", "align": "left"},
                            {"name": "預計出貨日", "label": "預計出貨日", "field": "預計出貨日"},
                            {"name": "預計出貨數量", "label": "預計出貨數量", "field": "預計出貨數量"},
                            {"name": "狀態", "label": "狀態", "field": "狀態"},
                            {"name": "備註", "label": "備註", "field": "備註", "align": "left"},
                        ],
                        rows=display_rows,
                    ).classes("w-full")

              orders_search_input.on_value_change(lambda e: update_orders_list())
              orders_status_select.on_value_change(lambda e: update_orders_list())
              update_orders_list()
              refs["update_orders_list"] = update_orders_list

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
                  "w-full p-3 mb-4 bg-[#e8f6f5] border border-[#bfe6e3]"
              ):
                ui.label(
                    "依「訂單資訊」的預計出貨日排序，並用「商品組合"
                    "(BOM)」展開子件用量，跟目前庫存比對——原料/半成品不夠"
                    "的品項會標示出來，方便提前備料。同一套邏輯跟「儀表板"
                    "」「訂單出貨」共用，數字會一致。"
                ).classes("text-xs text-teal-800")

              packing_reminder_container = ui.column().classes("w-full gap-2 mb-4")

              packing_stats_label = ui.label().classes(
                  "text-xs text-zinc-500 mb-2"
              )
              packing_schedule_container = ui.column().classes("w-full mb-6")

              ui.label("原物料備料需求（未來 30 天內，依建議下單日排序）").classes(
                  "text-sm font-bold text-zinc-700 mb-2"
              )
              packing_material_container = ui.column().classes("w-full")

              def update_packing_schedule():
                packing_schedule_container.clear()
                packing_material_container.clear()

                orders = app_state.get("orders", [])
                configured = app_state.get("orders_configured", False)
                if not configured:
                  packing_reminder_container.clear()
                  packing_stats_label.text = ""
                  with packing_schedule_container:
                    ui.label(
                        "尚未設定 Google Sheets，請見「系統設定」的設定"
                        "說明。"
                    ).classes("text-xs text-zinc-400")
                  return

                items_map = app_state.get("items_map", {})
                bom_map = app_state.get("bom_map", {})
                df = app_state["df"].copy()
                settings = app_state["settings"]

                if not df.empty:
                  stock_by_item = df.groupby("品號", as_index=False)[
                      "庫存數量"
                  ].sum()
                  stock_lookup = dict(
                      zip(stock_by_item["品號"], stock_by_item["庫存數量"])
                  )
                else:
                  stock_lookup = {}

                announcements = compute_dashboard_announcements(
                    orders, items_map, bom_map, stock_lookup, settings,
                    horizon_days=14,
                )
                _render_announcements(
                    packing_reminder_container, announcements, ["production"]
                )

                result = compute_order_demand_alerts(
                    orders, items_map, bom_map, stock_lookup, settings,
                    horizon_days=30,
                )
                shortfall_by_item = {
                    r["品號"]: r for r in result["finished_goods_shortfall"]
                }

                today = datetime.now().date()
                schedule_rows = sorted(
                    result["orders_in_horizon"], key=lambda o: o["預計出貨日"]
                )
                packing_stats_label.text = (
                    f"未來 30 天共 {len(schedule_rows)} 筆待包裝/出貨訂單"
                )

                with packing_schedule_container:
                  if not schedule_rows:
                    ui.label("未來 30 天內沒有待處理的訂單").classes(
                        "text-xs text-zinc-400"
                    )
                  else:
                    for o in schedule_rows:
                      days_left = (o["預計出貨日"] - today).days
                      has_shortage = o["品號"] in shortfall_by_item
                      severity = (
                          "danger" if has_shortage and days_left <= 3
                          else "warning" if has_shortage
                          else "success" if days_left <= 1
                          else "info"
                      )
                      status_text = (
                          "⚠ 成品庫存不足，需先確認能否即時生產/組裝"
                          if has_shortage else "成品庫存足夠，可直接安排包裝"
                      )
                      _severity_box(
                          severity,
                          f"{o['預計出貨日'].isoformat()}（{days_left}天後）"
                          f"｜{o.get('品名') or o['品號']}｜數量"
                          f" {o['預計出貨數量']:g}｜{status_text}",
                      )

                with packing_material_container:
                  if not result["raw_material_shortfall"]:
                    ui.label("目前沒有原物料缺口").classes(
                        "text-xs text-zinc-400"
                    )
                  else:
                    ui.table(
                        columns=[
                            {"name": "品號", "label": "品號", "field": "品號", "align": "left"},
                            {"name": "品名", "label": "品名", "field": "品名", "align": "left"},
                            {"name": "未來需求量(含損耗)", "label": "未來需求量(含損耗)", "field": "未來需求量(含損耗)"},
                            {"name": "現有庫存", "label": "現有庫存", "field": "現有庫存"},
                            {"name": "缺口", "label": "缺口", "field": "缺口"},
                            {"name": "採購前置天數", "label": "採購前置天數", "field": "採購前置天數"},
                            {"name": "建議下單日", "label": "建議下單日", "field": "建議下單日"},
                        ],
                        rows=result["raw_material_shortfall"],
                    ).classes("w-full")

              update_packing_schedule()
              refs["update_packing_schedule"] = update_packing_schedule

          # ==================================================
          # 5. 採購分析與決策支援
          # ==================================================
          with ui.tab_panel(tab_procurement):
            ui.label("採購分析與決策支援").classes(
                "text-lg font-bold text-zinc-900 tracking-wide mb-2"
            )
            with ui.row().classes(
                "w-full p-3 mb-4 bg-[#fff8e6] border border-[#f0dca0]"
            ):
              ui.label(
                  "「建議採購量」只用「安全庫存 − 現有庫存」計算，"
                  "沒有把訂單需求算進去；含訂單/BOM 展開的完整版本在"
                  "「生產排程」的「原物料備料需求」表，兩者算法"
                  "不同、用途也不同（這裡是長期安全庫存基準，第4點是短期"
                  "訂單驅動的緊急採購）。「供應商歷史採購單價走勢」仍是"
                  "規劃中，需要額外記錄歷次採購單價，目前 A1/Sheets 都還"
                  "沒有這份資料。"
              ).classes("text-xs text-amber-800")

            with ui.tabs().classes("w-full mb-2") as sub_tabs_procurement:
              tab_procurement_51 = ui.tab("建議採購量")
              tab_procurement_53 = ui.tab("庫存週轉")
              tab_procurement_54 = ui.tab("進貨明細")
              tab_procurement_55 = ui.tab("月產銷分析")

            with ui.tab_panels(
                sub_tabs_procurement, value=tab_procurement_51
            ).classes("w-full bg-transparent"):
              with ui.tab_panel(tab_procurement_51):
                with ui.card().classes(
                    "w-full p-6 bg-white border border-[#e2e1dc] shadow-none"
                    " rounded-none"
                ):
                  procurement_reminder_container = ui.column().classes("w-full gap-2 mb-4")

                  ui.label("建議採購量（安全庫存基準）").classes(
                      "text-sm font-bold text-zinc-700 mb-2"
                  )

                  with ui.row().classes("items-center gap-3 flex-wrap mb-2"):
                    procurement_cat_options = ["全部分類"] + app_state["categories"]
                    procurement_cat_select = ui.select(
                        options=procurement_cat_options, value="全部分類"
                    ).classes(
                        "bg-[#f7f6f2] text-zinc-900 rounded-none px-3 py-1"
                        " text-xs font-bold border border-[#e2e1dc]"
                    )
                    procurement_search_input = ui.input(
                        placeholder="輸入品號或品名關鍵字..."
                    ).classes("w-64 text-xs")
                    procurement_scope_select = ui.select(
                        options=["僅顯示需要採購", "顯示全部商品"],
                        value="僅顯示需要採購",
                    ).classes(
                        "bg-[#f7f6f2] text-zinc-900 rounded-none px-3 py-1"
                        " text-xs font-bold border border-[#e2e1dc]"
                    )

                  with ui.row().classes(
                      "w-full p-2 mb-2 bg-[#e8f6f5] border border-[#bfe6e3]"
                  ):
                    ui.label(
                        "⚠ 預設只顯示「現有庫存 ≤ 安全庫存」的品項。如果你"
                        "覺得少了很多商品，多半是因為那些商品在 A1 商品主檔"
                        "裡沒有設定安全存量(SafetyStock)——沒設定就沒有基準"
                        "可以判斷要不要採購，這裡自動略過。切成「顯示全部"
                        "商品」可以看到所有品項（含未設定安全存量的），"
                        "方便你確認、或去 A1 補上安全存量設定。"
                    ).classes("text-xs text-teal-800")

                  procurement_stats_label = ui.label().classes(
                      "text-xs text-zinc-500 mb-3"
                  )
                  procurement_list_container = ui.column().classes("w-full")

                  def update_procurement_list():
                    procurement_list_container.clear()
                    df = app_state["df"].copy()
                    items_map = app_state.get("items_map", {})
                    bom_map = app_state.get("bom_map", {})

                    orders = app_state.get("orders", [])
                    if app_state.get("orders_configured", False):
                      if not df.empty:
                        stock_by_item_r = df.groupby("品號", as_index=False)[
                            "庫存數量"
                        ].sum()
                        stock_lookup_r = dict(
                            zip(stock_by_item_r["品號"], stock_by_item_r["庫存數量"])
                        )
                      else:
                        stock_lookup_r = {}
                      announcements = compute_dashboard_announcements(
                          orders, items_map, bom_map, stock_lookup_r,
                          app_state["settings"], horizon_days=14,
                      )
                      _render_announcements(
                          procurement_reminder_container, announcements,
                          ["procurement", "incoming"],
                      )
                    else:
                      procurement_reminder_container.clear()

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

                    show_all = procurement_scope_select.value == "顯示全部商品"

                    rows = []
                    for item_id, info in items_map.items():
                      safety_stock = info.get("SafetyStock")
                      try:
                        safety_stock = float(safety_stock)
                      except (TypeError, ValueError):
                        safety_stock = 0.0
                      current_stock = stock_lookup.get(item_id, 0.0)
                      net_need = round(safety_stock - current_stock, 2)

                      if not show_all and (safety_stock <= 0 or net_need <= 0):
                        continue

                      rows.append({
                          "品號": item_id,
                          "品名": info.get("Name"),
                          "商品分類": info.get("CategoryName") or "未分類",
                          "現有庫存": ceil_qty(current_stock),
                          "安全庫存": ceil_qty(safety_stock),
                          "建議採購量（簡化版）": ceil_qty(max(net_need, 0)),
                          "參考前置天數": lead_time_by_child.get(
                              item_id, default_lead_time
                          ),
                      })

                    if procurement_cat_select.value and procurement_cat_select.value != "全部分類":
                      rows = [
                          r for r in rows
                          if r["商品分類"] == procurement_cat_select.value
                      ]

                    keyword = (procurement_search_input.value or "").strip().lower()
                    if keyword:
                      rows = [
                          r for r in rows
                          if keyword in str(r["品號"]).lower()
                          or keyword in str(r["品名"]).lower()
                      ]

                    rows.sort(key=lambda r: r["建議採購量（簡化版）"], reverse=True)
                    procurement_stats_label.text = (
                        f"共 {len(rows)} 項"
                        + ("（顯示全部商品）" if show_all else "（僅顯示需要採購的品項）")
                    )

                    with procurement_list_container:
                      if not rows:
                        ui.label(
                            "目前沒有符合條件的品項，或商品主檔尚未設定安全庫存"
                        ).classes("text-xs text-zinc-400")
                      else:
                        ui.table(
                            columns=[
                                {"name": c, "label": c, "field": c, "align": "left" if c in ("品號", "品名", "商品分類") else "right"}
                                for c in rows[0].keys()
                            ],
                            rows=rows,
                            pagination=10,
                        ).classes("w-full")

                  procurement_cat_select.on_value_change(lambda e: update_procurement_list())
                  procurement_search_input.on_value_change(lambda e: update_procurement_list())
                  procurement_scope_select.on_value_change(lambda e: update_procurement_list())
                  update_procurement_list()
                  refs["update_procurement_list"] = update_procurement_list
                  refs["procurement_cat_select"] = procurement_cat_select

              with ui.tab_panel(tab_procurement_53):
                with ui.card().classes(
                    "w-full p-6 bg-white border border-[#e2e1dc] shadow-none"
                    " rounded-none"
                ):
                  ui.label("庫存週轉率／滯銷品分析").classes(
                      "text-sm font-bold text-zinc-700 mb-2"
                  )
                  with ui.row().classes(
                      "w-full p-3 mb-4 bg-[#e8f6f5] border border-[#bfe6e3]"
                  ):
                    ui.label(
                        "資料來源可以是 Google Sheets「銷售歷史」分頁，或直接用"
                        "手冊記載的 GetSales／GetSaleReturns 端點向 A1 即時抓取"
                        "（下方按鈕）。用近 3 個月平均月銷量算週轉天數 = 現有"
                        "庫存 ÷ 日均銷量；週轉天數超過設定值（預設 90 天）或"
                        "完全沒賣出過但還有庫存，標記為滯銷。"
                    ).classes("text-xs text-teal-800")

                  with ui.row().classes("items-center gap-3 flex-wrap mb-2"):
                    turnover_source_label = ui.label().classes(
                        "text-xs text-zinc-500"
                    )

                    def handle_fetch_sales_from_a1():
                      token = get_a1_token()
                      if not token:
                        ui.notify("無法登入 A1，請確認 API 憑證", color="warning")
                        return
                      ui.notify(
                          "開始向 A1 抓取近 3 個月銷貨/銷退資料，需要幾秒到"
                          "十幾秒，請稍候…",
                          color="info",
                      )
                      try:
                        rows = fetch_sales_history_from_a1(token, months_back=3)
                      except Exception as e:
                        ui.notify(f"抓取失敗：{e}", color="negative")
                        return
                      if not rows:
                        ui.notify(
                            "A1 近 3 個月沒有查到銷貨資料（可能該期間確實無"
                            "交易，或帳號無此權限）",
                            color="warning",
                        )
                        return
                      app_state["sales_history"] = rows
                      app_state["sales_history_configured"] = True
                      app_state["sales_history_source"] = "鼎新 A1（GetSales/GetSaleReturns，近3個月）"
                      ui.notify(
                          f"已從 A1 抓取 {len(rows)} 筆銷售歷史彙總資料",
                          color="positive",
                      )
                      update_turnover_list()

                    ui.button(
                        "從 A1 抓取近3個月銷售歷史", on_click=handle_fetch_sales_from_a1
                    ).classes("sync-btn px-3 py-1 text-xs rounded-none")

                  turnover_stats_label = ui.label().classes(
                      "text-xs text-zinc-500 mb-3"
                  )
                  turnover_table_container = ui.column().classes("w-full")

                  def update_turnover_list():
                    turnover_table_container.clear()
                    sales_history = app_state.get("sales_history", [])
                    configured = app_state.get("sales_history_configured", False)
                    turnover_source_label.text = (
                        f"目前資料來源：{app_state.get('sales_history_source', '尚未設定')}"
                    )

                    if not configured:
                      turnover_stats_label.text = ""
                      with turnover_table_container:
                        ui.label(
                            "尚未設定 Google Sheets，請見「系統設定」的設定"
                            "說明。"
                        ).classes("text-xs text-zinc-400")
                      return

                    df = app_state["df"].copy()
                    items_map = app_state.get("items_map", {})
                    if not df.empty:
                      stock_by_item = df.groupby("品號", as_index=False)[
                          "庫存數量"
                      ].sum()
                      stock_lookup = dict(
                          zip(stock_by_item["品號"], stock_by_item["庫存數量"])
                      )
                    else:
                      stock_lookup = {}

                    slow_moving_days = app_state["settings"]["slow_moving_days"]
                    turnover_rows = compute_turnover_metrics(
                        sales_history, stock_lookup, items_map, slow_moving_days
                    )
                    slow_count = sum(1 for r in turnover_rows if r["滯銷"] == "是")
                    turnover_stats_label.text = (
                        f"共 {len(turnover_rows)} 項有銷售紀錄的品項｜"
                        f"其中 {slow_count} 項判定為滯銷"
                    )

                    with turnover_table_container:
                      if not turnover_rows:
                        ui.label("尚無銷售歷史資料").classes(
                            "text-xs text-zinc-400"
                        )
                      else:
                        ui.table(
                            columns=[
                                {"name": "品號", "label": "品號", "field": "品號", "align": "left"},
                                {"name": "品名", "label": "品名", "field": "品名", "align": "left"},
                                {"name": "近3月平均月銷", "label": "近3月平均月銷", "field": "近3月平均月銷"},
                                {"name": "現有庫存", "label": "現有庫存", "field": "現有庫存"},
                                {"name": "庫存週轉天數", "label": "庫存週轉天數", "field": "庫存週轉天數"},
                                {"name": "滯銷", "label": "滯銷", "field": "滯銷"},
                            ],
                            rows=turnover_rows,
                        ).classes("w-full")

                  update_turnover_list()
                  refs["update_turnover_list"] = update_turnover_list

              with ui.tab_panel(tab_procurement_54):
                with ui.card().classes(
                    "w-full p-6 bg-white border border-[#e2e1dc] shadow-none"
                    " rounded-none"
                ):
                  ui.label("進貨明細").classes(
                      "text-sm font-bold text-zinc-700 mb-2"
                  )
                  with ui.row().classes(
                      "w-full p-3 mb-4 bg-[#e8f6f5] border border-[#bfe6e3]"
                  ):
                    ui.label(
                        "資料來源：Google Sheets「進貨明細」分頁。A1 的"
                        "「進銷存報表」雖然能在後台匯出進退貨明細表，但那是"
                        "後台網頁報表，這份 API 串接手冊（1.0.35）裡 Receives "
                        "只有 [Post] 上傳、沒有對應的查詢端點，所以沒辦法直接"
                        "用 API 抓，改用 Sheet 維護（可以把 A1 匯出的報表複製"
                        "貼上進來，比一筆一筆手動輸入快）。"
                    ).classes("text-xs text-teal-800")

                  with ui.row().classes("items-center gap-3 flex-wrap mb-3"):
                    receiving_search_input = ui.input(
                        placeholder="輸入品號、品名或供應商..."
                    ).classes("w-64 text-xs")

                  receiving_stats_label = ui.label().classes(
                      "text-xs text-zinc-500 mb-3"
                  )
                  receiving_table_container = ui.column().classes("w-full")

                  def update_receivings_list():
                    receiving_table_container.clear()
                    receivings = app_state.get("receivings", [])
                    configured = app_state.get("receivings_configured", False)

                    if not configured:
                      receiving_stats_label.text = ""
                      with receiving_table_container:
                        ui.label(
                            "尚未設定 Google Sheets，請見「系統設定」的設定"
                            "說明。"
                        ).classes("text-xs text-zinc-400")
                      return

                    rows = list(receivings)
                    keyword = (receiving_search_input.value or "").strip().lower()
                    if keyword:
                      rows = [
                          r for r in rows
                          if keyword in str(r["品號"]).lower()
                          or keyword in str(r.get("品名", "")).lower()
                          or keyword in str(r.get("供應商", "")).lower()
                      ]

                    rows = sorted(rows, key=lambda r: r["進貨日期"], reverse=True)
                    display_rows = [
                        {**r, "進貨日期": r["進貨日期"].isoformat()}
                        for r in rows
                    ]

                    total_qty = sum(r["進貨數量"] for r in rows)
                    receiving_stats_label.text = (
                        f"共 {len(rows)} 筆進貨紀錄｜總進貨數量 {total_qty:g}"
                    )

                    with receiving_table_container:
                      if not rows:
                        ui.label("目前沒有符合條件的進貨紀錄").classes(
                            "text-xs text-zinc-400"
                        )
                      else:
                        ui.table(
                            columns=[
                                {"name": "進貨日期", "label": "進貨日期", "field": "進貨日期"},
                                {"name": "品號", "label": "品號", "field": "品號", "align": "left"},
                                {"name": "品名", "label": "品名", "field": "品名", "align": "left"},
                                {"name": "進貨數量", "label": "進貨數量", "field": "進貨數量"},
                                {"name": "單價", "label": "單價", "field": "單價"},
                                {"name": "供應商", "label": "供應商", "field": "供應商", "align": "left"},
                                {"name": "備註", "label": "備註", "field": "備註", "align": "left"},
                            ],
                            rows=display_rows,
                        ).classes("w-full")

                  receiving_search_input.on_value_change(lambda e: update_receivings_list())
                  update_receivings_list()
                  refs["update_receivings_list"] = update_receivings_list

              with ui.tab_panel(tab_procurement_55):
                with ui.card().classes(
                    "w-full p-6 bg-white border border-[#e2e1dc] shadow-none"
                    " rounded-none"
                ):
                  ui.label("月產銷分析").classes(
                      "text-lg font-bold text-zinc-900 tracking-wide mb-2"
                  )
                  with ui.row().classes(
                      "w-full p-3 mb-4 bg-[#e8f6f5] border border-[#bfe6e3]"
                  ):
                    ui.label(
                        "分析範圍只看「成品」跟「組合品」兩個分類（原料/物料/費用"
                        "類不是賣給客戶的品項，不計入）——不只是畫面篩選，連基準"
                        "銷量、去年比例回推這些底層計算都只用這個範圍，匯出的 "
                        "xlsx 也只會有這個範圍。按「計算」時會自動向 A1 抓取"
                        "「近3個月」與「去年同月」的真實銷售資料（不需要先去 5.3 "
                        "點、也不用整年硬抓，只拿這兩段還是需要幾秒到十幾秒，"
                        "不想自動抓可以在下方關掉）。邏輯：基準預估銷量 = (去年同期銷量 + "
                        "近3個月平均銷量) ÷ 2；若有填目標營業額，會等比例"
                        "校正每個品項的預估量，讓總營收貼近目標；建議採購"
                        "時間 = 目標月份第一天 − 採購前置天數（抓 BOM 表"
                        "設定，沒有則用系統預設值）。"
                    ).classes("text-xs text-teal-800")

                  with ui.row().classes("items-center gap-3 flex-wrap mb-2"):
                    forecast_month_select = ui.select(
                        options=generate_month_options(6),
                        value=generate_month_options(6)[0],
                        label="預估月份",
                    ).classes(
                        "bg-[#f7f6f2] text-zinc-900 rounded-none px-3 py-1"
                        " text-xs font-bold border border-[#e2e1dc] w-32"
                    )
                    forecast_revenue_input = ui.number(
                        label="目標營業額（選填，不填則用基準預估量）",
                        min=0,
                        step=1000,
                    ).classes("w-64")
                    forecast_last_year_revenue_input = ui.number(
                        label="去年目標營業額（選填，品號對不上時用比例回推）",
                        min=0,
                        step=1000,
                    ).classes("w-72")
                    forecast_auto_fetch_checkbox = ui.checkbox(
                        "自動從 A1 抓最新銷售資料（建議開啟）",
                        value=True,
                    ).classes("text-xs")
                    forecast_calc_button = ui.button("計算").classes(
                        "sync-btn px-4 py-2 text-xs rounded-none"
                    )

                  forecast_fetch_status_label = ui.label().classes(
                      "text-xs text-zinc-500 mb-1"
                  )
                  ui.label(
                      "⚠ 若去年的品號編碼跟今年不一樣，系統會完全比對不到"
                      "「去年同期銷量」（一律顯示為 0）。這時可以填「去年"
                      "目標營業額」，系統會用「近3個月」的營收佔比當權重，"
                      "把這個數字依同樣比例分攤回各品項、換算回數量，當作"
                      "去年同期的替代估計值——這是推算，不是實際數字，總表"
                      "會用「去年銷量來源」欄位標明。"
                  ).classes("text-xs text-zinc-500 mb-3")

                  ui.label("通路占比分配（選填，用於拆分下方「分通路表」）").classes(
                      "text-sm font-bold text-zinc-700 mb-2"
                  )
                  with ui.row().classes(
                      "w-full p-3 mb-2 bg-[#fff8e6] border border-[#f0dca0]"
                  ):
                    ui.label(
                        "不是逐品項拆通路（目前沒有這麼細的資料），而是把下方"
                        "「總表」算出來的預估總量/總成本/總營收，依你填的"
                        "占比直接等比例分攤到各通路，簡單好懂。不需要剛好等於"
                        "100%，但建議盡量接近，不然各通路加起來會跟總表對不起來。"
                    ).classes("text-xs text-amber-800")

                  channel_pct_inputs = {}
                  with ui.row().classes("w-full gap-3 flex-wrap mb-2"):
                    for channel in PRODUCTION_SALES_CHANNELS:
                      with ui.column().classes("gap-0"):
                        ui.label(channel).classes("text-xs text-zinc-600")
                        channel_pct_inputs[channel] = ui.number(
                            value=0, min=0, max=100, step=1,
                        ).classes("w-20")

                  channel_pct_total_label = ui.label().classes(
                      "text-xs text-zinc-500 mb-4"
                  )

                  def update_channel_pct_total():
                    total = sum(
                        (inp.value or 0) for inp in channel_pct_inputs.values()
                    )
                    channel_pct_total_label.text = f"目前合計：{total:g}%"
                    channel_pct_total_label.classes(
                        replace="text-xs mb-4 "
                        + ("text-teal-700" if abs(total - 100) < 0.01 else "text-amber-700")
                    )

                  for inp in channel_pct_inputs.values():
                    inp.on_value_change(lambda e: update_channel_pct_total())
                  update_channel_pct_total()

                  forecast_summary_row = ui.row().classes("w-full gap-4 mb-4 flex-wrap")
                  forecast_note_label = ui.label().classes(
                      "text-xs text-zinc-500 mb-3"
                  )
                  forecast_export_row = ui.row().classes("w-full mb-2")

                  ui.label("總表（逐品項）").classes(
                      "text-sm font-bold text-zinc-700 mb-2"
                  )
                  forecast_category_row = ui.row().classes("w-full gap-2 mb-3 flex-wrap")
                  forecast_table_container = ui.column().classes("w-full mb-6")

                  ui.label("分通路表（依占比等比例分攤）").classes(
                      "text-sm font-bold text-zinc-700 mb-2"
                  )
                  forecast_channel_table_container = ui.column().classes("w-full")

                  # 用來在按分類按鈕時重新篩選，不用重新整包計算一次
                  forecast_state = {"result": None, "active_category": "全部（成品＋組合品）", "channel_rows": []}

                  FORECAST_COLUMNS = [
                      {"name": "品號", "label": "品號", "field": "品號", "align": "left"},
                      {"name": "品名", "label": "品名", "field": "品名", "align": "left"},
                      {"name": "商品分類", "label": "商品分類", "field": "商品分類", "align": "left"},
                      {"name": "去年同期銷量", "label": "去年同期銷量", "field": "去年同期銷量"},
                      {"name": "去年銷量來源", "label": "去年銷量來源", "field": "去年銷量來源", "align": "left"},
                      {"name": "近3月平均銷量", "label": "近3月平均銷量", "field": "近3月平均銷量"},
                      {"name": "目標採購量", "label": "目標採購量", "field": "目標採購量"},
                      {"name": "單位成本", "label": "單位成本", "field": "單位成本"},
                      {"name": "預估總成本", "label": "預估總成本", "field": "預估總成本"},
                      {"name": "建議採購時間", "label": "建議採購時間", "field": "建議採購時間"},
                  ]
                  CHANNEL_COLUMNS = [
                      {"name": "通路", "label": "通路", "field": "通路", "align": "left"},
                      {"name": "佔比(%)", "label": "佔比(%)", "field": "佔比(%)"},
                      {"name": "目標營業額", "label": "目標營業額", "field": "目標營業額"},
                      {"name": "目標採購量", "label": "目標採購量", "field": "目標採購量"},
                      {"name": "預估總成本", "label": "預估總成本", "field": "預估總成本"},
                      {"name": "預估總營收", "label": "預估總營收", "field": "預估總營收"},
                  ]

                  def render_forecast_table():
                    forecast_table_container.clear()
                    result = forecast_state["result"]
                    if result is None:
                      return
                    rows = result["rows"]
                    active = forecast_state["active_category"]
                    if active == "成品":
                      rows = [
                          r for r in rows
                          if "成品" in (r["商品分類"] or "")
                      ]
                    elif active == "組合品":
                      rows = [
                          r for r in rows
                          if "組合品" in (r["商品分類"] or "")
                      ]
                    # active == "全部（成品＋組合品）" 時不額外篩選，因為
                    # result["rows"] 本來就已經只含這兩類（見計算階段的
                    # is_finished_or_combo_category 篩選）
                    with forecast_table_container:
                      if not rows:
                        ui.label("這個分類目前沒有預估資料").classes(
                            "text-xs text-zinc-400"
                        )
                      else:
                        ui.table(
                            columns=FORECAST_COLUMNS, rows=rows, pagination=10,
                        ).classes("w-full")

                  FORECAST_CATEGORY_OPTIONS = ["全部（成品＋組合品）", "成品", "組合品"]

                  def render_forecast_category_buttons():
                    forecast_category_row.clear()
                    result = forecast_state["result"]
                    if result is None:
                      return

                    def make_handler(cat):
                      def handler():
                        forecast_state["active_category"] = cat
                        render_forecast_category_buttons()
                        render_forecast_table()
                      return handler

                    with forecast_category_row:
                      for cat in FORECAST_CATEGORY_OPTIONS:
                        is_active = cat == forecast_state["active_category"]
                        ui.button(cat, on_click=make_handler(cat)).classes(
                            "px-3 py-1 text-xs rounded-none "
                            + (
                                "sync-btn"
                                if is_active
                                else "bg-[#f7f6f2] text-zinc-700 border border-[#e2e1dc]"
                            )
                        )

                  def render_channel_table():
                    forecast_channel_table_container.clear()
                    rows = forecast_state["channel_rows"]
                    with forecast_channel_table_container:
                      if not rows:
                        ui.label(
                            "尚未填寫通路佔比，或尚未按「計算」"
                        ).classes("text-xs text-zinc-400")
                      else:
                        ui.table(
                            columns=CHANNEL_COLUMNS, rows=rows, pagination=10,
                        ).classes("w-full")

                  def handle_calc_forecast():
                    forecast_summary_row.clear()
                    forecast_export_row.clear()
                    forecast_state["result"] = None
                    forecast_state["active_category"] = "全部（成品＋組合品）"
                    forecast_state["channel_rows"] = []

                    target_month = forecast_month_select.value

                    if forecast_auto_fetch_checkbox.value:
                      token = get_a1_token()
                      if token:
                        forecast_fetch_status_label.text = (
                            "正在向 A1 抓取近3個月與去年同月銷售資料，"
                            "需要幾秒到十幾秒…"
                        )
                        try:
                          fetched_rows = fetch_sales_history_for_forecast(
                              token, target_month
                          )
                        except Exception as e:
                          fetched_rows = []
                          ui.notify(f"自動抓取失敗，改用現有資料：{e}", color="warning")
                        if fetched_rows:
                          app_state["sales_history"] = fetched_rows
                          app_state["sales_history_configured"] = True
                          app_state["sales_history_source"] = (
                              "鼎新 A1（GetSales/GetSaleReturns，5.5 自動抓取）"
                          )
                          forecast_fetch_status_label.text = (
                              f"已從 A1 抓取 {len(fetched_rows)} 筆銷售歷史彙總資料"
                          )
                        else:
                          forecast_fetch_status_label.text = (
                              "A1 沒有查到資料，改用目前已有的銷售歷史"
                          )
                      else:
                        forecast_fetch_status_label.text = (
                            "無法登入 A1，改用目前已有的銷售歷史"
                        )
                    else:
                      forecast_fetch_status_label.text = "使用目前已有的銷售歷史（未勾選自動抓取）"

                    sales_history = app_state.get("sales_history", [])
                    if not app_state.get("sales_history_configured", False):
                      forecast_note_label.text = (
                          "尚未有銷售歷史資料來源，請先到 5.3 設定 Google "
                          "Sheets 或打開上方「自動從 A1 抓取」。"
                      )
                      forecast_category_row.clear()
                      forecast_table_container.clear()
                      render_channel_table()
                      return

                    items_map = app_state.get("items_map", {})
                    bom_map = app_state.get("bom_map", {})
                    target_revenue = forecast_revenue_input.value or 0
                    last_year_target_revenue = forecast_last_year_revenue_input.value or 0

                    result = compute_monthly_production_sales_forecast(
                        sales_history, items_map, bom_map, target_month,
                        target_revenue, app_state["settings"],
                        last_year_target_revenue=last_year_target_revenue,
                    )
                    forecast_state["result"] = result

                    forecast_note_label.text = (
                        f"參考去年同期（{result['last_year_ym']}）＋近3個月"
                        f"（{'、'.join(reversed(result['recent_months']))}）"
                        f"平均｜共 {len(result['rows'])} 項有預估值的品項"
                        + (
                            f"｜營收校正倍數：{result['scale_factor']}"
                            if target_revenue else "｜未填目標營業額，使用基準預估量"
                        )
                        + (
                            "｜去年同期已用比例推算補齊"
                            if last_year_target_revenue else ""
                        )
                    )

                    with forecast_summary_row:
                      summary_cards = [
                          ("目標採購總量", f"{result['total_est_qty']:,.0f}"),
                          ("預估總成本", f"NT$ {result['total_est_cost']:,.0f}"),
                          ("預估總營收（核對用）", f"NT$ {result['total_est_revenue']:,.0f}"),
                          (
                              "建議最早開始採購日",
                              result["earliest_order_date"] or "－",
                          ),
                      ]
                      for label, value in summary_cards:
                        with ui.column().classes(
                            "bg-[#f7f6f2] border border-[#e2e1dc] p-4"
                            " min-w-[200px] flex-1"
                        ):
                          ui.label(label).classes("text-xs text-zinc-500 mb-1")
                          ui.label(value).classes(
                              "text-xl font-black text-zinc-900"
                          )

                    channel_pcts = {
                        ch: (inp.value or 0)
                        for ch, inp in channel_pct_inputs.items()
                    }
                    channel_rows, total_pct = compute_channel_breakdown(
                        result, channel_pcts, target_revenue
                    )
                    forecast_state["channel_rows"] = channel_rows

                    if result["rows"]:
                      def handle_export_forecast():
                        try:
                          xlsx_bytes = multi_sheet_xlsx_bytes({
                              f"{target_month}總表": result["rows"],
                              f"{target_month}分通路表": channel_rows,
                          })
                          ui.download(
                              xlsx_bytes,
                              f"月產銷分析_{target_month}.xlsx",
                              media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                          )
                        except Exception as e:
                          ui.notify(f"匯出失敗：{e}", color="negative")

                      with forecast_export_row:
                        ui.button(
                            "匯出 xlsx（總表＋分通路表）",
                            on_click=handle_export_forecast,
                        ).classes("sync-btn px-3 py-1 text-xs rounded-none")

                    render_forecast_category_buttons()
                    render_forecast_table()
                    render_channel_table()

                  forecast_calc_button.on_click(handle_calc_forecast)


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
                    GOOGLE_SHEETS_CREDENTIALS_JSON and GOOGLE_SHEET_ID
                )
                ui.label(
                    "狀態："
                    + (
                        "已設定，四份資料（BOM表／訂單資訊／銷售歷史／"
                        "進貨明細）共用同一份 Google Sheet 讀取"
                        if sheets_configured
                        else "尚未設定。BOM表會退回本機 Excel 過渡方案；"
                        "訂單資訊、銷售歷史、進貨明細目前沒有備援來源，"
                        "相關頁面會顯示「尚未設定」"
                    )
                ).classes(
                    "text-sm mb-2 "
                    + ("text-teal-700" if sheets_configured else "text-amber-700")
                )
                with ui.column().classes("gap-1"):
                  ui.label(
                      f"Sheet ID：{GOOGLE_SHEET_ID or '（未設定）'}"
                  ).classes("text-xs text-zinc-600")
                  ui.label(
                      "服務帳號金鑰："
                      + ("已設定" if GOOGLE_SHEETS_CREDENTIALS_JSON else "（未設定）")
                  ).classes("text-xs text-zinc-600")
                  for label, tab_name, source_state in (
                      ("BOM表", BOM_GOOGLE_SHEET_TAB, app_state.get("bom_source")),
                      (
                          "訂單資訊", ORDERS_GOOGLE_SHEET_TAB,
                          "Google Sheets" if app_state.get("orders_configured") else "尚未設定",
                      ),
                      (
                          "銷售歷史", SALES_HISTORY_GOOGLE_SHEET_TAB,
                          app_state.get("sales_history_source", "尚未設定"),
                      ),
                      (
                          "進貨明細", RECEIVING_GOOGLE_SHEET_TAB,
                          "Google Sheets" if app_state.get("receivings_configured") else "尚未設定",
                      ),
                      (
                          "通路銷售明細", CHANNEL_SALES_GOOGLE_SHEET_TAB,
                          "Google Sheets" if app_state.get("channel_sales_configured") else "尚未設定",
                      ),
                  ):
                    ui.label(
                        f"分頁「{tab_name}」（{label}）目前來源：{source_state}"
                    ).classes("text-xs text-zinc-600")
                ui.label(
                    "設定方式：環境變數 GOOGLE_SHEETS_CREDENTIALS_JSON（服務"
                    "帳號金鑰 JSON 內容）、GOOGLE_SHEET_ID（五份資料共用同一"
                    "個 Sheet ID，只是分頁不同）。分頁名稱可用"
                    "BOM_GOOGLE_SHEET_TAB／ORDERS_GOOGLE_SHEET_TAB／"
                    "SALES_HISTORY_GOOGLE_SHEET_TAB／"
                    "RECEIVING_GOOGLE_SHEET_TAB／"
                    "CHANNEL_SALES_GOOGLE_SHEET_TAB 自訂，預設分別是"
                    "「BOM表」「訂單資訊」「銷售歷史」「進貨明細」"
                    "「通路銷售明細」。"
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

                with ui.row().classes("items-center gap-3 mb-3"):
                  ui.label("滯銷判定門檻（庫存週轉天數超過此值視為滯銷）").classes(
                      "text-sm text-zinc-700"
                  )
                  slow_moving_input = ui.number(
                      value=app_state["settings"]["slow_moving_days"],
                      min=7,
                      max=365,
                      step=1,
                  ).classes("w-24")

                def handle_save_settings():
                  app_state["settings"]["low_stock_alert_ratio"] = (
                      ratio_input.value or 1.0
                  )
                  app_state["settings"]["default_lead_time_days"] = (
                      lead_time_input.value or 0
                  )
                  app_state["settings"]["slow_moving_days"] = (
                      slow_moving_input.value or 90
                  )
                  ui.notify("參數已更新（僅本次服務執行期間有效）", color="positive")
                  for ref_key in (
                      "update_dashboard",
                      "update_procurement_list",
                      "update_packing_schedule",
                      "update_turnover_list",
                  ):
                    if ref_key in refs:
                      refs[ref_key]()

                ui.button("儲存參數", on_click=handle_save_settings).classes(
                    "sync-btn px-4 py-2 text-xs rounded-none"
                )


  # header（含公司切換分頁）已經搬到函式最前面建立，這裡只需要在所有
  # render_xxx 函式都定義好之後，觸發一次「畫面初始內容」即可。
  company_tabs.set_value(ACTIVE_COMPANY_LABEL)
  render_hai_tao_ke_page()


ui.run(port=8080, title="興聖集團 A1 智慧進銷存總管理系統", host="0.0.0.0")
