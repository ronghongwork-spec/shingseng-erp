import asyncio
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

from nicegui import app, run, ui
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


def compute_tax_and_total(subtotal, tax_type, tax_rate=0.05):
  """依A1的課稅別規則，從「品項金額合計」算出稅額跟總金額（含稅），供
  訂單/銷貨單/採購單/進貨單的手動填寫表單共用，避免四個地方各寫一次、
  容易改一個忘記改另一個。

  課稅別代碼統一用A1手冊的定義：
    0.免發票／3.免稅 → 不計稅，總金額=品項金額合計
    1.應稅外加 → 品項金額視為「未稅」，稅額另外加在總金額之上
    4.應稅內含 → 品項金額視為「已含稅」，總金額=品項金額合計，稅額用
      反推的方式從裡面拆出來
  稅率固定用5%（手冊範例皆用此稅率）。
  回傳 (稅額, 總金額)，皆四捨五入到小數點後2位。
  """
  if tax_type in ("0", "3"):
    return 0.0, round(subtotal, 2)
  if tax_type == "1":
    tax = round(subtotal * tax_rate, 2)
    return tax, round(subtotal + tax, 2)
  # "4" 應稅內含
  tax = round(subtotal - subtotal / (1 + tax_rate), 2)
  return tax, round(subtotal, 2)


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
# SHOPLINE 官網訂單串接（目前有3個據點，各自獨立帳密）
# 文件：https://open-api.docs.shoplineapp.com/docs/search-orders
# .env 需設定：
#   興聖(股)公司／官網(海濤客)：
#     SHOPLINE_XINGSHENG_ACCESS_TOKEN=
#     SHOPLINE_XINGSHENG_USER_AGENT=（已實測 "Xingsheng-ERP" 可用）
#   興聖(股)公司／官網(JDH)：
#     SHOPLINE_XINGSHENG_JDH_ACCESS_TOKEN=
#     SHOPLINE_XINGSHENG_JDH_USER_AGENT=
#   芙萊柏(股)公司／官網-B'f：
#     SHOPLINE_FULAIBO_ACCESS_TOKEN=
#     SHOPLINE_FULAIBO_USER_AGENT=
# -------------------------------------------------------------------------
SHOPLINE_API_DOMAIN = "https://open.shopline.io"
SHOPLINE_XINGSHENG_ACCESS_TOKEN = os.environ.get("SHOPLINE_XINGSHENG_ACCESS_TOKEN", "")
SHOPLINE_XINGSHENG_USER_AGENT = os.environ.get("SHOPLINE_XINGSHENG_USER_AGENT", "")
SHOPLINE_XINGSHENG_JDH_ACCESS_TOKEN = os.environ.get("SHOPLINE_XINGSHENG_JDH_ACCESS_TOKEN", "")
SHOPLINE_XINGSHENG_JDH_USER_AGENT = os.environ.get("SHOPLINE_XINGSHENG_JDH_USER_AGENT", "")
SHOPLINE_FULAIBO_ACCESS_TOKEN = os.environ.get("SHOPLINE_FULAIBO_ACCESS_TOKEN", "")
SHOPLINE_FULAIBO_USER_AGENT = os.environ.get("SHOPLINE_FULAIBO_USER_AGENT", "")


SHOPLINE_ORDER_STATUSES = ["pending", "confirmed"]  # 待處理+已確認
SHOPLINE_LOOKBACK_DAYS = 90  # 訂單建立時間往前推3個月

# -------------------------------------------------------------------------
# 分公司｜每日出貨（抓鼎新A1銷貨單，依銷貨單建立日期彙總品項數量＝揀貨表）
# 興聖/容鴻/芙萊柏在A1系統裡是各自獨立租戶，帳密跟海濤客的 A1_API_KEY /
# A1_API_PASSWORD 不是同一組，要各自設定。
# .env 需設定（目前只有興聖，其餘公司之後陸續補）：
#   A1_XINGSHENG_API_KEY=
#   A1_XINGSHENG_API_PASSWORD=
# -------------------------------------------------------------------------
A1_XINGSHENG_API_KEY = os.environ.get("A1_XINGSHENG_API_KEY", "")
A1_XINGSHENG_API_PASSWORD = os.environ.get("A1_XINGSHENG_API_PASSWORD", "")

# 每日出貨手動填寫的通路清單（跟出貨明細一起匯出，不會反查任何API）
DAILY_SHIPPING_CHANNELS = ["全家", "7-11", "黑貓", "新竹", "順豐", "海外", "其它"]

# -------------------------------------------------------------------------
# 分公司｜採購分析（建議採購量／庫存週轉／月產銷分析，共用A1帳密。
# 進貨明細目前跳過，因為A1沒有查詢API、要另外維護Google Sheet，容鴻/
# 芙萊柏是否有這份資料還待確認）
# .env 需設定（目前有容鴻、芙萊柏；興聖如果也要這個功能，可以直接沿用
# 上面 A1_XINGSHENG_API_KEY/PASSWORD，不用重複設定）：
#   A1_RONGHONG_API_KEY=
#   A1_RONGHONG_API_PASSWORD=
#   A1_FULAIBO_API_KEY=
#   A1_FULAIBO_API_PASSWORD=
# -------------------------------------------------------------------------
A1_RONGHONG_API_KEY = os.environ.get("A1_RONGHONG_API_KEY", "")
A1_RONGHONG_API_PASSWORD = os.environ.get("A1_RONGHONG_API_PASSWORD", "")
A1_FULAIBO_API_KEY = os.environ.get("A1_FULAIBO_API_KEY", "")
A1_FULAIBO_API_PASSWORD = os.environ.get("A1_FULAIBO_API_PASSWORD", "")
PROCUREMENT_ANALYSIS_LOOKBACK_MONTHS = 3  # 銷售歷史往前抓幾個月，用來算週轉/預估銷量

# 送貨方式關鍵字分類：先比對已知物流商關鍵字，再比對溫層關鍵字，
# 兩者獨立判斷後組合（例如「黑貓常溫」= 黑貓 + 常溫），沒對到已知物流商
# 關鍵字時，直接用原始送貨方式名稱當分組，避免資料被錯誤合併。
SHOPLINE_DELIVERY_CARRIER_KEYWORDS = [
    "黑貓", "全家", "7-11", "7-ELEVEN", "小七", "萊爾富", "OK超商", "郵局", "宅配通",
]
SHOPLINE_TEMP_KEYWORDS = [("常溫", "常溫"), ("冷藏", "冷藏"), ("低溫", "低溫"), ("冷凍", "冷凍")]


def classify_shopline_delivery_method(label):
  """依關鍵字判斷送貨方式分組。同一物流商不同溫層會分開（黑貓常溫／黑貓
  低溫不會合併），沒對到已知物流商關鍵字時直接用原始標籤當分組名稱。
  """
  label = (label or "").strip()
  if not label:
    return "未標示"
  carrier = next((kw for kw in SHOPLINE_DELIVERY_CARRIER_KEYWORDS if kw in label), None)
  temp = next((disp for kw, disp in SHOPLINE_TEMP_KEYWORDS if kw in label), None)
  base = carrier or label
  return f"{base}{temp}" if temp else base


def fetch_shopline_orders(access_token, user_agent, statuses, created_after):
  """抓SHOPLINE官網訂單原始資料（不做聚合），依statuses(list)+
  created_after（UTC時間字串"YYYY-MM-DD HH:MM:SS"）篩選，自動翻頁抓完。
  回傳 (orders, error_message)。
  """
  if not access_token or not user_agent:
    return None, "尚未設定 access_token / user_agent"
  try:
    all_orders = []
    page = 1
    while True:
      resp = requests.get(
          f"{SHOPLINE_API_DOMAIN}/v1/orders/search",
          params={
              "statuses[]": statuses,
              "created_after": created_after,
              "per_page": 50,
              "page": page,
          },
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
    return all_orders, None
  except Exception as e:
    return None, str(e)


def compute_shopline_stats(orders):
  """算總覽統計：待處理/已確認筆數（全部訂單，不受篩選影響），以及全部
  送貨方式名單（用來讓下拉選單的選項順序/項目固定，不會因為篩選而忽多
  忽少）。"""
  stats = {"pending": 0, "confirmed": 0, "delivery_counts": {}}
  for o in orders:
    st = o.get("status")
    if st in ("pending", "confirmed"):
      stats[st] += 1
    delivery_label = (
        (o.get("order_delivery") or {}).get("name_translations") or {}
    ).get("zh-hant", "")
    group = classify_shopline_delivery_method(delivery_label)
    stats["delivery_counts"][group] = stats["delivery_counts"].get(group, 0) + 1
  return stats


def compute_shopline_delivery_counts(orders, status_filter="all"):
  """算「送貨方式→筆數」，但只算status_filter篩選後的訂單。用來讓送貨
  方式下拉選單的數字跟著目前選擇的待處理/已確認/全部即時更新，例如選
  「待處理」時，各送貨方式的筆數加總起來會等於待處理總筆數。
  """
  if status_filter in ("pending", "confirmed"):
    filtered = [o for o in orders if o.get("status") == status_filter]
  else:
    filtered = orders
  counts = {}
  for o in filtered:
    label = (
        (o.get("order_delivery") or {}).get("name_translations") or {}
    ).get("zh-hant", "")
    group = classify_shopline_delivery_method(label)
    counts[group] = counts.get(group, 0) + 1
  return counts


def compute_shopline_sku_rows(
    orders, status_filter="all", delivery_filter="all", keyword="",
    date_from="", date_to="",
):
  """依status_filter(全部/待處理/已確認) + delivery_filter(全送貨方式/
  特定送貨方式) + keyword(商品名稱關鍵字) + date_from/date_to(訂單建立
  時間範圍，台灣時間，格式YYYY-MM-DD，任一留空代表不限) 篩選訂單後，依
  SKU加總商品需求數量。這是給海濤客食品工廠請備貨/採購用的主要清單。
  """
  filtered = orders
  if status_filter in ("pending", "confirmed"):
    filtered = [o for o in filtered if o.get("status") == status_filter]
  if delivery_filter and delivery_filter != "all":
    def _matches_delivery(o):
      label = (
          (o.get("order_delivery") or {}).get("name_translations") or {}
      ).get("zh-hant", "")
      return classify_shopline_delivery_method(label) == delivery_filter
    filtered = [o for o in filtered if _matches_delivery(o)]
  if date_from or date_to:
    def _order_date_tw(o):
      raw = o.get("created_at", "") or ""
      try:
        dt_utc = datetime.fromisoformat(raw.replace("Z", "+00:00"))
        return (dt_utc + timedelta(hours=8)).strftime("%Y-%m-%d")
      except (ValueError, TypeError):
        return ""
    def _in_range(o):
      d = _order_date_tw(o)
      if not d:
        return False
      if date_from and d < date_from:
        return False
      if date_to and d > date_to:
        return False
      return True
    filtered = [o for o in filtered if _in_range(o)]

  sku_agg = {}
  for o in filtered:
    for item in (o.get("subtotal_items") or []):
      for line in explode_shopline_item(item):
        entry = sku_agg.setdefault(
            line["SKU"],
            {"商品": line["商品"], "SKU": line["SKU"], "細項": line["細項"], "需求數量": 0},
        )
        # 細項理論上同一個SKU都是同一組固定內容，但保險起見：如果目前是
        # 空的、這次抓到有值，就補上去，避免因為抓到的第一筆剛好沒細項
        # 資料而漏掉。
        if not entry["細項"] and line["細項"]:
          entry["細項"] = line["細項"]
        entry["需求數量"] += line["數量"]

  rows = list(sku_agg.values())
  keyword = (keyword or "").strip()
  if keyword:
    rows = [r for r in rows if keyword in r["商品"]]
  rows.sort(key=lambda r: r["需求數量"], reverse=True)
  return rows


def _extract_zh_translation_list(fields_translations):
  """共用小工具：從SHOPLINE的xxx_translations格式(例如
  {"zh-hant": [...] 或 "zh-hant": "..."}) 撈出中文字串清單，統一格式、
  過濾空字串。extract_shopline_item_detail() 跟 explode_shopline_item()
  都會用到。
  """
  zh_list = (fields_translations or {}).get("zh-hant") or []
  if isinstance(zh_list, str):
    zh_list = [zh_list]
  return [s for s in zh_list if s]


def explode_shopline_item(item):
  """把單一 subtotal_item 展開成「實際要出貨/備貨的品項」清單，格式統一
  為 [{"SKU","商品","細項","數量"}, ...]：

  - 一般商品(item_type="Product"等)：展開成自己一筆，數量就是最外層
    quantity。
  - 客製化特惠組(item_type="ProductSet"，客人自己選內容的組合，最外層
    sku是空的)：拆解成 child_products 裡的每個子項目各一筆，數量＝
    item_data.selected_child_products 裡對應子項目的quantity，再乘上
    最外層quantity（訂了幾組這種特惠組本身），這樣才能算出實際要備多少
    貨，而不是把整組特惠組當成1個獨立品項。
  """
  if item.get("item_type") == "ProductSet":
    child_products = item.get("child_products") or []
    if not child_products:
      return []
    selected = (item.get("item_data") or {}).get("selected_child_products") or []
    qty_by_child_id = {s.get("child_product_id"): s.get("quantity") or 0 for s in selected}
    set_multiplier = item.get("quantity") or 1

    lines = []
    for child in child_products:
      child_id = child.get("id")
      qty_per_set = qty_by_child_id.get(child_id, 0)
      if qty_per_set <= 0:
        continue
      detail = "、".join(_extract_zh_translation_list(child.get("fields_translations")))
      sku = child.get("sku") or child_id or "(無SKU)"
      name = detail or sku
      lines.append({
          "SKU": sku,
          "商品": name,
          "細項": detail,
          "數量": qty_per_set * set_multiplier,
      })
    return lines

  sku = item.get("sku") or item.get("item_id") or "(無SKU)"
  title = item.get("title_translations") or {}
  name = title.get("zh-hant") or title.get("en") or sku
  qty = item.get("quantity") or 0
  detail = extract_shopline_item_detail(item)
  return [{"SKU": sku, "商品": name, "細項": detail, "數量": qty}]


def extract_shopline_item_detail(item):
  """從訂單商品明細(subtotal_items的單一item)裡撈出「細項」說明，支援
  兩種不同結構：

  1. 一般組合/禮盒商品(item_type="Product")：細項放在最外層
     item["fields_translations"]["zh-hant"]，字串陣列，例如
     「烏金醬*2 (禮盒組)」，沒有細項的一般商品這裡會是空字典{}。

  2. 客製化特惠組(item_type="ProductSet"，讓客人自選內容的組合，最外層
     sku是空的)：細項要從 item["child_products"][] 裡撈，每個子項目
     自己的 fields_translations.zh-hant 才是內容描述。這裡回傳的是給
     單一item摘要用的字串；實際彙總數量時 explode_shopline_item() 會
     分開處理每個子項目各自的數量，不會走這條路徑。
  """
  # 先試一般組合商品的結構（最外層）
  top_level = _extract_zh_translation_list(item.get("fields_translations"))
  if top_level:
    return "、".join(top_level)

  # 再試客製化特惠組(ProductSet)的結構
  child_products = item.get("child_products") or []
  if not child_products:
    return ""

  selected = (item.get("item_data") or {}).get("selected_child_products") or []
  qty_by_child_id = {s.get("child_product_id"): s.get("quantity") for s in selected}

  parts = []
  for child in child_products:
    names = _extract_zh_translation_list(child.get("fields_translations"))
    name = "、".join(names)
    if not name:
      continue
    qty = qty_by_child_id.get(child.get("id"))
    parts.append(f"{name}*{qty}" if qty else name)
  return "、".join(parts)


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
  return get_a1_token_for(API_KEY, API_PASSWORD)


def get_a1_token_for(api_key, api_password):
  """跟 get_a1_token() 同一套登入邏輯，但帳密用傳入的參數，而不是寫死
  海濤客的 API_KEY/API_PASSWORD，讓其他分公司（各自獨立的A1租戶）也能
  共用這支登入函式。
  """
  url = f"{A1_BASE_URL}/Login"
  headers = {"Content-Type": "application/json"}
  body = {"UserName": api_key, "Password": api_password}

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


def fetch_daily_shipping_items(api_key, api_password, start_date, end_date):
  """抓鼎新A1銷貨單(GetSales)，依「銷貨單建立日期」(TradeDate)在
  start_date~end_date區間內的資料，依品號加總數量，回傳揀貨表用的清單：
  [{"品號","品名","數量"}, ...]。
  回傳 (rows, error_message)；成功時 error_message 是 None。
  """
  if not api_key or not api_password:
    return None, "尚未設定 A1 API Key / Password"
  token = get_a1_token_for(api_key, api_password)
  if not token:
    return None, "A1登入失敗，請確認API Key/Password是否正確"

  try:
    qty_by_item = {}
    name_by_item = {}
    window_start = start_date
    while window_start <= end_date:
      window_end = min(window_start + timedelta(days=6), end_date)
      for sale in fetch_sales_details_range(
          token, window_start, window_end, "/Sales/PaginationQuery", "SaleDetails"
      ):
        for detail in sale.get("SaleDetails", []) or []:
          item_id = detail.get("ItemDetailID")
          if not item_id:
            continue
          try:
            qty = float(detail.get("Qty") or 0)
          except (TypeError, ValueError):
            qty = 0.0
          qty_by_item[item_id] = qty_by_item.get(item_id, 0) + qty
          if detail.get("ItemName"):
            name_by_item[item_id] = detail["ItemName"]
      window_start = window_end + timedelta(days=1)

    rows = [
        {"品號": item_id, "品名": name_by_item.get(item_id, ""), "數量": qty}
        for item_id, qty in qty_by_item.items()
    ]
    rows.sort(key=lambda r: r["數量"], reverse=True)
    return rows, None
  except Exception as e:
    return None, str(e)


def fetch_procurement_analysis_data(api_key, api_password, months_back=PROCUREMENT_ANALYSIS_LOOKBACK_MONTHS):
  """分公司採購分析的共用資料來源：登入A1後一次抓齊
    - items_map：商品主檔（品名/分類/成本/安全庫存等，來自 fetch_items_map）
    - stock_lookup：{品號: 現有庫存(跨倉庫加總)}
    - sales_history：近months_back個月的銷貨淨額，[{"年月","品號","品名",
      "銷售數量","銷售金額"}, ...]（銷貨-銷退，銷售金額用來反推平均售價）
  回傳 (data_dict, error_message)；成功時 error_message 是 None。
  """
  if not api_key or not api_password:
    return None, "尚未設定 A1 API Key / Password"
  token = get_a1_token_for(api_key, api_password)
  if not token:
    return None, "A1登入失敗，請確認API Key/Password是否正確"

  try:
    items_map = fetch_items_map(token)
    if not items_map:
      return None, "取得商品主檔失敗或商品清單為空"

    # ---- 庫存：StockBatch，依品號分批查詢，跨倉庫加總 ----
    stock_lookup = {}
    all_item_ids = list(items_map.keys())
    batch_size = 100  # 手冊：ItemIDs 一次最多可傳 100 筆
    batches = [all_item_ids[i:i + batch_size] for i in range(0, len(all_item_ids), batch_size)]
    for batch in batches:
      pagination = 1
      more = True
      while more and pagination <= MAX_STOCK_PAGES:
        resp = requests.post(
            f"{A1_BASE_URL}/Stock/Batch",
            json={"Pagination": pagination, "ItemIDs": batch},
            headers={"Content-Type": "application/json", "Authorization": token},
            timeout=REQUEST_TIMEOUT,
        )
        if resp.status_code != 200:
          break
        data = resp.json()
        rows = data if isinstance(data, list) else data.get("Data", [])
        for row in rows:
          iid = row.get("ItemID")
          if not iid:
            continue
          stock_lookup[iid] = stock_lookup.get(iid, 0) + (row.get("Qty") or 0)
        more = data.get("More", False) if isinstance(data, dict) else False
        pagination += 1

    # ---- 銷售歷史：近N個月銷貨-銷退，含金額(算平均售價用) ----
    end_date = datetime.now().date()
    start_date = end_date - timedelta(days=months_back * 30)
    qty_by_month_item = defaultdict(float)
    amount_by_month_item = defaultdict(float)
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
          qty = float(detail.get("Qty") or 0)
          amount = float(detail.get("Amount") or 0)
          qty_by_month_item[(year_month, item_id)] += qty
          amount_by_month_item[(year_month, item_id)] += amount
          if detail.get("ItemName"):
            name_by_item[item_id] = detail["ItemName"]
      for ret in fetch_sales_details_range(
          token, window_start, window_end, "/SaleReturns/PaginationQuery", "SaleReturnDetails"
      ):
        year_month = str(ret.get("TradeDate", ""))[:7].replace("/", "-")
        for detail in ret.get("SaleReturnDetails", []) or []:
          item_id = detail.get("ItemDetailID")
          if not item_id:
            continue
          qty = float(detail.get("Qty") or 0)
          amount = float(detail.get("Amount") or 0)
          qty_by_month_item[(year_month, item_id)] -= qty
          amount_by_month_item[(year_month, item_id)] -= amount
          if detail.get("ItemName"):
            name_by_item[item_id] = detail["ItemName"]
      window_start = window_end + timedelta(days=1)

    sales_history = [
        {
            "年月": year_month,
            "品號": item_id,
            "品名": items_map.get(item_id, {}).get("Name") or name_by_item.get(item_id, ""),
            "銷售數量": round(qty, 2),
            "銷售金額": round(amount_by_month_item.get((year_month, item_id), 0), 2),
        }
        for (year_month, item_id), qty in qty_by_month_item.items()
    ]

    return {
        "items_map": items_map,
        "stock_lookup": stock_lookup,
        "sales_history": sales_history,
    }, None
  except Exception as e:
    return None, str(e)


def compute_suggested_procurement(items_map, stock_lookup):
  """建議採購量（安全庫存基準）＝安全庫存(A1商品主檔設定) − 現有庫存，
  只列出「有設定安全庫存」且「算出來需要採購」的品項。沒設定安全庫存
  的品項沒有基準可比較，直接略過（不是庫存=0，是根本沒有這個判斷依據）。
  """
  rows = []
  for item_id, info in items_map.items():
    safety_stock = info.get("SafetyStock")
    if safety_stock is None:
      continue
    try:
      safety_stock = float(safety_stock)
    except (TypeError, ValueError):
      continue
    current_stock = stock_lookup.get(item_id, 0) or 0
    suggested_qty = safety_stock - current_stock
    if suggested_qty <= 0:
      continue
    unit_cost = info.get("StdPurPrice") or 0
    rows.append({
        "品號": item_id,
        "品名": info.get("Name", ""),
        "商品分類": info.get("CategoryName") or "未分類",
        "現有庫存": ceil_qty(current_stock),
        "安全庫存": ceil_qty(safety_stock),
        "建議採購量": ceil_qty(suggested_qty),
        "預估採購成本": round(suggested_qty * unit_cost, 2),
    })
  rows.sort(key=lambda r: r["建議採購量"], reverse=True)
  return rows


def compute_simple_monthly_forecast(items_map, sales_history, months_for_avg=PROCUREMENT_ANALYSIS_LOOKBACK_MONTHS):
  """簡化版月產銷分析：只分析「成品」跟「組合品」（見
  is_finished_or_combo_category），不做BOM展開成原物料——直接用商品
  本身的成本(StdPurPrice)和依銷售歷史反推的平均售價(銷售金額÷銷售數量)
  來估算，適合本身就有明確成本/售價、不需要拆解用料的商品。
  用近months_for_avg個月的平均銷量，估算「下個月大概要備多少貨」。
  """
  by_item = defaultdict(list)
  for row in sales_history:
    item_id = row["品號"]
    if not is_finished_or_combo_category(items_map.get(item_id, {}).get("CategoryName")):
      continue
    by_item[item_id].append(row)

  rows = []
  for item_id, records in by_item.items():
    records_sorted = sorted(records, key=lambda r: r["年月"])
    recent = records_sorted[-months_for_avg:]
    total_qty = sum(r["銷售數量"] for r in recent)
    total_amount = sum(r.get("銷售金額", 0) for r in recent)
    avg_qty = total_qty / len(recent) if recent else 0
    avg_price = (total_amount / total_qty) if total_qty > 0 else 0

    info = items_map.get(item_id, {})
    unit_cost = info.get("StdPurPrice") or 0
    est_qty = ceil_qty(avg_qty)

    rows.append({
        "品號": item_id,
        "品名": info.get("Name") or (records[-1]["品名"] if records else ""),
        "商品分類": info.get("CategoryName") or "未分類",
        f"近{months_for_avg}月平均銷量": est_qty,
        "平均售價": round(avg_price, 2),
        "商品成本": round(unit_cost, 2),
        "預估營收": round(est_qty * avg_price, 2),
        "預估成本": round(est_qty * unit_cost, 2),
        "預估毛利": round(est_qty * (avg_price - unit_cost), 2),
    })
  rows.sort(key=lambda r: r["預估營收"], reverse=True)
  return rows


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
        # 手冊的欄位說明表格寫「Details」，但手冊自己給的範例JSON卻是
        # 「OrderDetails」，兩者矛盾。實測發現用「Details」上傳會被A1
        # 判斷成空的品項清單（回傳400002「訂單單身不可空白」），改成
        # OrderDetails 才是實際有效的欄位名稱。
        "OrderDetails": details,
    }
    payloads.append((display_id, payload))

  return payloads, skipped


def upload_order_to_a1(token, payload):
  """Orders[Post]：上傳單張訂單。回傳 (成功與否, 訊息)。

  409（唯一辨識碼重複）視為「這張訂單先前已經上傳過」，不當成錯誤，讓
  呼叫端可以正常統計、不用擔心重複按會出亂子。
  """
  return _upload_document_to_a1(token, "/Orders", payload, "訂單")


def upload_sale_to_a1(token, payload):
  """Sales[Post]：上傳單張銷貨單。回傳 (成功與否, 訊息)。"""
  return _upload_document_to_a1(token, "/Sales", payload, "銷貨單")


def upload_purchase_to_a1(token, payload):
  """Purchases[Post]：上傳單張採購單。回傳 (成功與否, 訊息)。"""
  return _upload_document_to_a1(token, "/Purchases", payload, "採購單")


def upload_receive_to_a1(token, payload):
  """Receives[Post]：上傳單張進貨單。回傳 (成功與否, 訊息)。"""
  return _upload_document_to_a1(token, "/Receives", payload, "進貨單")


def _upload_document_to_a1(token, path, payload, doc_label):
  """訂單/銷貨單/採購單/進貨單上傳共用邏輯，差別只在API路徑跟錯誤訊息
  裡的單據名稱。409（唯一辨識碼重複）視為「這張單先前已經上傳過」，不
  當成錯誤，避免重複按會出亂子。
  """
  url = f"{A1_BASE_URL}{path}"
  headers = {"Content-Type": "application/json", "Authorization": token}
  try:
    response = requests.post(
        url, json=payload, headers=headers, timeout=REQUEST_TIMEOUT
    )
    if response.status_code == 200:
      return True, "上傳成功"
    if response.status_code == 409:
      return False, f"{doc_label}編號重複（先前應該已經上傳過，正常現象）"
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


def _get_gspread_client():
  """建立可讀寫的 gspread client（共用給讀取跟寫入用）。
  回傳 None 代表沒有設定 Google Sheets 憑證。
  """
  if not GOOGLE_SHEETS_CREDENTIALS_JSON or not GOOGLE_SHEET_ID:
    return None
  import gspread
  from google.oauth2.service_account import Credentials

  creds_info = json.loads(GOOGLE_SHEETS_CREDENTIALS_JSON)
  creds = Credentials.from_service_account_info(
      creds_info,
      # 用完整的 spreadsheets 權限（含寫入），不是唯讀，這樣同一組憑證
      # 讀取跟「手動建立訂單」要寫回Sheet兩邊都能共用，不用維護兩種
      # scope的client。
      scopes=["https://www.googleapis.com/auth/spreadsheets"],
  )
  return gspread.authorize(creds)


def _fetch_google_sheet_records(tab_name):
  """共用的 Google Sheet 讀取邏輯（BOM／訂單資訊／銷售歷史都靠這支）。

  回傳 None 代表「沒有設定 Google Sheets」（呼叫端應退回其他備援來源）；
  回傳空 list 代表「有設定，但讀取失敗，或該分頁本來就沒有資料」。
  這兩種情況要分開，才不會把「沒設定」誤判成「設定了但是空的」。

  設定步驟：
  1. Google Cloud Console 建立服務帳號，下載 JSON 金鑰
  2. 把金鑰 JSON 的完整內容存進環境變數 GOOGLE_SHEETS_CREDENTIALS_JSON
  3. 把該服務帳號的 email（金鑰 JSON 裡的 client_email）加入 Google
     Sheet 的「共用」名單，權限「編輯者」（不是唯讀，因為手動建立訂單
     功能需要寫回這份Sheet）
  4. 設定 GOOGLE_SHEET_ID（Sheet 網址中 /d/ 與 /edit 中間那串）
  5. 分頁名稱要跟 BOM_GOOGLE_SHEET_TAB／ORDERS_GOOGLE_SHEET_TAB／
     SALES_HISTORY_GOOGLE_SHEET_TAB 一致，欄位標題要跟本檔案裡的
     BOM_COL_* / ORDER_COL_* / SALES_COL_* 常數完全一致
  6. requirements.txt 需要加上 gspread、google-auth
  """
  gc = _get_gspread_client()
  if gc is None:
    return None

  try:
    sh = gc.open_by_key(GOOGLE_SHEET_ID)
    ws = sh.worksheet(tab_name)
    return ws.get_all_records()
  except Exception as e:
    print(f"讀取 Google Sheet「{tab_name}」分頁失敗：{e}")
    return []


def append_rows_to_google_sheet(tab_name, rows):
  """把 rows（list[dict]，key要對到Sheet的欄位標題）依照該分頁目前的
  欄位順序，逐列附加到分頁最後面。用來讓「手動建立訂單」等功能可以把
  新資料直接寫回Sheet，跟既有的訂單資訊共用同一份，不用另外維護一份
  紀錄。

  回傳 (成功筆數, 錯誤訊息)；成功時錯誤訊息是 None。
  沒設定Google Sheets憑證時，回傳 (0, "尚未設定Google Sheets")。
  """
  gc = _get_gspread_client()
  if gc is None:
    return 0, "尚未設定 Google Sheets"
  if not rows:
    return 0, None

  try:
    sh = gc.open_by_key(GOOGLE_SHEET_ID)
    ws = sh.worksheet(tab_name)
    header = ws.row_values(1)
    if not header:
      return 0, f"「{tab_name}」分頁沒有標題列，無法對應欄位順序"

    values = [[str(row.get(col, "")) for col in header] for row in rows]
    ws.append_rows(values, value_input_option="USER_ENTERED")
    return len(values), None
  except Exception as e:
    return 0, str(e)


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

  # 訂單出貨提醒：同一張訂單如果有好幾個品項，原始資料(orders_in_horizon)
  # 會是一個品項一列，這裡先依「訂單編號」分組，同一張訂單只顯示一行
  # （取最早的預計出貨日），不逐項列出品名/數量明細，畫面才不會被拆成
  # 一堆瑣碎的品項提醒。沒有訂單編號的品項(用品號當識別碼)不會被誤併，
  # 分組鍵固定用「訂單編號 or 品號」跟原本order_label邏輯一致。
  orders_by_key = {}
  for o in result["orders_in_horizon"]:
    key = o["訂單編號"] or o["品號"]
    if key not in orders_by_key or o["預計出貨日"] < orders_by_key[key]["預計出貨日"]:
      orders_by_key[key] = o

  shipping = []
  for key, o in sorted(orders_by_key.items(), key=lambda kv: kv[1]["預計出貨日"]):
    days_left = (o["預計出貨日"] - today).days
    severity = "danger" if days_left <= 1 else "warning" if days_left <= 3 else "info"
    shipping.append({
        "text": f"訂單 {key} 需於 {o['預計出貨日'].isoformat()} 出貨",
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
            f"生產組裝確認：{r.get('品名') or item_id}，原料已備妥，建議"
            f"安排組裝／生產（最早出貨日 {r['最早出貨日']}）"
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

  只分析「成品」跟「組合品」兩個分類（見 is_finished_or_combo_category），
  原料/物料/費用類不是賣給客戶的品項，週轉率對這些品項沒有意義，排除
  掉才不會讓滯銷清單充滿一堆原物料雜訊。
  """
  by_item = defaultdict(list)
  for row in sales_history:
    item_id = row["品號"]
    if not is_finished_or_combo_category(items_map.get(item_id, {}).get("CategoryName")):
      continue
    by_item[item_id].append(row)

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


def render_section_placeholder(title, hint="此區尚未串接資料來源，敬請期待"):
  """訂單出貨／每日出貨／調撥紀錄／退換貨記錄 共用的「還沒串API」佔位畫面。
  等之後陸續串接各分公司/各通路的 API 時，把對應區塊換成真的資料表格即可，
  版面（分頁結構）不用重搭。
  """
  with ui.card().classes(
      "w-full p-10 bg-white border border-[#e6e1d4] shadow-[0_1px_3px_rgba(42,40,35,0.06)]"
      " rounded-lg text-center"
  ):
    ui.label(title).classes("text-sm font-bold text-zinc-700 mb-2")
    ui.label(hint).classes("text-xs text-zinc-500")

async def render_shopline_channel(access_token, user_agent, channel_title, restock_target=None):
  """SHOPLINE官網訂單通路的共用畫面（興聖官網(海濤客)／官網(JDH)／
  芙萊柏官網-B'f 都呼叫這支，只是傳入的access_token/user_agent/標題不同）。
  抓近3個月「待處理」+「已確認」訂單，畫面：
    1. 篩選工具列（放在最上方，一進分頁就看得到）：狀態(全部/待處理/
       已確認，各自帶筆數)、送貨方式下拉選單(全送貨方式+各送貨方式，
       數字會依目前選的狀態即時重算)、商品關鍵字搜尋、匯出xlsx
    2. 商品需求彙總表：依目前篩選條件即時算出的SKU加總結果，商品名稱
       欄位可排序，右下角可切換每頁顯示筆數(10/30/50/全部)，不用一直
       往下滑。所有篩選都是在「已經抓好的資料」裡做，不會重打API。
  每次「切換到這個分頁」都會重新打一次API抓最新資料；分頁內也有「重新
  整理」按鈕，不用離開分頁再切回來也能手動重抓一次。

  restock_target：這個通路的商品要向誰請備貨/採購，會顯示在商品需求
  彙總的標題裡（不同通路賣的商品可能來自不同工廠/公司，備貨對象不一定
  跟通路本身掛在哪個公司頁籤一樣）。沒傳的話預設用channel_title。

  這支是async函式，實際打API的地方都用 run.io_bound() 包起來，讓抓資料
  這段「同步阻塞」的過程丟到背景執行緒跑，不會卡住整個伺服器的事件
  迴圈──不然一個人切到這個分頁在等API回應時，全部人的畫面都會跟著
  卡住沒反應。
  """
  restock_target = restock_target or channel_title
  def fetch_data():
    created_after = (
        datetime.utcnow() - timedelta(days=SHOPLINE_LOOKBACK_DAYS)
    ).strftime("%Y-%m-%d %H:%M:%S")
    return fetch_shopline_orders(
        access_token, user_agent, SHOPLINE_ORDER_STATUSES, created_after,
    )

  orders, error = await run.io_bound(fetch_data)
  if error:
    render_section_placeholder(f"訂單出貨－{channel_title}", f"抓取失敗：{error}")
    return
  if not orders:
    render_section_placeholder(
        f"訂單出貨－{channel_title}", "近3個月沒有待處理或已確認的訂單"
    )
    return

  stats = compute_shopline_stats(orders)
  last_updated = (datetime.utcnow() + timedelta(hours=8)).strftime("%Y-%m-%d %H:%M:%S")
  state = {
      "status_filter": "all", "delivery_filter": "all", "keyword": "",
      "date_from": "", "date_to": "",
  }

  # ---- 更新時間 + 手動重新整理按鈕 ----
  refresh_row = ui.row().classes("w-full items-center gap-3 mb-2")

  def render_refresh_row():
    nonlocal last_updated
    refresh_row.clear()
    with refresh_row:
      ui.label(f"資料更新時間：{last_updated}").classes("text-xs text-zinc-400")

      async def handle_refresh():
        nonlocal orders, stats, last_updated
        new_orders, err = await run.io_bound(fetch_data)
        if err:
          ui.notify(f"重新整理失敗：{err}", color="negative")
          return
        orders = new_orders or []
        stats = compute_shopline_stats(orders)
        last_updated = (datetime.utcnow() + timedelta(hours=8)).strftime("%Y-%m-%d %H:%M:%S")
        render_refresh_row()
        render_toolbar()
        refresh_results()
        ui.notify("已重新整理", color="positive")

      ui.button("重新整理", icon="refresh", on_click=handle_refresh).props(
          "dense no-caps unelevated"
      ).classes("px-3 py-1 rounded-lg text-xs").style(
          "background:#ffffff !important; color:#4b5563 !important;"
          " border:1px solid #e6e1d4;"
      )

  render_refresh_row()

  # ---- 篩選工具列：放在最上方（在說明文字跟結果表格之前）----
  toolbar_container = ui.row().classes("w-full items-end gap-3 mb-1 flex-wrap")

  def render_toolbar():
    toolbar_container.clear()
    with toolbar_container:
      status_options = [
          ("全部", "all"),
          (f"待處理 ({stats['pending']})", "pending"),
          (f"已確認 ({stats['confirmed']})", "confirmed"),
      ]
      with ui.row().classes("gap-2"):
        for label, value in status_options:
          is_active = value == state["status_filter"]
          bg = "#5bc0be" if is_active else "#ffffff"
          fg = "#ffffff" if is_active else "#4b5563"
          border = "#5bc0be" if is_active else "#e6e1d4"
          # 用行內 style 強制指定顏色，避免被 NiceGUI/Quasar 按鈕預設的
          # 文字顏色蓋掉（Tailwind的class在這裡權重會輸給Quasar內建樣式，
          # 導致文字顏色跟背景一樣、整個看起來像空白按鈕）。
          ui.button(
              label, on_click=lambda v=value: set_status_filter(v)
          ).props("dense no-caps unelevated").classes(
              "px-4 py-1 rounded-lg"
          ).style(
              f"background:{bg} !important; color:{fg} !important;"
              f" border:1px solid {border};"
          )

      delivery_counts_now = compute_shopline_delivery_counts(
          orders, state["status_filter"]
      )
      total_now = sum(delivery_counts_now.values())
      # 選項名單固定用全部送貨方式（避免篩選後某個方式暫時是0筆就從
      # 選單裡消失），但數字用目前狀態篩選後的筆數，兩者分開處理
      delivery_select_options = {"all": f"全送貨方式 ({total_now})"}
      for method in sorted(stats["delivery_counts"].keys()):
        cnt = delivery_counts_now.get(method, 0)
        delivery_select_options[method] = f"{method} ({cnt})"
      ui.select(
          options=delivery_select_options,
          value=state["delivery_filter"],
          on_change=set_delivery_filter,
          label="送貨方式",
      ).props("dense outlined").classes("w-48")

      ui.input(
          label="搜尋商品關鍵字",
          value=state["keyword"],
          on_change=set_keyword,
      ).props("dense outlined clearable").classes("w-56")

      ui.input(
          label="建立日期 起",
          value=state["date_from"],
          on_change=set_date_from,
      ).props('dense outlined clearable type="date"').classes("w-40")

      ui.input(
          label="建立日期 迄",
          value=state["date_to"],
          on_change=set_date_to,
      ).props('dense outlined clearable type="date"').classes("w-40")

  ui.label(
      f"{channel_title}・近{SHOPLINE_LOOKBACK_DAYS}天訂單建立時間"
  ).classes("text-xs text-zinc-500 mb-3")

  results_container = ui.column().classes("w-full")

  def refresh_results():
    results_container.clear()
    rows = compute_shopline_sku_rows(
        orders, state["status_filter"], state["delivery_filter"], state["keyword"],
        state["date_from"], state["date_to"],
    )
    with results_container:
      with ui.card().classes(
          "w-full p-6 bg-white border border-[#e6e1d4] shadow-[0_1px_3px_rgba(42,40,35,0.06)] rounded-lg"
      ):
        with ui.row().classes("w-full items-center justify-between mb-3"):
          ui.label(
              f"未出貨商品需求彙總（共 {len(rows)} 個品項，"
              f"供向{restock_target}請備貨/採購用）"
          ).classes("text-sm font-bold text-zinc-700")

          def handle_export():
            try:
              xlsx_bytes = rows_to_xlsx_bytes(rows, sheet_name="商品需求彙總")
              ui.download(
                  xlsx_bytes,
                  f"{channel_title}商品需求彙總.xlsx",
                  media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              )
            except Exception as e:
              ui.notify(f"匯出失敗：{e}", color="negative")

          ui.button("匯出 xlsx", on_click=handle_export).classes(
              "sync-btn px-3 py-1 text-xs rounded-lg"
          )

        if not rows:
          ui.label("目前沒有符合篩選條件的品項").classes("text-xs text-zinc-400")
        else:
          # rows-per-page-options 用 Quasar 慣例：0 代表「全部」，
          # 表格右下角會有內建的每頁筆數切換選單(10/30/50/全部)。
          ui.table(
              columns=[
                  {"name": "SKU", "label": "SKU", "field": "SKU", "align": "left", "sortable": True},
                  {"name": "商品", "label": "商品", "field": "商品", "align": "left", "sortable": True},
                  {"name": "細項", "label": "細項", "field": "細項", "align": "left", "sortable": True},
                  {"name": "需求數量", "label": "需求數量", "field": "需求數量", "align": "right", "sortable": True},
              ],
              rows=rows,
              row_key="SKU",
              pagination={"rowsPerPage": 10, "sortBy": "需求數量", "descending": True},
          ).classes("w-full").props(':rows-per-page-options="[10,30,50,0]"')

  def set_status_filter(v):
    state["status_filter"] = v
    render_toolbar()
    refresh_results()

  def set_delivery_filter(e):
    state["delivery_filter"] = e.value
    refresh_results()

  def set_keyword(e):
    state["keyword"] = e.value
    refresh_results()

  def set_date_from(e):
    state["date_from"] = e.value or ""
    refresh_results()

  def set_date_to(e):
    state["date_to"] = e.value or ""
    refresh_results()

  render_toolbar()
  refresh_results()



# 訂單出貨底下的通路子分頁，之後每個通路會各自串不同的訂單來源API。
# 各分公司實際通路不完全一樣，用字典各自設定；沒特別列出的公司預設
# 用完整4通路。
ORDER_CHANNELS_BY_COMPANY = {
    # 興聖沒有蝦皮通路，官網部分有兩個獨立SHOPLINE據點：海濤客品牌、JDH
    "興聖(股)公司": ["官網(海濤客)", "官網(JDH)", "經銷", "其它"],
    # 芙萊柏的官網通路標籤用他們自己的代稱
    "芙萊柏(股)公司": ["官網-B'f", "蝦皮", "經銷", "其它"],
}
DEFAULT_ORDER_CHANNELS = ["SHOPLINE官網", "蝦皮", "經銷", "其它"]

# (公司, 通路標籤) -> 呼叫render_shopline_channel()要用的(access_token, user_agent)
SHOPLINE_CHANNEL_CREDENTIALS = {
    ("興聖(股)公司", "官網(海濤客)"): (SHOPLINE_XINGSHENG_ACCESS_TOKEN, SHOPLINE_XINGSHENG_USER_AGENT),
    ("興聖(股)公司", "官網(JDH)"): (SHOPLINE_XINGSHENG_JDH_ACCESS_TOKEN, SHOPLINE_XINGSHENG_JDH_USER_AGENT),
    ("芙萊柏(股)公司", "官網-B'f"): (SHOPLINE_FULAIBO_ACCESS_TOKEN, SHOPLINE_FULAIBO_USER_AGENT),
}

# (公司, 通路標籤) -> 這個通路的商品需求彙總要向誰請備貨/採購（不同
# 通路賣的商品可能來自不同工廠/公司，備貨對象不一定跟頁籤本身的公司
# 一樣，例如興聖官網(JDH)是要跟容鴻請備貨）。沒列出的通路，預設用該
# 頁籤所屬的公司名稱。
SHOPLINE_RESTOCK_TARGET = {
    ("興聖(股)公司", "官網(海濤客)"): "海濤客食品工廠",
    ("興聖(股)公司", "官網(JDH)"): "容鴻(股)公司",
}

# 公司 -> 每日出貨要用的A1帳密(api_key, api_password)。目前只有興聖，
# 容鴻/芙萊柏之後拿到帳密再補進來即可，沒列出的公司會顯示佔位畫面。
DAILY_SHIPPING_CREDENTIALS = {
    "興聖(股)公司": (A1_XINGSHENG_API_KEY, A1_XINGSHENG_API_PASSWORD),
}

# 公司 -> 採購分析要用的A1帳密(api_key, api_password)。目前是容鴻、
# 芙萊柏；興聖如果之後也要這個功能，可以直接沿用A1_XINGSHENG_API_KEY/
# PASSWORD，把這行也加進來即可。
PROCUREMENT_ANALYSIS_CREDENTIALS = {
    "容鴻(股)公司": (A1_RONGHONG_API_KEY, A1_RONGHONG_API_PASSWORD),
    "芙萊柏(股)公司": (A1_FULAIBO_API_KEY, A1_FULAIBO_API_PASSWORD),
}

# 公司名稱轉成安全的英文代碼，用來組CSS class名稱（中文當class名稱在
# 部分瀏覽器/選擇器語法下容易出錯，改用英文代碼比較保險）
COMPANY_SLUGS = {
    "興聖(股)公司": "xingsheng",
    "容鴻(股)公司": "ronghong",
    "芙萊柏(股)公司": "fulaibo",
}

async def render_daily_shipping(api_key, api_password, company_label):
  """分公司／每日出貨：抓鼎新A1銷貨單，依「銷貨單建立日期」區間彙總
  品項數量＝揀貨表；下方另外提供一張「通路分類數量」讓人員手動填寫
  （全家/7-11/黑貓/新竹/順豐/海外/其它，純手填，不會反查任何系統），
  兩張表最後可以合併匯出成一份xlsx（兩個分頁）。

  這支是async函式，實際打A1 API的地方用run.io_bound()包起來，避免
  卡住整個伺服器的事件迴圈。
  """
  today_tw = (datetime.utcnow() + timedelta(hours=8)).date()
  state = {"date_from": today_tw.isoformat(), "date_to": today_tw.isoformat()}
  picking_rows_holder = {"rows": []}
  channel_inputs = {}

  ui.label(f"{company_label}・每日出貨（依銷貨單建立日期彙總）").classes(
      "text-xs text-zinc-500 mb-3"
  )

  date_row = ui.row().classes("w-full items-end gap-3 mb-4 flex-wrap")

  results_container = ui.column().classes("w-full")

  async def load_and_render():
    picking_container.clear()
    try:
      d_from = datetime.strptime(state["date_from"], "%Y-%m-%d").date()
      d_to = datetime.strptime(state["date_to"], "%Y-%m-%d").date()
    except (ValueError, TypeError):
      with picking_container:
        ui.label("日期格式錯誤，請重新選擇").classes("text-xs text-red-500")
      return
    if d_from > d_to:
      with picking_container:
        ui.label("「起」不能晚於「迄」，請重新選擇").classes("text-xs text-red-500")
      return

    with picking_container:
      with ui.row().classes("items-center gap-2 p-4"):
        ui.spinner(size="20px").classes("text-zinc-400")
        ui.label("抓取中…").classes("text-xs text-zinc-500")

    rows, error = await run.io_bound(
        fetch_daily_shipping_items, api_key, api_password, d_from, d_to
    )
    picking_container.clear()
    if error:
      with picking_container:
        ui.label(f"抓取失敗：{error}").classes("text-xs text-red-500")
      return
    picking_rows_holder["rows"] = rows or []
    with picking_container:
      if not rows:
        ui.label("此區間沒有銷貨單資料").classes("text-xs text-zinc-400")
      else:
        ui.table(
            columns=[
                {"name": "品號", "label": "品號", "field": "品號", "align": "left", "sortable": True},
                {"name": "品名", "label": "品名", "field": "品名", "align": "left", "sortable": True},
                {"name": "數量", "label": "數量", "field": "數量", "align": "right", "sortable": True},
            ],
            rows=rows,
            row_key="品號",
            pagination={"rowsPerPage": 10, "sortBy": "數量", "descending": True},
        ).classes("w-full").props(':rows-per-page-options="[10,30,50,0]"')

  async def set_date_from(e):
    state["date_from"] = e.value or ""
    await load_and_render()

  async def set_date_to(e):
    state["date_to"] = e.value or ""
    await load_and_render()

  with date_row:
    ui.input(
        label="銷貨單建立日期 起", value=state["date_from"], on_change=set_date_from,
    ).props('dense outlined type="date"').classes("w-44")
    ui.input(
        label="銷貨單建立日期 迄", value=state["date_to"], on_change=set_date_to,
    ).props('dense outlined type="date"').classes("w-44")

  with results_container:
    with ui.card().classes(
        "w-full p-6 bg-white border border-[#e6e1d4] shadow-[0_1px_3px_rgba(42,40,35,0.06)] rounded-lg mb-4"
    ):
      ui.label("揀貨表（依品號加總數量）").classes("text-sm font-bold text-zinc-700 mb-3")
      picking_container = ui.column().classes("w-full")

    with ui.card().classes(
        "w-full p-6 bg-white border border-[#e6e1d4] shadow-[0_1px_3px_rgba(42,40,35,0.06)] rounded-lg"
    ):
      with ui.row().classes("w-full items-center justify-between mb-3"):
        ui.label("通路分類數量（人員手動填寫，不會自動計算）").classes(
            "text-sm font-bold text-zinc-700"
        )

        def handle_export():
          try:
            channel_rows = [
                {"通路": ch, "數量": int(inp.value or 0)}
                for ch, inp in channel_inputs.items()
            ]
            xlsx_bytes = multi_sheet_xlsx_bytes({
                "揀貨表": picking_rows_holder["rows"],
                "通路分類數量": channel_rows,
            })
            ui.download(
                xlsx_bytes,
                f"{company_label}每日出貨_{state['date_from']}_{state['date_to']}.xlsx",
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
          except Exception as e:
            ui.notify(f"匯出失敗：{e}", color="negative")

        ui.button("匯出 xlsx（揀貨表＋通路分類數量）", on_click=handle_export).classes(
            "sync-btn px-3 py-1 text-xs rounded-lg"
        )

      with ui.row().classes("w-full gap-3 flex-wrap"):
        for ch in DAILY_SHIPPING_CHANNELS:
          with ui.column().classes("gap-1"):
            ui.label(ch).classes("text-xs text-zinc-500")
            channel_inputs[ch] = ui.input(value="0").props(
                'dense outlined type="number"'
            ).classes("w-24")

  await load_and_render()


def render_order_channels_tabs(company_name):
  """「訂單出貨」分頁底下的通路子分頁（SHOPLINE官網／蝦皮／經銷／其它
  等，依公司不同而不同）。一樣做成懶載入：只有實際點進某個通路，才會
  去打那個通路的API，不會切到「訂單出貨」就把底下所有通路一次全部
  打完。

  注意：分頁列(ui.tabs)一定要在內容容器(channel_body)「之前」建立，
  這樣分頁列在畫面排列順序上才會排在內容前面（天生在最上方）。順序
  反過來的話，分頁列雖然視覺上看起來應該在上面，但因為content容器的
  位置已經先佔走了，分頁列反而會被排到內容下方——這是之前修過的同一
  種bug，這裡也要注意。
  """
  order_channels = ORDER_CHANNELS_BY_COMPANY.get(
      company_name, DEFAULT_ORDER_CHANNELS
  )

  async def handle_channel_change(ch):
    channel_body.clear()
    with channel_body:
      with ui.row().classes("w-full items-center gap-2 p-8 justify-center"):
        ui.spinner(size="24px").classes("text-zinc-400")
        ui.label("資料抓取中，請稍候…").classes("text-xs text-zinc-500")
    await asyncio.sleep(0)

    channel_body.clear()
    shopline_creds = SHOPLINE_CHANNEL_CREDENTIALS.get((company_name, ch))
    restock_target = SHOPLINE_RESTOCK_TARGET.get((company_name, ch), company_name)
    with channel_body:
      if shopline_creds:
        await render_shopline_channel(
            shopline_creds[0], shopline_creds[1], ch, restock_target
        )
      else:
        render_section_placeholder(
            f"訂單出貨－{ch}",
            f"「{ch}」通路的訂單 API 尚未串接，敬請期待",
        )

  with ui.tabs(on_change=lambda e: handle_channel_change(e.value)).props(
      "dense no-caps"
  ).classes("w-full") as channel_tabs:
    for ch in order_channels:
      ui.tab(ch)

  channel_body = ui.column().classes("w-full")
  channel_tabs.set_value(order_channels[0])


COMPANIES = ["興聖(股)公司", "海濤客食品工業(股)公司", "容鴻(股)公司", "芙萊柏(股)公司"]
ACTIVE_COMPANY_LABEL = "海濤客食品工業(股)公司"

COMPANY_TAB_COLORS = {
    "興聖(股)公司": {"text": "#5bc0be", "active_bg": "#5bc0be"},
    "海濤客食品工業(股)公司": {"text": "#e0824a", "active_bg": "#e0824a"},
    "容鴻(股)公司": {"text": "#8e7cc3", "active_bg": "#8e7cc3"},
    "芙萊柏(股)公司": {"text": "#5b8fc0", "active_bg": "#5b8fc0"},
}

# -------------------------------------------------------------------------
# App 切換器（雲端進銷存／雲端電商訂單／雲端會計／報表分析 四個獨立頁面）
# 每個App是一個獨立的 @ui.page 路由，共用同一個Render服務、同一組
# Basic Auth登入，不用重複登入。之後陸續把內容搬過去對應的App時，這份
# 清單也要記得同步更新標籤名稱/路徑。
# -------------------------------------------------------------------------
APP_SWITCHER_ITEMS = [
    ("雲端進銷存", "/"),
    ("雲端電商訂單", "/orders"),
    ("雲端會計", "/accounting"),
    ("報表分析", "/analytics"),
]


def render_app_switcher(active_path):
  """四個App之間的切換器，放在畫面最上方。用ui.link做頁面跳轉（不是
  NiceGUI的tab_panels切換，因為這是四個「不同網址」的獨立頁面，要真的
  換頁，不是同一頁裡切換內容）。
  """
  with ui.row().classes(
      "w-full items-center gap-1 bg-[#f7f5ef] border-b border-[#e6e1d4]"
      " px-8 py-2 sticky top-0 z-[60]"
  ):
    for label, path in APP_SWITCHER_ITEMS:
      is_active = path == active_path
      ui.link(label, path).classes(
          "px-3 py-1 text-xs no-underline rounded-lg "
          + (
              "bg-[#2a2823] text-white font-bold"
              if is_active
              else "text-zinc-600 hover:bg-[#ece6d6]"
          )
      )


def inject_global_theme_css():
  """全站共用的字型/配色/元件樣式（暖石色底、襯線標題、柔和圓角+陰影）。
  四個App頁面（進銷存／電商訂單／會計／報表分析）都呼叫這支，確保視覺
  風格一致，不會有的頁面有質感、有的頁面看起來像沒套用到設計。
  """
  ui.add_head_html("""
        <link rel="preconnect" href="https://fonts.googleapis.com">
        <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
        <link href="https://fonts.googleapis.com/css2?family=Noto+Serif+TC:wght@600;700&family=Noto+Sans+TC:wght@400;500;700&family=IBM+Plex+Mono:wght@400;500&display=swap" rel="stylesheet">
        <style>
            :root {
              --bg: #f2f0ea;
              --surface: #ffffff;
              --surface-2: #f7f5ef;
              --border: #e6e1d4;
              --ink: #2a2823;
              --muted: #8a8577;
              --font-display: 'Noto Serif TC', serif;
              --font-body: 'Noto Sans TC', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
              --font-mono: 'IBM Plex Mono', ui-monospace, SFMono-Regular, monospace;
            }
            body { background-color: var(--bg); color: var(--ink); font-family: var(--font-body); }

            /* 表格：白底卡片、柔和陰影取代生硬的滿版邊框，數字用等寬字體 */
            .q-table__container { background-color: var(--surface) !important; border: 1px solid var(--border); border-radius: 10px; box-shadow: 0 1px 3px rgba(42,40,35,.06) !important; overflow: hidden; }
            .q-table th { color: var(--muted) !important; font-weight: 700 !important; font-size: 12px; letter-spacing: .02em; background: var(--surface-2) !important; border-bottom: 1px solid var(--border) !important; }
            .q-table td { color: var(--ink) !important; border-bottom: 1px solid var(--surface-2) !important; font-family: var(--font-body); }
            .q-table tbody tr:hover td { background: var(--surface-2) !important; }

            /* 分頁：底線取代粗黑框，字重輕一點更沉穩 */
            .q-tabs { border-bottom: 1px solid var(--border); }
            .q-tab { color: var(--muted) !important; font-weight: 500; text-transform: none; font-family: var(--font-body); }
            .q-tab--active { color: var(--ink) !important; font-weight: 700; }
            .q-tab-indicator { background: #5bc0be !important; height: 2px !important; }

            /* 按鈕：柔和圓角，主要動作按鈕用品牌色 */
            .q-btn { border-radius: 6px !important; font-family: var(--font-body); }
            .sync-btn { background-color: #4f9d9b !important; color: #ffffff !important; font-weight: 700; }

            /* 輸入框/下拉選單也統一柔和圓角 */
            .q-field__control { border-radius: 6px !important; }

            /* 標題用襯線字型，跟內文區分出層次 */
            h1, h2, .text-lg.font-bold, .text-xl.font-bold { font-family: var(--font-display); }
        </style>
    """)


def inject_company_tab_css():
  """公司分頁的顏色樣式（.company-tab-0～3），四個App頁面共用同一套，
  只要呼叫這支就能套用，不用每個頁面各自重複寫一次CSS字串。
  """
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


def render_company_switcher_placeholder(app_label):
  """公司切換列的骨架版本——四個新App目前還沒把實際內容搬過去，先讓
  「App切換＋公司切換」兩排並列的架構長出來，公司分頁點下去先顯示佔位
  畫面。之後搬內容過來時，把render_section_placeholder(...)那行換成
  真正的頁面渲染函式即可，不用動這裡的分頁結構。

  注意：公司分頁列(ui.tabs)一定要在內容容器(content_container)「之前」
  建立，這樣分頁列在畫面排列順序上才會天生排在內容前面（最上方）。順序
  反過來的話，分頁列雖然視覺上看起來應該在上面，內容容器的位置卻已經
  先佔走了，分頁列反而會被排到內容下方——這是本專案已經踩過兩次的坑，
  這裡先直接避開。
  """
  inject_company_tab_css()

  def handle_company_change(e):
    content_container.clear()
    with content_container:
      with ui.column().classes(
          "w-full p-8 max-w-[1600px] mx-auto items-center justify-center"
      ):
        with ui.card().classes(
            "w-full p-16 bg-white border border-[#e6e1d4]"
            " shadow-[0_1px_3px_rgba(42,40,35,0.06)] rounded-lg text-center"
        ):
          ui.label(f"{app_label}｜{e.value}").classes(
              "text-lg font-bold text-zinc-900 mb-2"
          )
          ui.label("這個App的內容還沒搬過來，敬請期待").classes(
              "text-sm text-zinc-500"
          )

  with ui.row().classes(
      "w-full flex flex-nowrap items-center bg-white border-b border-[#e6e1d4]"
      " px-8 py-3 sticky top-[41px] z-40"
  ):
    with ui.tabs(on_change=handle_company_change).props(
        "dense no-caps"
    ).classes("flex-shrink-0") as company_tabs:
      for i, c in enumerate(COMPANIES):
        ui.tab(c).classes(f"company-tab-{i}")

  content_container = ui.column().classes("w-full")
  company_tabs.set_value(COMPANIES[0])

  return content_container, company_tabs


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
  inject_global_theme_css()
  render_app_switcher("/")

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
  inject_company_tab_css()

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
      " border-b border-[#e6e1d4] px-8 py-4 sticky top-[41px] z-50"
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
            "w-full p-16 bg-white border border-[#e6e1d4] shadow-[0_1px_3px_rgba(42,40,35,0.06)]"
            " rounded-lg text-center"
        ):
          ui.label(name).classes("text-lg font-bold text-zinc-900 mb-2")
          ui.label("此分公司尚未串接 A1 API，敬請期待").classes(
              "text-sm text-zinc-500"
          )

  async def render_procurement_analysis(api_key, api_password, company_label):
    """分公司／採購分析：建議採購量／庫存週轉／進貨明細(先佔位)／月產銷
    分析，共用一次A1資料抓取(fetch_procurement_analysis_data)。跟海濤客
    那份的差異：
      - 銷售歷史直接即時從A1銷貨單算，不用另外維護Google Sheet
      - 月產銷分析不做BOM展開，只看成品/組合品本身的成本/售價
      - 進貨明細因為A1沒有查詢API、要另外維護Sheet，資料來源還沒確認，
        先維持佔位畫面

    這支是async函式，實際打A1 API的地方（load_data）用run.io_bound()包
    起來，避免卡住整個伺服器的事件迴圈；底下4個子分頁本身只是對同一份
    已抓好的資料做計算，不會額外打API，維持原本一次全部畫出來的做法。
    """
    data_holder = {"data": None, "error": None}
    last_updated = {"time": ""}

    async def load_data():
      data, error = await run.io_bound(
          fetch_procurement_analysis_data, api_key, api_password
      )
      data_holder["data"] = data
      data_holder["error"] = error
      last_updated["time"] = (datetime.utcnow() + timedelta(hours=8)).strftime("%Y-%m-%d %H:%M:%S")

    await load_data()

    refresh_row = ui.row().classes("w-full items-center gap-3 mb-3")

    def render_refresh_row():
      refresh_row.clear()
      with refresh_row:
        ui.label(f"資料更新時間：{last_updated['time']}").classes("text-xs text-zinc-400")

        async def handle_refresh():
          await load_data()
          render_refresh_row()
          render_tabs_content()
          if data_holder["error"]:
            ui.notify(f"重新整理失敗：{data_holder['error']}", color="negative")
          else:
            ui.notify("已重新整理", color="positive")

        ui.button("重新整理", icon="refresh", on_click=handle_refresh).props(
            "dense no-caps unelevated"
        ).classes("px-3 py-1 rounded-lg text-xs").style(
            "background:#ffffff !important; color:#4b5563 !important;"
            " border:1px solid #e6e1d4;"
        )

    render_refresh_row()

    body_container = ui.column().classes("w-full")

    def render_tabs_content():
      body_container.clear()
      with body_container:
        if data_holder["error"]:
          render_section_placeholder(
              "採購分析", f"抓取失敗：{data_holder['error']}"
          )
          return
        data = data_holder["data"]
        items_map = data["items_map"]
        stock_lookup = data["stock_lookup"]
        sales_history = data["sales_history"]

        with ui.tabs().props("dense no-caps").classes("w-full") as pa_tabs:
          pa_tab_procure = ui.tab("建議採購量")
          pa_tab_turnover = ui.tab("庫存週轉")
          pa_tab_receiving = ui.tab("進貨明細")
          pa_tab_forecast = ui.tab("月產銷分析")

        with ui.tab_panels(pa_tabs, value=pa_tab_procure).classes(
            "w-full bg-transparent"
        ):
          with ui.tab_panel(pa_tab_procure):
            rows = compute_suggested_procurement(items_map, stock_lookup)
            with ui.card().classes(
                "w-full p-6 bg-white border border-[#e6e1d4] shadow-[0_1px_3px_rgba(42,40,35,0.06)] rounded-lg"
            ):
              with ui.row().classes("w-full items-center justify-between mb-3"):
                ui.label(
                    f"建議採購量（安全庫存基準，共 {len(rows)} 項需要採購）"
                ).classes("text-sm font-bold text-zinc-700")

                def handle_export_procure():
                  try:
                    xlsx_bytes = rows_to_xlsx_bytes(rows, sheet_name="建議採購量")
                    ui.download(
                        xlsx_bytes, f"{company_label}建議採購量.xlsx",
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                  except Exception as e:
                    ui.notify(f"匯出失敗：{e}", color="negative")

                ui.button("匯出 xlsx", on_click=handle_export_procure).classes(
                    "sync-btn px-3 py-1 text-xs rounded-lg"
                )
              if not rows:
                ui.label("目前沒有品項需要採購（或商品主檔都沒設定安全庫存）").classes(
                    "text-xs text-zinc-400"
                )
              else:
                ui.table(
                    columns=[
                        {"name": c, "label": c, "field": c,
                         "align": "left" if c in ("品號", "品名", "商品分類") else "right",
                         "sortable": True}
                        for c in rows[0].keys()
                    ],
                    rows=rows, row_key="品號",
                    pagination={"rowsPerPage": 10, "sortBy": "建議採購量", "descending": True},
                ).classes("w-full").props(':rows-per-page-options="[10,30,50,0]"')

          with ui.tab_panel(pa_tab_turnover):
            rows = compute_turnover_metrics(sales_history, stock_lookup, items_map)
            slow_count = sum(1 for r in rows if r["滯銷"] == "是")
            with ui.card().classes(
                "w-full p-6 bg-white border border-[#e6e1d4] shadow-[0_1px_3px_rgba(42,40,35,0.06)] rounded-lg"
            ):
              with ui.row().classes("w-full items-center justify-between mb-3"):
                ui.label(
                    f"庫存週轉率／滯銷品分析（共 {len(rows)} 項，其中 "
                    f"{slow_count} 項判定為滯銷；只分析成品/組合品）"
                ).classes("text-sm font-bold text-zinc-700")

                def handle_export_turnover():
                  try:
                    xlsx_bytes = rows_to_xlsx_bytes(rows, sheet_name="庫存週轉")
                    ui.download(
                        xlsx_bytes, f"{company_label}庫存週轉分析.xlsx",
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                  except Exception as e:
                    ui.notify(f"匯出失敗：{e}", color="negative")

                ui.button("匯出 xlsx", on_click=handle_export_turnover).classes(
                    "sync-btn px-3 py-1 text-xs rounded-lg"
                )
              if not rows:
                ui.label("目前沒有成品/組合品的銷售歷史資料").classes("text-xs text-zinc-400")
              else:
                ui.table(
                    columns=[
                        {"name": c, "label": c, "field": c,
                         "align": "left" if c in ("品號", "品名", "滯銷") else "right",
                         "sortable": True}
                        for c in rows[0].keys()
                    ],
                    rows=rows, row_key="品號",
                    pagination={"rowsPerPage": 10, "sortBy": "庫存週轉天數", "descending": True},
                ).classes("w-full").props(':rows-per-page-options="[10,30,50,0]"')

          with ui.tab_panel(pa_tab_receiving):
            render_section_placeholder(
                "進貨明細",
                "資料來源尚未確認（A1沒有進貨查詢API，需要額外維護"
                "Google Sheet），敬請期待",
            )

          with ui.tab_panel(pa_tab_forecast):
            rows = compute_simple_monthly_forecast(items_map, sales_history)
            with ui.card().classes(
                "w-full p-6 bg-white border border-[#e6e1d4] shadow-[0_1px_3px_rgba(42,40,35,0.06)] rounded-lg"
            ):
              with ui.row().classes(
                  "w-full p-3 mb-3 bg-[#fff8e6] border border-[#f0dca0]"
              ):
                ui.label(
                    "簡化版：只分析成品/組合品，用商品本身的成本與依銷售"
                    "歷史反推的平均售價估算，不做BOM展開成原物料。"
                ).classes("text-xs text-amber-800")
              with ui.row().classes("w-full items-center justify-between mb-3"):
                ui.label(f"月產銷分析（共 {len(rows)} 項）").classes(
                    "text-sm font-bold text-zinc-700"
                )

                def handle_export_forecast():
                  try:
                    xlsx_bytes = rows_to_xlsx_bytes(rows, sheet_name="月產銷分析")
                    ui.download(
                        xlsx_bytes, f"{company_label}月產銷分析.xlsx",
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                  except Exception as e:
                    ui.notify(f"匯出失敗：{e}", color="negative")

                ui.button("匯出 xlsx", on_click=handle_export_forecast).classes(
                    "sync-btn px-3 py-1 text-xs rounded-lg"
                )
              if not rows:
                ui.label("目前沒有成品/組合品的銷售歷史資料").classes("text-xs text-zinc-400")
              else:
                ui.table(
                    columns=[
                        {"name": c, "label": c, "field": c,
                         "align": "left" if c in ("品號", "品名", "商品分類") else "right",
                         "sortable": True}
                        for c in rows[0].keys()
                    ],
                    rows=rows, row_key="品號",
                    pagination={"rowsPerPage": 10, "sortBy": "預估營收", "descending": True},
                ).classes("w-full").props(':rows-per-page-options="[10,30,50,0]"')

    render_tabs_content()


  def render_channel_company_page(company_name):
    """興聖(股)公司／容鴻(股)公司／芙萊柏(股)公司 共用的頁面骨架：
    儀表板／調撥紀錄／退換貨記錄／採購分析。

    「訂單出貨」「每日出貨」已經搬到「雲端電商訂單」App（/orders路由）
    去了，這裡不再顯示——共用的render_order_channels_tabs()/
    render_daily_shipping()等函式已經提升到模組層級，兩邊都能呼叫，
    內容邏輯本身沒有重複維護兩份。

    刻意做成「懶載入」：切換公司的當下只建立分頁導覽本身（很便宜），
    實際會打API抓資料的內容（採購分析）要等使用者真的點進那個分頁才
    觸發。原本的寫法是用ui.tab_panels把所有分頁內容一次全部建好，即使
    畫面上只顯示一個分頁，其他分頁的內容(含API呼叫)背地裡早就全部執行
    完了——這代表「切換一次公司」會同時觸發好幾個分頁各自的API呼叫，
    互相排隊等，才會覺得切換很鈍。
    """
    content_container.clear()
    accent = COMPANY_TAB_COLORS.get(company_name, {}).get("active_bg", "#5bc0be")
    slug = COMPANY_SLUGS.get(company_name, "default")
    tabs_class = f"section-tabs-{slug}"

    if company_name in ("容鴻(股)公司", "芙萊柏(股)公司"):
      SECTION_TABS = ["儀表板", "商品資訊", "調撥紀錄", "退換貨記錄", "採購分析"]
    else:
      SECTION_TABS = ["儀表板", "調撥紀錄", "退換貨記錄", "採購分析"]

    with content_container:
      with ui.column().classes("w-full p-8 max-w-[1600px] mx-auto gap-4"):
        ui.label(company_name).classes(
            "text-lg font-bold text-zinc-900"
        )
        with ui.tabs(on_change=lambda e: handle_section_change(e.value)).props(
            "dense no-caps"
        ).classes(f"w-full {tabs_class}") as section_tabs:
          for t in SECTION_TABS:
            ui.tab(t)
        # 只套用在這組分頁自己身上（用.section-tabs-xxx限定範圍），
        # 不會影響到最上面的公司切換列或其他分公司頁面的分頁顏色
        ui.add_head_html(
            f"<style>.{tabs_class}.q-tabs .q-tab--active {{ color: {accent} !important; }}"
            f" .{tabs_class}.q-tabs .q-tab-indicator {{ background: {accent} !important; }}</style>"
        )

        section_body = ui.column().classes("w-full")

        def render_loading():
          with ui.row().classes("w-full items-center gap-2 p-8 justify-center"):
            ui.spinner(size="24px").classes("text-zinc-400")
            ui.label("資料抓取中，請稍候…").classes("text-xs text-zinc-500")

        async def handle_section_change(tab_label):
          section_body.clear()
          with section_body:
            render_loading()
          # 讓上面的載入中畫面先真的畫出來，再開始跑會卡住的API呼叫，
          # 使用者才不會覺得畫面「整個沒反應」。
          await asyncio.sleep(0)

          if tab_label == "儀表板":
            section_body.clear()
            with section_body:
              render_section_placeholder(
                  "儀表板", "尚未串接此分公司的庫存／訂單資料，敬請期待"
              )
          elif tab_label == "商品資訊":
            section_body.clear()
            with section_body:
              with ui.tabs().props("dense no-caps").classes("w-full mb-2") as product_sub_tabs:
                ui.tab("庫存查詢")
              with ui.tab_panels(product_sub_tabs, value="庫存查詢").classes(
                  "w-full bg-transparent"
              ):
                with ui.tab_panel("庫存查詢"):
                  render_section_placeholder(
                      "庫存查詢", "此分公司尚未開通 A1 API，敬請期待"
                  )
          elif tab_label == "調撥紀錄":
            section_body.clear()
            with section_body:
              render_section_placeholder("調撥紀錄")
          elif tab_label == "退換貨記錄":
            section_body.clear()
            with section_body:
              render_section_placeholder("退換貨記錄")
          elif tab_label == "採購分析":
            section_body.clear()
            procurement_creds = PROCUREMENT_ANALYSIS_CREDENTIALS.get(company_name)
            if procurement_creds:
              with section_body:
                await render_procurement_analysis(
                    procurement_creds[0], procurement_creds[1], company_name,
                )
            else:
              with section_body:
                render_section_placeholder("採購分析")

        section_tabs.set_value("儀表板")

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
              "sync-btn px-4 py-2 text-xs rounded-lg"
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
                "w-full p-6 bg-white border border-[#e6e1d4] shadow-[0_1px_3px_rgba(42,40,35,0.06)]"
                " rounded-lg"
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
                      f"text-[11px] px-2 py-0.5 rounded-lg font-bold"
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
                    "sync-btn px-4 py-1 text-xs rounded-lg mt-3 self-end"
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
                        "sync-btn px-3 py-1 text-xs rounded-lg mt-3"
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

                # ---- 1.1/1.2：訂單需求相關 ----
                today = datetime.now().date()
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
                        ("shipping", "local_shipping", "訂單出貨提醒", "#2563eb"),
                        ("production", "factory", "生產組裝確認", "#9333ea"),
                        ("finished_goods", "receipt_long", "建議採購成品（母件）", "#db2777"),
                    ]
                    any_announcement = False
                    for key, icon_name, label, accent_color in category_labels:
                      items = announcements.get(key, [])[:6]
                      if not items:
                        continue
                      with ui.row().classes("items-center gap-1 mt-1"):
                        ui.icon(icon_name, size="16px").classes("text-zinc-500")
                        ui.label(label).classes(
                            "text-xs font-bold text-zinc-600"
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
              tab_inventory = ui.tab("庫存查詢")
              tab_bom = ui.tab("商品組合(BOM)")
              tab_lotno = ui.tab("批號效期")

            with ui.tab_panels(sub_tabs, value=tab_inventory).classes(
                "w-full bg-transparent"
            ):
              # ---------------- 2.3：庫存即時查詢 ----------------
              with ui.tab_panel(tab_inventory):
                with ui.card().classes(
                    "w-full p-6 bg-white border border-[#e6e1d4] shadow-[0_1px_3px_rgba(42,40,35,0.06)]"
                    " rounded-lg"
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
                          "bg-[#f7f6f2] text-zinc-900 rounded-lg px-3 py-1"
                          " text-xs font-bold border border-[#e6e1d4]"
                      )

                      cat_options = ["全部分類"] + app_state["categories"]
                      cat_select = ui.select(
                          options=cat_options, value="全部分類"
                      ).classes(
                          "bg-[#f7f6f2] text-zinc-900 rounded-lg px-3 py-1"
                          " text-xs font-bold border border-[#e6e1d4]"
                      )

                      search_input = ui.input(
                          placeholder="輸入品號或品名關鍵字..."
                      ).classes("w-64 text-xs")

                  stats_label = ui.label().classes("text-xs text-zinc-500 mb-3")

                  with ui.row().classes(
                      "w-full items-center gap-3 mb-4 p-3 bg-[#f7f6f2] border"
                      " border-[#e6e1d4] flex-wrap"
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
                        "sync-btn px-3 py-1 text-xs rounded-lg"
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
                    "w-full p-6 bg-white border border-[#e6e1d4] shadow-[0_1px_3px_rgba(42,40,35,0.06)]"
                    " rounded-lg"
                ):
                  ui.label("商品組合資訊（BOM）").classes(
                      "text-lg font-bold text-zinc-900 tracking-wide mb-2"
                  )
                  with ui.row().classes(
                      "w-full p-3 mb-4 bg-[#fff8e6] border border-[#f0dca0]"
                  ):
                    ui.label(
                        "鼎新 A1 目前的 API／後台匯出都只有組合品「主件」，"
                        "沒有「子件＋用量＋損耗率＋前置天數」明細，這些資料"
                        "改由「商品組合明細」Google Sheet 維護（尚未設定"
                        "Google Sheets 時，暫時退回本機 Excel 上傳作為過渡）。"
                    ).classes("text-xs text-amber-800")

                  bom_source_badge = ui.label().classes(
                      "text-xs font-bold mb-2"
                  )

                  with ui.row().classes(
                      "w-full items-center gap-3 flex-wrap mb-4 p-3"
                      " bg-[#f7f6f2] border border-[#e6e1d4]"
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
                    ).classes("sync-btn px-3 py-1 text-xs rounded-lg")

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
                            "w-full border border-[#e6e1d4] text-sm"
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
                            "w-full border border-[#e6e1d4] text-sm text-zinc-500"
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
                              f"商品組合明細中有 {len(orphan_ids)} 個主件品號，在"
                              f"目前 A1 商品主檔中查無此品號（可能是品號填錯，"
                              f"或該商品已停售）：{shown}{more}"
                          ).classes("text-xs text-red-700")

                  combo_search_input.on_value_change(lambda e: update_combo_list())
                  update_combo_list()

                  refs["update_combo_list"] = update_combo_list

              # ---------------- 2.3：批號／效期追蹤 ----------------
              with ui.tab_panel(tab_lotno):
                with ui.card().classes(
                    "w-full p-6 bg-white border border-[#e6e1d4] shadow-[0_1px_3px_rgba(42,40,35,0.06)]"
                    " rounded-lg"
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
                    ).classes("sync-btn px-3 py-1 text-xs rounded-lg")

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
          # 3. 訂單出貨（建立訂單／未出訂單查詢 兩個子分頁）
          # ==================================================
          with ui.tab_panel(tab_orders):
            with ui.tabs().classes("w-full mb-2") as orders_sub_tabs:
              tab_create_order = ui.tab("建立訂單")
              tab_unshipped_query = ui.tab("未出訂單查詢")
              tab_create_sale = ui.tab("銷貨單")
              tab_create_purchase = ui.tab("採購單")
              tab_create_receive = ui.tab("進貨單")

            with ui.tab_panels(orders_sub_tabs, value=tab_create_order).classes(
                "w-full bg-transparent"
            ):
              with ui.tab_panel(tab_create_order):
                # -----------------------------------------------------------
                # 手動建立訂單：給「經銷／團購」這種沒有系統資料來源、業務用
                # 電話/LINE臨時談成的訂單用——沒有任何自動化資料可以帶入，
                # 只能人工填。做成表單直接送出 Orders[Post]，避免還要切到
                # A1 系統裡重新輸入一次。客戶/商品選單直接用app_state裡已經
                # 抓好的資料（跟「商品資訊」「庫存查詢」用同一份），不用額外
                # 再打一次API。
                # -----------------------------------------------------------
                with ui.card().classes(
                    "w-full p-6 bg-white border border-[#e6e1d4]"
                    " shadow-[0_1px_3px_rgba(42,40,35,0.06)] rounded-lg mb-4"
                ):
                  ui.label("手動建立訂單（經銷／團購臨時接單）").classes(
                      "text-base font-bold text-zinc-900 mb-2"
                  )
                  ui.label(
                      "沒有系統資料來源的臨時訂單，在這裡填寫送出，會建立成 A1"
                      "的「訂單」（Orders[Post]，只記錄未來要出貨的承諾，不會"
                      "馬上扣庫存，出貨時仍要另外在 A1 開銷貨單），同時會把"
                      "品項寫進「訂單資訊」Google Sheet，出現在上面的訂單查詢"
                      "/缺貨提醒裡，食品廠可以直接看到還有多少未出、何時要出。"
                  ).classes("text-xs text-zinc-500 mb-3")

                  customer_options = {
                      cid: f"{cid} - {name}"
                      for cid, name in (app_state.get("customers_map") or {}).items()
                  }
                  item_options = {
                      iid: f"{iid} - {info.get('Name', '')}"
                      for iid, info in (app_state.get("items_map") or {}).items()
                  }

                  MANUAL_TAXTYPE_OPTIONS = {
                      "0": "免發票", "1": "應稅外加", "3": "免稅", "4": "應稅內含",
                  }

                  with ui.row().classes("items-end gap-3 flex-wrap mb-3"):
                    manual_customer_select = ui.select(
                        options=customer_options, label="客戶", with_input=True,
                    ).props("dense outlined").classes("w-72")
                    manual_predate_input = ui.input(
                        label="預交日期",
                        value=(datetime.now().date() + timedelta(days=3)).isoformat(),
                    ).props('dense outlined type="date"').classes("w-44")
                    manual_taxtype_select = ui.select(
                        options=MANUAL_TAXTYPE_OPTIONS,
                        label="課稅別", value="0",
                    ).props("dense outlined").classes("w-32")

                  manual_lines_container = ui.column().classes("w-full gap-2 mb-2")
                  manual_line_rows = []

                  manual_total_label = ui.label("合計金額：0").classes(
                      "text-sm font-bold text-zinc-700 mb-3"
                  )

                  def update_manual_total():
                    total = sum(
                        float(r["amount"].value or 0) for r in manual_line_rows
                    )
                    manual_total_label.text = f"合計金額：{total:,.0f}"

                  def add_manual_line():
                    with manual_lines_container:
                      with ui.row().classes("w-full items-center gap-2") as row:
                        item_sel = ui.select(
                            options=item_options, label="商品", with_input=True,
                        ).props("dense outlined").classes("flex-1")
                        qty_input = ui.number(
                            label="數量", value=1, min=0,
                        ).props("dense outlined").classes("w-24")
                        amount_input = ui.number(
                            label="金額", value=0, min=0,
                        ).props("dense outlined").classes("w-28")
                        memo_input = ui.input(
                            label="備註",
                        ).props("dense outlined").classes("w-36")

                        def remove_this_line():
                          manual_lines_container.remove(row)
                          manual_line_rows.remove(entry)
                          update_manual_total()

                        ui.button(icon="close", on_click=remove_this_line).props(
                            "flat dense round"
                        )
                    entry = {
                        "item_sel": item_sel, "qty": qty_input, "amount": amount_input,
                        "memo": memo_input,
                    }
                    manual_line_rows.append(entry)
                    amount_input.on_value_change(lambda e: update_manual_total())
                    update_manual_total()

                  with ui.row().classes("gap-2 mb-3"):
                    ui.button("+ 新增品項", on_click=add_manual_line).classes(
                        "px-3 py-1 text-xs rounded-lg"
                    )

                  add_manual_line()

                  manual_created_container = ui.column().classes("w-full gap-1 mt-2")
                  manual_created_list = []

                  def render_manual_created_list():
                    manual_created_container.clear()
                    if not manual_created_list:
                      return
                    with manual_created_container:
                      ui.label(
                          f"本次工作階段已建立 {len(manual_created_list)} 張"
                          "（這份清單只是操作確認、重新整理頁面就會消失；正式"
                          "持久的紀錄在「未出訂單查詢」分頁跟「訂單資訊」"
                          "Google Sheet 裡都查得到）"
                      ).classes("text-xs text-zinc-400")
                      for rec in reversed(manual_created_list[-10:]):
                        ui.label(
                            f"{rec['訂單編號']}｜{rec['客戶']}｜"
                            f"{rec['品項數']}個品項｜{rec['金額']:,.0f}元｜"
                            f"預交 {rec['預交日期']}"
                        ).classes("text-xs text-zinc-600")

                  with ui.dialog() as manual_confirm_dialog, ui.card().classes(
                      "min-w-[360px] max-w-[90vw] p-5"
                  ):
                    manual_confirm_dialog_body = ui.column().classes("w-full gap-2")
                    with ui.row().classes("w-full justify-end gap-2 mt-4"):
                      ui.button(
                          "取消", on_click=manual_confirm_dialog.close
                      ).classes("px-4 py-1 text-xs rounded-lg")
                      manual_confirm_button = ui.button("確認建立").classes(
                          "sync-btn px-4 py-1 text-xs rounded-lg"
                      )

                  manual_pending_payload = {"value": None}

                  def reset_manual_form():
                    manual_customer_select.value = None
                    manual_predate_input.value = (
                        datetime.now().date() + timedelta(days=3)
                    ).isoformat()
                    manual_taxtype_select.value = "0"
                    manual_lines_container.clear()
                    manual_line_rows.clear()
                    add_manual_line()

                  def handle_manual_submit_click():
                    if not manual_customer_select.value:
                      ui.notify("請選擇客戶", color="warning")
                      return
                    if not manual_line_rows:
                      ui.notify("請至少新增一個品項", color="warning")
                      return
                    try:
                      pre_date = datetime.strptime(
                          manual_predate_input.value, "%Y-%m-%d"
                      ).date()
                    except (ValueError, TypeError):
                      ui.notify("預交日期格式錯誤", color="warning")
                      return

                    details = []
                    subtotal = 0.0
                    for i, r in enumerate(manual_line_rows, start=1):
                      item_id = r["item_sel"].value
                      qty = float(r["qty"].value or 0)
                      amount = float(r["amount"].value or 0)
                      if not item_id or qty <= 0:
                        ui.notify(f"第 {i} 行商品或數量沒填好", color="warning")
                        return
                      line = {
                          "ID": i, "ItemID": item_id, "Qty": qty, "Amount": amount,
                          "PreDeliveryDate": pre_date.strftime("%Y/%m/%d"),
                      }
                      memo_val = (r["memo"].value or "").strip()
                      if memo_val:
                        line["Memo"] = memo_val
                      details.append(line)
                      subtotal += amount

                    tax_type = manual_taxtype_select.value
                    total_tax, total_sale_amount = compute_tax_and_total(
                        subtotal, tax_type
                    )

                    order_id = f"WEB{datetime.now().strftime('%Y%m%d%H%M%S%f')}"
                    payload = {
                        "ID": order_id,
                        "TradeDate": datetime.now().date().strftime("%Y/%m/%d"),
                        "CustomerID": manual_customer_select.value,
                        "TaxType": tax_type,
                        "TotalTax": total_tax,
                        "TotalSaleAmount": total_sale_amount,
                        "PreDeliveryDate": pre_date.strftime("%Y/%m/%d"),
                        # 手冊表格寫「Details」，範例JSON卻是「OrderDetails」，
                        # 兩者矛盾，實測後者才是正確欄位名稱（見
                        # build_order_upload_payloads的說明）。
                        "OrderDetails": details,
                    }
                    manual_pending_payload["value"] = payload

                    manual_confirm_dialog_body.clear()
                    with manual_confirm_dialog_body:
                      ui.label("確認建立訂單").classes(
                          "text-base font-bold text-zinc-900"
                      )
                      ui.label(
                          f"客戶：{customer_options.get(manual_customer_select.value, '')}"
                      ).classes("text-xs text-zinc-700")
                      ui.label(f"預交日期：{pre_date.isoformat()}").classes(
                          "text-xs text-zinc-700"
                      )
                      ui.label(
                          f"課稅別：{MANUAL_TAXTYPE_OPTIONS[tax_type]}"
                          f"　稅額：{total_tax:,.0f}"
                      ).classes("text-xs text-zinc-700")
                      ui.label(
                          f"共 {len(details)} 個品項，總金額（含稅）"
                          f" {total_sale_amount:,.0f}"
                      ).classes("text-xs text-zinc-700")
                      ui.label(
                          "送出後會直接寫入 A1 正式系統，請確認無誤。"
                      ).classes("text-xs text-amber-700 mt-2")
                    manual_confirm_dialog.open()

                  def handle_manual_confirm_click():
                    payload = manual_pending_payload["value"]
                    if not payload:
                      manual_confirm_dialog.close()
                      return
                    token = get_a1_token()
                    if not token:
                      ui.notify("無法登入 A1，請確認 API 憑證", color="warning")
                      return
                    ok, msg = upload_order_to_a1(token, payload)
                    manual_confirm_dialog.close()
                    if ok:
                      ui.notify(f"訂單建立成功（{payload['ID']}）", color="positive")

                      # 同時把每個品項寫成一列，附加到「訂單資訊」Google
                      # Sheet——這樣食品廠原本就在看的訂單出貨查詢/缺貨提醒
                      # 會自動看到這張新訂單，不用另外開地方查。就算Sheet
                      # 沒設定或寫入失敗，A1訂單已經成功建立，不會因此卡住，
                      # 只會另外提示一次警告。
                      items_map_now = app_state.get("items_map") or {}
                      sheet_rows = [
                          {
                              ORDER_COL_NO: payload["ID"],
                              ORDER_COL_ITEM_ID: d["ItemID"],
                              ORDER_COL_ITEM_NAME: items_map_now.get(
                                  d["ItemID"], {}
                              ).get("Name", ""),
                              ORDER_COL_DUE_DATE: payload["PreDeliveryDate"],
                              ORDER_COL_QTY: d["Qty"],
                              ORDER_COL_STATUS: "未出貨",
                              ORDER_COL_CUSTOMER_ID: payload["CustomerID"],
                              ORDER_COL_AMOUNT: d["Amount"],
                              ORDER_COL_MEMO: "經銷/團購臨時接單，網頁手動建立",
                          }
                          for d in payload["OrderDetails"]
                      ]
                      sheet_count, sheet_err = append_rows_to_google_sheet(
                          ORDERS_GOOGLE_SHEET_TAB, sheet_rows
                      )
                      if sheet_err:
                        ui.notify(
                            f"A1訂單已建立成功，但寫入Google Sheet失敗："
                            f"{sheet_err}（訂單編號：{payload['ID']}，可以自己"
                            f"手動補登記到Sheet）",
                            color="warning",
                        )
                      else:
                        ui.notify(
                            f"已同步寫入「訂單資訊」Sheet（{sheet_count}列）",
                            color="positive",
                        )
                        if refs.get("update_orders_list"):
                          refs["update_orders_list"]()

                      manual_created_list.append({
                          "訂單編號": payload["ID"],
                          "客戶": customer_options.get(
                              payload["CustomerID"], payload["CustomerID"]
                          ),
                          "預交日期": payload["PreDeliveryDate"],
                          "金額": payload["TotalSaleAmount"],
                          "品項數": len(payload["OrderDetails"]),
                      })
                      render_manual_created_list()
                      reset_manual_form()
                    else:
                      ui.notify(f"建立失敗：{msg}", color="negative")
                      print(f"[手動建立訂單] 上傳失敗：{payload['ID']}｜{msg}")

                  manual_confirm_button.on_click(handle_manual_confirm_click)

                  ui.button(
                      "建立訂單", on_click=handle_manual_submit_click,
                  ).classes(
                      "px-4 py-2 text-xs rounded-lg bg-amber-600 text-white font-bold"
                  )

                  render_manual_created_list()

              with ui.tab_panel(tab_unshipped_query):
                  orders_reminder_container = ui.column().classes("w-full gap-2 mb-4")

                  with ui.row().classes("items-center gap-3 flex-wrap mb-3"):
                    orders_search_input = ui.input(
                        placeholder="輸入品號、品名或訂單編號..."
                    ).classes("w-64 text-xs")
                    orders_status_select = ui.select(
                        options=["全部狀態", "未出貨", "備貨中", "已出貨"],
                        value="全部狀態",
                    ).classes(
                        "bg-[#f7f6f2] text-zinc-900 rounded-lg px-3 py-1"
                        " text-xs font-bold border border-[#e6e1d4]"
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

              # ---------------------------------------------------------
              # 銷貨單／採購單／進貨單：邏輯跟「建立訂單」一樣（表單+確認
              # 彈窗+送出），但不統計已上傳紀錄，單純當上傳管道。
              # ---------------------------------------------------------
              customer_options_doc = {
                  cid: f"{cid} - {name}"
                  for cid, name in (app_state.get("customers_map") or {}).items()
              }
              supplier_options_doc = {
                  sid: f"{sid} - {name}"
                  for sid, name in (app_state.get("suppliers_map") or {}).items()
              }
              item_options_doc = {
                  iid: f"{iid} - {info.get('Name', '')}"
                  for iid, info in (app_state.get("items_map") or {}).items()
              }
              warehouse_options_doc = {
                  w: w for w in (app_state.get("warehouses") or [])
              }

              with ui.tab_panel(tab_create_sale):
                with ui.card().classes(
                    "w-full p-6 bg-white border border-[#e6e1d4]"
                    " shadow-[0_1px_3px_rgba(42,40,35,0.06)] rounded-lg"
                ):
                  ui.label("上傳銷貨單").classes(
                      "text-base font-bold text-zinc-900 mb-2"
                  )
                  ui.label(
                      "會建立成 A1 的「銷貨單」（Sales[Post]），會馬上扣庫存"
                      "、視同已完成的銷售，跟「訂單」不一樣。這裡不會留存"
                      "上傳紀錄，純粹是送出管道，正式紀錄請至 A1 查看。"
                  ).classes("text-xs text-zinc-500 mb-3")

                  with ui.row().classes("items-end gap-3 flex-wrap mb-3"):
                    sale_customer_select = ui.select(
                        options=customer_options_doc, label="客戶", with_input=True,
                    ).props("dense outlined").classes("w-64")
                    sale_payment_select = ui.select(
                        options={
                            "1": "現金", "2": "信用卡", "3": "轉帳", "M": "賒銷(月結)",
                        },
                        label="收款方式", value="M",
                    ).props("dense outlined").classes("w-32")
                    sale_taxtype_select = ui.select(
                        options=MANUAL_TAXTYPE_OPTIONS, label="課稅別", value="0",
                    ).props("dense outlined").classes("w-32")

                  sale_lines_container = ui.column().classes("w-full gap-2 mb-2")
                  sale_line_rows = []
                  sale_total_label = ui.label("合計金額：0").classes(
                      "text-sm font-bold text-zinc-700 mb-3"
                  )

                  def update_sale_total():
                    total = sum(float(r["amount"].value or 0) for r in sale_line_rows)
                    sale_total_label.text = f"合計金額：{total:,.0f}"

                  def add_sale_line():
                    with sale_lines_container:
                      with ui.row().classes("w-full items-center gap-2") as row:
                        item_sel = ui.select(
                            options=item_options_doc, label="商品", with_input=True,
                        ).props("dense outlined").classes("flex-1")
                        qty_input = ui.number(label="數量", value=1, min=0).props(
                            "dense outlined"
                        ).classes("w-20")
                        amount_input = ui.number(label="金額", value=0, min=0).props(
                            "dense outlined"
                        ).classes("w-24")
                        warehouse_sel = ui.select(
                            options=warehouse_options_doc, label="倉庫",
                        ).props("dense outlined").classes("w-28")
                        lotno_input = ui.input(label="批號").props(
                            "dense outlined"
                        ).classes("w-24")
                        isfree_checkbox = ui.checkbox("贈品")
                        memo_input = ui.input(label="備註").props(
                            "dense outlined"
                        ).classes("w-28")

                        def remove_this_line():
                          sale_lines_container.remove(row)
                          sale_line_rows.remove(entry)
                          update_sale_total()

                        ui.button(icon="close", on_click=remove_this_line).props(
                            "flat dense round"
                        )
                    entry = {
                        "item_sel": item_sel, "qty": qty_input, "amount": amount_input,
                        "warehouse": warehouse_sel, "lotno": lotno_input,
                        "isfree": isfree_checkbox, "memo": memo_input,
                    }
                    sale_line_rows.append(entry)
                    amount_input.on_value_change(lambda e: update_sale_total())
                    update_sale_total()

                  with ui.row().classes("gap-2 mb-3"):
                    ui.button("+ 新增品項", on_click=add_sale_line).classes(
                        "px-3 py-1 text-xs rounded-lg"
                    )
                  add_sale_line()

                  with ui.dialog() as sale_confirm_dialog, ui.card().classes(
                      "min-w-[360px] max-w-[90vw] p-5"
                  ):
                    sale_confirm_dialog_body = ui.column().classes("w-full gap-2")
                    with ui.row().classes("w-full justify-end gap-2 mt-4"):
                      ui.button(
                          "取消", on_click=sale_confirm_dialog.close
                      ).classes("px-4 py-1 text-xs rounded-lg")
                      sale_confirm_button = ui.button("確認建立").classes(
                          "sync-btn px-4 py-1 text-xs rounded-lg"
                      )

                  sale_pending_payload = {"value": None}

                  def reset_sale_form():
                    sale_customer_select.value = None
                    sale_payment_select.value = "M"
                    sale_taxtype_select.value = "0"
                    sale_lines_container.clear()
                    sale_line_rows.clear()
                    add_sale_line()

                  def handle_sale_submit_click():
                    if not sale_customer_select.value:
                      ui.notify("請選擇客戶", color="warning")
                      return
                    if not sale_line_rows:
                      ui.notify("請至少新增一個品項", color="warning")
                      return
                    details = []
                    subtotal = 0.0
                    for i, r in enumerate(sale_line_rows, start=1):
                      item_id = r["item_sel"].value
                      qty = float(r["qty"].value or 0)
                      amount = float(r["amount"].value or 0)
                      warehouse = r["warehouse"].value
                      if not item_id or qty <= 0 or not warehouse:
                        ui.notify(
                            f"第 {i} 行商品／數量／倉庫沒填好（倉庫必填）",
                            color="warning",
                        )
                        return
                      line = {
                          "ID": i, "ItemID": item_id, "Qty": qty, "Amount": amount,
                          "Warehouse": warehouse, "IsFree": bool(r["isfree"].value),
                      }
                      if (r["lotno"].value or "").strip():
                        line["LotNo"] = r["lotno"].value.strip()
                      if (r["memo"].value or "").strip():
                        line["Memo"] = r["memo"].value.strip()
                      details.append(line)
                      subtotal += amount

                    tax_type = sale_taxtype_select.value
                    total_tax, total_amount = compute_tax_and_total(subtotal, tax_type)

                    sale_id = f"WEB{datetime.now().strftime('%Y%m%d%H%M%S%f')}"
                    payload = {
                        "ID": sale_id,
                        "TradeDate": datetime.now().date().strftime("%Y/%m/%d"),
                        "CustomerID": sale_customer_select.value,
                        "Payment": sale_payment_select.value,
                        "TaxType": tax_type,
                        "TotalTax": total_tax,
                        "TotalSaleAmount": total_amount,
                        "SaleDetails": details,
                    }
                    sale_pending_payload["value"] = payload

                    sale_confirm_dialog_body.clear()
                    with sale_confirm_dialog_body:
                      ui.label("確認建立銷貨單").classes(
                          "text-base font-bold text-zinc-900"
                      )
                      ui.label(
                          f"客戶：{customer_options_doc.get(sale_customer_select.value, '')}"
                      ).classes("text-xs text-zinc-700")
                      ui.label(
                          f"課稅別：{MANUAL_TAXTYPE_OPTIONS[tax_type]}　"
                          f"稅額：{total_tax:,.0f}"
                      ).classes("text-xs text-zinc-700")
                      ui.label(
                          f"共 {len(details)} 個品項，總金額（含稅）"
                          f" {total_amount:,.0f}"
                      ).classes("text-xs text-zinc-700")
                      ui.label(
                          "銷貨單會馬上扣庫存，送出後無法在這裡復原，請確認"
                          "無誤。"
                      ).classes("text-xs text-amber-700 mt-2")
                    sale_confirm_dialog.open()

                  def handle_sale_confirm_click():
                    payload = sale_pending_payload["value"]
                    if not payload:
                      sale_confirm_dialog.close()
                      return
                    token = get_a1_token()
                    if not token:
                      ui.notify("無法登入 A1，請確認 API 憑證", color="warning")
                      return
                    ok, msg = upload_sale_to_a1(token, payload)
                    sale_confirm_dialog.close()
                    if ok:
                      ui.notify(f"銷貨單建立成功（{payload['ID']}）", color="positive")
                      reset_sale_form()
                    else:
                      ui.notify(f"建立失敗：{msg}", color="negative")
                      print(f"[銷貨單上傳] 失敗：{payload['ID']}｜{msg}")

                  sale_confirm_button.on_click(handle_sale_confirm_click)
                  ui.button("建立銷貨單", on_click=handle_sale_submit_click).classes(
                      "px-4 py-2 text-xs rounded-lg bg-amber-600 text-white font-bold"
                  )

              with ui.tab_panel(tab_create_purchase):
                with ui.card().classes(
                    "w-full p-6 bg-white border border-[#e6e1d4]"
                    " shadow-[0_1px_3px_rgba(42,40,35,0.06)] rounded-lg"
                ):
                  ui.label("上傳採購單").classes(
                      "text-base font-bold text-zinc-900 mb-2"
                  )
                  ui.label(
                      "會建立成 A1 的「採購單」（Purchases[Post]），只記錄"
                      "要跟廠商採購的品項，不影響庫存，等實際進貨、開進貨單"
                      "時才會增加庫存。這裡不會留存上傳紀錄，正式紀錄請至"
                      "A1 查看。"
                  ).classes("text-xs text-zinc-500 mb-3")

                  with ui.row().classes("items-end gap-3 flex-wrap mb-3"):
                    purchase_supplier_select = ui.select(
                        options=supplier_options_doc, label="廠商", with_input=True,
                    ).props("dense outlined").classes("w-64")
                    purchase_predate_input = ui.input(
                        label="預交日期",
                        value=(datetime.now().date() + timedelta(days=3)).isoformat(),
                    ).props('dense outlined type="date"').classes("w-44")
                    purchase_taxtype_select = ui.select(
                        options=MANUAL_TAXTYPE_OPTIONS, label="課稅別", value="0",
                    ).props("dense outlined").classes("w-32")

                  purchase_lines_container = ui.column().classes("w-full gap-2 mb-2")
                  purchase_line_rows = []
                  purchase_total_label = ui.label("合計金額：0").classes(
                      "text-sm font-bold text-zinc-700 mb-3"
                  )

                  def update_purchase_total():
                    total = sum(
                        float(r["amount"].value or 0) for r in purchase_line_rows
                    )
                    purchase_total_label.text = f"合計金額：{total:,.0f}"

                  def add_purchase_line():
                    with purchase_lines_container:
                      with ui.row().classes("w-full items-center gap-2") as row:
                        item_sel = ui.select(
                            options=item_options_doc, label="商品", with_input=True,
                        ).props("dense outlined").classes("flex-1")
                        qty_input = ui.number(label="數量", value=1, min=0).props(
                            "dense outlined"
                        ).classes("w-24")
                        amount_input = ui.number(label="金額", value=0, min=0).props(
                            "dense outlined"
                        ).classes("w-28")
                        memo_input = ui.input(label="備註").props(
                            "dense outlined"
                        ).classes("w-36")

                        def remove_this_line():
                          purchase_lines_container.remove(row)
                          purchase_line_rows.remove(entry)
                          update_purchase_total()

                        ui.button(icon="close", on_click=remove_this_line).props(
                            "flat dense round"
                        )
                    entry = {
                        "item_sel": item_sel, "qty": qty_input,
                        "amount": amount_input, "memo": memo_input,
                    }
                    purchase_line_rows.append(entry)
                    amount_input.on_value_change(lambda e: update_purchase_total())
                    update_purchase_total()

                  with ui.row().classes("gap-2 mb-3"):
                    ui.button("+ 新增品項", on_click=add_purchase_line).classes(
                        "px-3 py-1 text-xs rounded-lg"
                    )
                  add_purchase_line()

                  with ui.dialog() as purchase_confirm_dialog, ui.card().classes(
                      "min-w-[360px] max-w-[90vw] p-5"
                  ):
                    purchase_confirm_dialog_body = ui.column().classes("w-full gap-2")
                    with ui.row().classes("w-full justify-end gap-2 mt-4"):
                      ui.button(
                          "取消", on_click=purchase_confirm_dialog.close
                      ).classes("px-4 py-1 text-xs rounded-lg")
                      purchase_confirm_button = ui.button("確認建立").classes(
                          "sync-btn px-4 py-1 text-xs rounded-lg"
                      )

                  purchase_pending_payload = {"value": None}

                  def reset_purchase_form():
                    purchase_supplier_select.value = None
                    purchase_predate_input.value = (
                        datetime.now().date() + timedelta(days=3)
                    ).isoformat()
                    purchase_taxtype_select.value = "0"
                    purchase_lines_container.clear()
                    purchase_line_rows.clear()
                    add_purchase_line()

                  def handle_purchase_submit_click():
                    if not purchase_supplier_select.value:
                      ui.notify("請選擇廠商", color="warning")
                      return
                    if not purchase_line_rows:
                      ui.notify("請至少新增一個品項", color="warning")
                      return
                    try:
                      pre_date = datetime.strptime(
                          purchase_predate_input.value, "%Y-%m-%d"
                      ).date()
                    except (ValueError, TypeError):
                      ui.notify("預交日期格式錯誤", color="warning")
                      return

                    details = []
                    subtotal = 0.0
                    for i, r in enumerate(purchase_line_rows, start=1):
                      item_id = r["item_sel"].value
                      qty = float(r["qty"].value or 0)
                      amount = float(r["amount"].value or 0)
                      if not item_id or qty <= 0:
                        ui.notify(f"第 {i} 行商品或數量沒填好", color="warning")
                        return
                      line = {
                          "ID": i, "ItemID": item_id, "Qty": qty, "Amount": amount,
                          "PreDeliveryDate": pre_date.strftime("%Y/%m/%d"),
                      }
                      if (r["memo"].value or "").strip():
                        line["Memo"] = r["memo"].value.strip()
                      details.append(line)
                      subtotal += amount

                    tax_type = purchase_taxtype_select.value
                    total_tax, total_amount = compute_tax_and_total(subtotal, tax_type)

                    purchase_id = f"WEB{datetime.now().strftime('%Y%m%d%H%M%S%f')}"
                    payload = {
                        "ID": purchase_id,
                        "TradeDate": datetime.now().date().strftime("%Y/%m/%d"),
                        "SupplierID": purchase_supplier_select.value,
                        "TaxType": tax_type,
                        "TotalTax": total_tax,
                        "TotalAmount": total_amount,
                        "PreDeliveryDate": pre_date.strftime("%Y/%m/%d"),
                        "PurchaseDetails": details,
                    }
                    purchase_pending_payload["value"] = payload

                    purchase_confirm_dialog_body.clear()
                    with purchase_confirm_dialog_body:
                      ui.label("確認建立採購單").classes(
                          "text-base font-bold text-zinc-900"
                      )
                      ui.label(
                          f"廠商：{supplier_options_doc.get(purchase_supplier_select.value, '')}"
                      ).classes("text-xs text-zinc-700")
                      ui.label(f"預交日期：{pre_date.isoformat()}").classes(
                          "text-xs text-zinc-700"
                      )
                      ui.label(
                          f"共 {len(details)} 個品項，總金額（含稅）"
                          f" {total_amount:,.0f}"
                      ).classes("text-xs text-zinc-700")
                      ui.label(
                          "送出後會直接寫入 A1 正式系統，請確認無誤。"
                      ).classes("text-xs text-amber-700 mt-2")
                    purchase_confirm_dialog.open()

                  def handle_purchase_confirm_click():
                    payload = purchase_pending_payload["value"]
                    if not payload:
                      purchase_confirm_dialog.close()
                      return
                    token = get_a1_token()
                    if not token:
                      ui.notify("無法登入 A1，請確認 API 憑證", color="warning")
                      return
                    ok, msg = upload_purchase_to_a1(token, payload)
                    purchase_confirm_dialog.close()
                    if ok:
                      ui.notify(f"採購單建立成功（{payload['ID']}）", color="positive")
                      reset_purchase_form()
                    else:
                      ui.notify(f"建立失敗：{msg}", color="negative")
                      print(f"[採購單上傳] 失敗：{payload['ID']}｜{msg}")

                  purchase_confirm_button.on_click(handle_purchase_confirm_click)
                  ui.button(
                      "建立採購單", on_click=handle_purchase_submit_click,
                  ).classes(
                      "px-4 py-2 text-xs rounded-lg bg-amber-600 text-white font-bold"
                  )

              with ui.tab_panel(tab_create_receive):
                with ui.card().classes(
                    "w-full p-6 bg-white border border-[#e6e1d4]"
                    " shadow-[0_1px_3px_rgba(42,40,35,0.06)] rounded-lg"
                ):
                  ui.label("上傳進貨單").classes(
                      "text-base font-bold text-zinc-900 mb-2"
                  )
                  ui.label(
                      "會建立成 A1 的「進貨單」（Receives[Post]），會馬上"
                      "增加庫存。這裡不會留存上傳紀錄，正式紀錄請至 A1"
                      "查看。"
                  ).classes("text-xs text-zinc-500 mb-3")
                  ui.label(
                      "手冊裡這個欄位名稱有兩種寫法（欄位表寫"
                      "「ReceiveDetails」，範例JSON寫成「ReceiDetails」，"
                      "兩者不一致），這裡先用手冊表格的完整拼法，如果送出"
                      "後跟「上傳訂單」一開始遇到的狀況一樣出現"
                      "「400002訂單單身不可空白」，代表要改成另一種拼法，"
                      "跟我說一聲我立刻改。"
                  ).classes("text-xs text-amber-700 mb-3")

                  with ui.row().classes("items-end gap-3 flex-wrap mb-3"):
                    receive_supplier_select = ui.select(
                        options=supplier_options_doc, label="廠商", with_input=True,
                    ).props("dense outlined").classes("w-64")
                    receive_payment_select = ui.select(
                        options={"1": "現金", "M": "賒進(月結)"},
                        label="付款方式", value="M",
                    ).props("dense outlined").classes("w-32")
                    receive_taxtype_select = ui.select(
                        options=MANUAL_TAXTYPE_OPTIONS, label="課稅別", value="0",
                    ).props("dense outlined").classes("w-32")

                  receive_lines_container = ui.column().classes("w-full gap-2 mb-2")
                  receive_line_rows = []
                  receive_total_label = ui.label("合計金額：0").classes(
                      "text-sm font-bold text-zinc-700 mb-3"
                  )

                  def update_receive_total():
                    total = sum(
                        float(r["amount"].value or 0) for r in receive_line_rows
                    )
                    receive_total_label.text = f"合計金額：{total:,.0f}"

                  def add_receive_line():
                    with receive_lines_container:
                      with ui.row().classes("w-full items-center gap-2") as row:
                        item_sel = ui.select(
                            options=item_options_doc, label="商品", with_input=True,
                        ).props("dense outlined").classes("flex-1")
                        qty_input = ui.number(label="數量", value=1, min=0).props(
                            "dense outlined"
                        ).classes("w-20")
                        amount_input = ui.number(label="金額", value=0, min=0).props(
                            "dense outlined"
                        ).classes("w-24")
                        warehouse_sel = ui.select(
                            options=warehouse_options_doc, label="倉庫",
                        ).props("dense outlined").classes("w-28")
                        purchaseno_input = ui.input(label="採購單號").props(
                            "dense outlined"
                        ).classes("w-28")
                        memo_input = ui.input(label="備註").props(
                            "dense outlined"
                        ).classes("w-28")

                        def remove_this_line():
                          receive_lines_container.remove(row)
                          receive_line_rows.remove(entry)
                          update_receive_total()

                        ui.button(icon="close", on_click=remove_this_line).props(
                            "flat dense round"
                        )
                    entry = {
                        "item_sel": item_sel, "qty": qty_input, "amount": amount_input,
                        "warehouse": warehouse_sel, "purchaseno": purchaseno_input,
                        "memo": memo_input,
                    }
                    receive_line_rows.append(entry)
                    amount_input.on_value_change(lambda e: update_receive_total())
                    update_receive_total()

                  with ui.row().classes("gap-2 mb-3"):
                    ui.button("+ 新增品項", on_click=add_receive_line).classes(
                        "px-3 py-1 text-xs rounded-lg"
                    )
                  add_receive_line()

                  with ui.dialog() as receive_confirm_dialog, ui.card().classes(
                      "min-w-[360px] max-w-[90vw] p-5"
                  ):
                    receive_confirm_dialog_body = ui.column().classes("w-full gap-2")
                    with ui.row().classes("w-full justify-end gap-2 mt-4"):
                      ui.button(
                          "取消", on_click=receive_confirm_dialog.close
                      ).classes("px-4 py-1 text-xs rounded-lg")
                      receive_confirm_button = ui.button("確認建立").classes(
                          "sync-btn px-4 py-1 text-xs rounded-lg"
                      )

                  receive_pending_payload = {"value": None}

                  def reset_receive_form():
                    receive_supplier_select.value = None
                    receive_payment_select.value = "M"
                    receive_taxtype_select.value = "0"
                    receive_lines_container.clear()
                    receive_line_rows.clear()
                    add_receive_line()

                  def handle_receive_submit_click():
                    if not receive_supplier_select.value:
                      ui.notify("請選擇廠商", color="warning")
                      return
                    if not receive_line_rows:
                      ui.notify("請至少新增一個品項", color="warning")
                      return
                    details = []
                    subtotal = 0.0
                    for i, r in enumerate(receive_line_rows, start=1):
                      item_id = r["item_sel"].value
                      qty = float(r["qty"].value or 0)
                      amount = float(r["amount"].value or 0)
                      warehouse = r["warehouse"].value
                      if not item_id or qty <= 0 or not warehouse:
                        ui.notify(
                            f"第 {i} 行商品／數量／倉庫沒填好（倉庫必填）",
                            color="warning",
                        )
                        return
                      line = {
                          "ID": i, "ItemID": item_id, "Qty": qty, "Amount": amount,
                          "Warehouse": warehouse,
                      }
                      if (r["purchaseno"].value or "").strip():
                        line["PurchaseNo"] = r["purchaseno"].value.strip()
                      if (r["memo"].value or "").strip():
                        line["Memo"] = r["memo"].value.strip()
                      details.append(line)
                      subtotal += amount

                    tax_type = receive_taxtype_select.value
                    total_tax, total_amount = compute_tax_and_total(subtotal, tax_type)

                    receive_id = f"WEB{datetime.now().strftime('%Y%m%d%H%M%S%f')}"
                    payload = {
                        "ID": receive_id,
                        "TradeDate": datetime.now().date().strftime("%Y/%m/%d"),
                        "SupplierID": receive_supplier_select.value,
                        "Payment": receive_payment_select.value,
                        "TaxType": tax_type,
                        "TotalTax": total_tax,
                        "TotalAmount": total_amount,
                        "ReceiveDetails": details,
                    }
                    receive_pending_payload["value"] = payload

                    receive_confirm_dialog_body.clear()
                    with receive_confirm_dialog_body:
                      ui.label("確認建立進貨單").classes(
                          "text-base font-bold text-zinc-900"
                      )
                      ui.label(
                          f"廠商：{supplier_options_doc.get(receive_supplier_select.value, '')}"
                      ).classes("text-xs text-zinc-700")
                      ui.label(
                          f"共 {len(details)} 個品項，總金額（含稅）"
                          f" {total_amount:,.0f}"
                      ).classes("text-xs text-zinc-700")
                      ui.label(
                          "進貨單會馬上增加庫存，送出後無法在這裡復原，請"
                          "確認無誤。"
                      ).classes("text-xs text-amber-700 mt-2")
                    receive_confirm_dialog.open()

                  def handle_receive_confirm_click():
                    payload = receive_pending_payload["value"]
                    if not payload:
                      receive_confirm_dialog.close()
                      return
                    token = get_a1_token()
                    if not token:
                      ui.notify("無法登入 A1，請確認 API 憑證", color="warning")
                      return
                    ok, msg = upload_receive_to_a1(token, payload)
                    receive_confirm_dialog.close()
                    if ok:
                      ui.notify(f"進貨單建立成功（{payload['ID']}）", color="positive")
                      reset_receive_form()
                    else:
                      ui.notify(f"建立失敗：{msg}", color="negative")
                      print(f"[進貨單上傳] 失敗：{payload['ID']}｜{msg}")

                  receive_confirm_button.on_click(handle_receive_confirm_click)
                  ui.button(
                      "建立進貨單", on_click=handle_receive_submit_click,
                  ).classes(
                      "px-4 py-2 text-xs rounded-lg bg-amber-600 text-white font-bold"
                  )

          # ==================================================
          # 4. 生產與包裝排程
          # ==================================================
          with ui.tab_panel(tab_production):
            with ui.card().classes(
                "w-full p-6 bg-white border border-[#e6e1d4] shadow-[0_1px_3px_rgba(42,40,35,0.06)]"
                " rounded-lg"
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
                          "成品庫存不足，需先確認能否即時生產/組裝"
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
                    "w-full p-6 bg-white border border-[#e6e1d4] shadow-[0_1px_3px_rgba(42,40,35,0.06)]"
                    " rounded-lg"
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
                        "bg-[#f7f6f2] text-zinc-900 rounded-lg px-3 py-1"
                        " text-xs font-bold border border-[#e6e1d4]"
                    )
                    procurement_search_input = ui.input(
                        placeholder="輸入品號或品名關鍵字..."
                    ).classes("w-64 text-xs")
                    procurement_scope_select = ui.select(
                        options=["僅顯示需要採購", "顯示全部商品"],
                        value="僅顯示需要採購",
                    ).classes(
                        "bg-[#f7f6f2] text-zinc-900 rounded-lg px-3 py-1"
                        " text-xs font-bold border border-[#e6e1d4]"
                    )

                  with ui.row().classes(
                      "w-full p-2 mb-2 bg-[#e8f6f5] border border-[#bfe6e3]"
                  ):
                    ui.label(
                        "預設只顯示「現有庫存 ≤ 安全庫存」的品項。如果你"
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
                    "w-full p-6 bg-white border border-[#e6e1d4] shadow-[0_1px_3px_rgba(42,40,35,0.06)]"
                    " rounded-lg"
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
                    ).classes("sync-btn px-3 py-1 text-xs rounded-lg")

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
                    "w-full p-6 bg-white border border-[#e6e1d4] shadow-[0_1px_3px_rgba(42,40,35,0.06)]"
                    " rounded-lg"
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
                    "w-full p-6 bg-white border border-[#e6e1d4] shadow-[0_1px_3px_rgba(42,40,35,0.06)]"
                    " rounded-lg"
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
                        "bg-[#f7f6f2] text-zinc-900 rounded-lg px-3 py-1"
                        " text-xs font-bold border border-[#e6e1d4] w-32"
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
                        "sync-btn px-4 py-2 text-xs rounded-lg"
                    )

                  forecast_fetch_status_label = ui.label().classes(
                      "text-xs text-zinc-500 mb-1"
                  )
                  ui.label(
                      "若去年的品號編碼跟今年不一樣，系統會完全比對不到"
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
                            "px-3 py-1 text-xs rounded-lg "
                            + (
                                "sync-btn"
                                if is_active
                                else "bg-[#f7f6f2] text-zinc-700 border border-[#e6e1d4]"
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
                            "bg-[#f7f6f2] border border-[#e6e1d4] p-4"
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
                        ).classes("sync-btn px-3 py-1 text-xs rounded-lg")

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
                  "w-full p-6 bg-white border border-[#e6e1d4] shadow-[0_1px_3px_rgba(42,40,35,0.06)]"
                  " rounded-lg"
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
                  "w-full p-6 bg-white border border-[#e6e1d4] shadow-[0_1px_3px_rgba(42,40,35,0.06)]"
                  " rounded-lg"
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
                  "w-full p-6 bg-white border border-[#e6e1d4] shadow-[0_1px_3px_rgba(42,40,35,0.06)]"
                  " rounded-lg"
              ):
                ui.label("參數設定").classes(
                    "text-lg font-bold text-zinc-900 tracking-wide mb-3"
                )
                ui.label(
                    "這裡的設定目前只存在記憶體中，服務重啟（含 Render "
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
                    "sync-btn px-4 py-2 text-xs rounded-lg"
                )


  # header（含公司切換分頁）已經搬到函式最前面建立，這裡只需要在所有
  # render_xxx 函式都定義好之後，觸發一次「畫面初始內容」即可。
  company_tabs.set_value(ACTIVE_COMPANY_LABEL)
  render_hai_tao_ke_page()


# =============================================================================
# 以下三個是新規劃的App骨架（雲端電商訂單／雲端會計／報表分析），目前都
# 只有「App切換器 + 公司切換器」的架子，公司分頁點下去先顯示佔位畫面。
# 之後陸續把內容搬過來時，是把 render_company_switcher_placeholder() 內部
# handle_company_change() 裡「顯示佔位畫面」那段，換成真正的頁面渲染函式，
# 不用動這個檔案其他地方，也不用重新設計App切換器/公司切換器。
# =============================================================================
@ui.page("/orders")
def cloud_orders_dashboard():
  inject_global_theme_css()
  render_app_switcher("/orders")
  inject_company_tab_css()

  ORDERS_APP_COMPANIES = ("興聖(股)公司", "容鴻(股)公司", "芙萊柏(股)公司")

  def render_orders_company_content(company_name):
    content_container.clear()
    with content_container:
      with ui.column().classes("w-full p-8 max-w-[1600px] mx-auto gap-4"):
        if company_name not in ORDERS_APP_COMPANIES:
          with ui.card().classes(
              "w-full p-16 bg-white border border-[#e6e1d4]"
              " shadow-[0_1px_3px_rgba(42,40,35,0.06)] rounded-lg text-center"
          ):
            ui.label(company_name).classes(
                "text-lg font-bold text-zinc-900 mb-2"
            )
            ui.label(
                "這間公司的訂單出貨/每日出貨目前還在「雲端進銷存」App"
                "裡，請切到那邊查看，尚未搬過來這裡。"
            ).classes("text-sm text-zinc-500")
          return

        ui.label(company_name).classes("text-lg font-bold text-zinc-900")

        with ui.tabs(
            on_change=lambda e: handle_orders_section_change(e.value)
        ).props("dense no-caps").classes("w-full") as orders_section_tabs:
          ui.tab("訂單出貨")
          ui.tab("每日出貨")

        orders_section_body = ui.column().classes("w-full")

        async def handle_orders_section_change(tab_label):
          orders_section_body.clear()
          with orders_section_body:
            with ui.row().classes(
                "w-full items-center gap-2 p-8 justify-center"
            ):
              ui.spinner(size="24px").classes("text-zinc-400")
              ui.label("資料抓取中，請稍候…").classes("text-xs text-zinc-500")
          await asyncio.sleep(0)

          orders_section_body.clear()
          if tab_label == "訂單出貨":
            with orders_section_body:
              render_order_channels_tabs(company_name)
          elif tab_label == "每日出貨":
            daily_shipping_creds = DAILY_SHIPPING_CREDENTIALS.get(company_name)
            with orders_section_body:
              if daily_shipping_creds:
                await render_daily_shipping(
                    daily_shipping_creds[0], daily_shipping_creds[1], company_name,
                )
              else:
                render_section_placeholder("每日出貨")

        orders_section_tabs.set_value("訂單出貨")

  with ui.row().classes(
      "w-full flex flex-nowrap items-center bg-white border-b border-[#e6e1d4]"
      " px-8 py-3 sticky top-[41px] z-40"
  ):
    with ui.tabs(
        on_change=lambda e: render_orders_company_content(e.value)
    ).props("dense no-caps").classes("flex-shrink-0") as orders_company_tabs:
      for i, c in enumerate(COMPANIES):
        ui.tab(c).classes(f"company-tab-{i}")

  content_container = ui.column().classes("w-full")
  orders_company_tabs.set_value("興聖(股)公司")


@ui.page("/accounting")
def cloud_accounting_dashboard():
  inject_global_theme_css()
  render_app_switcher("/accounting")
  render_company_switcher_placeholder("雲端會計")


@ui.page("/analytics")
def sales_analytics_dashboard():
  inject_global_theme_css()
  render_app_switcher("/analytics")
  render_company_switcher_placeholder("報表分析")


ui.run(port=8080, title="興聖集團 A1 智慧進銷存總管理系統", host="0.0.0.0")
