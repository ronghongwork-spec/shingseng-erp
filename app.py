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


def fetch_items_map(token):
  """取得商品詳細資料（對應品名、分類、單位、平均成本等）

  手冊 Items[Get] 無傳入商品代號時，只回傳 ID/Name，要拿到 CategoryID、
  UnitName、StdPurPrice 等完整欄位，必須逐筆呼叫 Items/{ItemID}。
  商品數量多時逐一序列呼叫會很慢，這裡改用多執行緒平行抓取明細，
  並針對單筆失敗加入重試，避免暫時性網路錯誤讓某些商品被靜默漏掉。
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
          return item_id, detail_res.json()
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
    return get_mock_data(), [], []

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
    })

  print(
      f"庫存資料彙整完成：StockBatch 回傳 {len(covered_item_ids)} 個不同品號，"
      f"另補上 {len(missing_items)} 個從未有庫存異動的品號，"
      f"總計 {len(all_stock_data)} 列（品號 x 倉庫）"
  )

  # 防呆機制：若 API 無資料或連線失敗，回傳範例資料
  if not all_stock_data:
    return get_mock_data(), warehouses, list(categories_map.values())

  return pd.DataFrame(all_stock_data), warehouses, list(categories_map.values())


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
      },
      {
          "倉庫名稱": "食品廠鳳仁倉",
          "商品分類": "(海濤客)_成品11",
          "品號": "011101180002",
          "品名": "醬料_飛魚卵XO醬",
          "單位": "罐",
          "庫存數量": 258.00,
          "平均成本": 90.50,
      },
      {
          "倉庫名稱": "永福倉",
          "商品分類": "(海濤客)_原料21",
          "品號": "011102000001",
          "品名": "原料_干貝散裝",
          "單位": "公斤",
          "庫存數量": 500.00,
          "平均成本": 150.00,
      },
      {
          "倉庫名稱": "小琉球現場",
          "商品分類": "(海濤客)_限定組合99",
          "品號": "011109900001",
          "品名": "現場限定澎湃禮盒",
          "單位": "組",
          "庫存數量": 45.00,
          "平均成本": 680.00,
      },
  ])


def get_mock_bom_data(item_id):
  """BOM（商品組合）範例資料，待確認 A1 正式 API 端點後即可自動改為即時資料"""
  return [
      {"組成品號": "011102000001", "組成品名": "原料_干貝散裝", "用量": 0.05, "單位": "公斤"},
      {"組成品號": "011101180002", "組成品名": "醬料_飛魚卵XO醬", "用量": 1, "單位": "罐"},
  ]


def fetch_item_bom(token, item_id):
  """查詢單一商品的 BOM（組合品用量明細）

  TODO：目前「鼎新 A1 商務應用雲 POS API 串接手冊」中 BOM／商品組合對應的
  正式端點尚待確認，此處先以常見 RESTful 慣例試打 /Items/{ItemID}/BOM。
  正式串接前請對照手冊確認實際路徑與回傳欄位名稱（常見命名如 ItemBOM、
  Combination、ItemComponents 等），呼叫失敗時頁面會自動退回範例資料。
  """
  url = f"{A1_BASE_URL}/Items/{item_id}/BOM"
  headers = {"Authorization": token}
  try:
    response = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
    if response.status_code == 200:
      data = response.json()
      rows = data if isinstance(data, list) else data.get("Data", []) if isinstance(data, dict) else []
      return True, rows
    return False, f"[{response.status_code}] {response.text}"
  except requests.exceptions.RequestException as e:
    return False, str(e)


# -------------------------------------------------------------------------
# 興聖集團旗下分公司清單（右上角切換用）
# 目前僅「海濤客食品工業(股)公司」已完成 A1 API 串接，其餘分公司頁面預留、
# 之後陸續串接時只要比照 render_hai_tao_ke_page() 的寫法為它們各自建立
# render_xxx_page() 即可。
# -------------------------------------------------------------------------
COMPANIES = ["興聖(股)公司", "海濤客食品工業(股)公司", "容鴻(股)公司", "芙萊柏(股)公司"]
ACTIVE_COMPANY_LABEL = "海濤客食品工業(股)公司"


# 初始化全域狀態
initial_df, initial_whs, initial_cats = fetch_all_a1_inventory()
app_state = {
    "df": initial_df,
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
        df, whs, cats = fetch_all_a1_inventory()
        app_state["df"] = df
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
          tab_bom = ui.tab("BOM表（商品組合）")

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

                # 依品號彙整（同一品號在不同倉庫的庫存加總，
                # 品名/分類/單位/平均成本取第一筆即可，均來自商品主檔）
                catalog = df.groupby("品號", as_index=False).agg({
                    "品名": "first",
                    "商品分類": "first",
                    "單位": "first",
                    "平均成本": "first",
                    "庫存數量": "sum",
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
                      # 目前 A1 商品明細（Items/{ID}）未回傳圖片欄位，
                      # 這裡先以品名首字當作預留圖示，待確認 A1 是否有
                      # 圖片欄位、或改由人工上傳商品照片後再替換。
                      with ui.row().classes(
                          "w-full items-center justify-center mb-2"
                      ):
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

          # ---------------- 頁籤 3：BOM表（商品組合） ----------------
          with ui.tab_panel(tab_bom):
            with ui.card().classes(
                "w-full p-6 bg-white border border-[#e2e1dc] shadow-none"
                " rounded-none"
            ):
              ui.label("查詢商品的 BOM（組合品用量明細）").classes(
                  "text-sm font-bold text-zinc-700 mb-3"
              )

              with ui.row().classes("items-center gap-3 flex-wrap mb-2"):
                bom_item_options = {
                    f"{row['品號']}｜{row['品名']}": row["品號"]
                    for row in app_state["df"]
                    .drop_duplicates("品號")
                    .to_dict("records")
                }
                bom_select = ui.select(
                    options=list(bom_item_options.keys()),
                    with_input=True,
                    label="選擇商品",
                ).classes("w-96 text-xs")

                def handle_query_bom():
                  if not bom_select.value:
                    ui.notify("請先選擇商品", color="warning")
                    return
                  item_id = bom_item_options.get(bom_select.value)
                  token = get_a1_token()
                  bom_rows = None
                  if token:
                    ok, data = fetch_item_bom(token, item_id)
                    if ok and data:
                      bom_rows = data

                  if not bom_rows:
                    bom_rows = get_mock_bom_data(item_id)
                    bom_result_label.text = (
                        "（尚未取得 A1 正式 BOM 資料，以下為範例資料，"
                        "待確認 API 端點後將自動改為即時資料）"
                    )
                  else:
                    bom_result_label.text = (
                        f"A1 即時 BOM 資料，共 {len(bom_rows)} 筆組成"
                    )

                  bom_table_container.clear()
                  with bom_table_container:
                    ui.table(
                        columns=[
                            {
                                "name": "組成品號",
                                "label": "組成品號",
                                "field": "組成品號",
                                "align": "left",
                            },
                            {
                                "name": "組成品名",
                                "label": "組成品名",
                                "field": "組成品名",
                                "align": "left",
                            },
                            {"name": "用量", "label": "用量", "field": "用量"},
                            {"name": "單位", "label": "單位", "field": "單位"},
                        ],
                        rows=bom_rows,
                    ).classes("w-full")

                ui.button("查詢 BOM", on_click=handle_query_bom).classes(
                    "sync-btn px-3 py-1 text-xs rounded-none"
                )

              bom_result_label = ui.label().classes(
                  "text-xs text-zinc-500 mb-3"
              )
              bom_table_container = ui.column().classes("w-full")

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
