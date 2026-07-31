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
  商品數量多時逐一序列呼叫會很慢，這裡改用多執行緒平行抓取明細。
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

  def fetch_one(item_id):
    detail_url = f"{A1_BASE_URL}/Items/{item_id}"
    try:
      detail_res = requests.get(
          detail_url, headers=headers, timeout=REQUEST_TIMEOUT
      )
      if detail_res.status_code == 200:
        return item_id, detail_res.json()
      print(
          f"取得商品明細失敗 [{item_id}]"
          f" [{detail_res.status_code}]: {detail_res.text}"
      )
    except requests.exceptions.RequestException as e:
      print(f"取得商品明細失敗 [{item_id}]: {e}")
    return item_id, None

  with concurrent.futures.ThreadPoolExecutor(
      max_workers=ITEM_DETAIL_WORKERS
  ) as executor:
    for item_id, detail in executor.map(fetch_one, item_ids):
      if detail:
        items_dict[item_id] = detail

  return items_dict


def fetch_all_a1_inventory():
  """透過 StockBatch API 分頁完整抓取所有商品在各倉庫的庫存資料

  手冊 StockBatch[Post]：每頁固定 100 筆（依 品號+倉庫 組合計算），
  不傳 StartItemID/EndItemID/ItemIDs/WarehouseName 即代表查全部品號、全部倉庫。
  Response 為 {"Data": [...], "More": bool}，More=true 代表還有下一頁。
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

  all_stock_data = []
  pagination = 1
  more_data = True

  # 使用 StockBatch 分頁迴圈抓取全部庫存（含安全上限，避免 More 異常造成無窮迴圈）
  while more_data and pagination <= MAX_STOCK_PAGES:
    url = f"{A1_BASE_URL}/Stock/Batch"
    payload = {"Pagination": pagination}

    try:
      response = requests.post(
          url, json=payload, headers=headers, timeout=REQUEST_TIMEOUT
      )
      if response.status_code == 200:
        res_json = response.json()
        # 實際回傳結構依手冊為主，通常包在 Data 欄位中
        rows = res_json if isinstance(res_json, list) else res_json.get("Data", [])

        for row in rows:
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

        print(f"StockBatch 第 {pagination} 頁抓取完成，累計 {len(all_stock_data)} 筆")

        # 檢查是否還有下一頁（根據手冊 StockBatch 回傳格式）
        more_data = (
            res_json.get("More", False) if isinstance(res_json, dict) else False
        )
        pagination += 1
      else:
        print(f"StockBatch 抓取失敗 [{response.status_code}]: {response.text}")
        break
    except requests.exceptions.RequestException as e:
      print(f"StockBatch 請求異常: {e}")
      break

  if pagination > MAX_STOCK_PAGES:
    print(f"警告：StockBatch 已達分頁安全上限 {MAX_STOCK_PAGES} 頁，資料可能未抓取完整")

  # 防呆機制：若 API 無資料或連線失敗，回傳範例資料
  if not all_stock_data:
    return get_mock_data(), warehouses, list(categories_map.values())

  return pd.DataFrame(all_stock_data), warehouses, list(categories_map.values())


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
        </style>
    """)

  def handle_sync():
    df, whs, cats = fetch_all_a1_inventory()
    app_state["df"] = df
    if whs:
      app_state["warehouses"] = whs
    if cats:
      app_state["categories"] = cats

    # 更新下拉選單選項
    wh_select.options = ["全部倉庫"] + app_state["warehouses"]
    cat_select.options = ["全部分類"] + app_state["categories"]

    ui.notify(
        f"已成功從鼎新 A1 API 同步 {COMPANY_NAME} 最新庫存資料！"
        f"（{datetime.now().strftime('%H:%M:%S')}）",
        color="positive",
    )
    update_table()

  with ui.row().classes(
      "w-full items-center justify-between bg-white border-b border-[#e2e1dc]"
      " px-8 py-4 sticky top-0 z-50"
  ):
    ui.label(f"興聖集團｜{COMPANY_NAME}｜A1 智慧庫存總管理系統").classes(
        "text-base font-black tracking-wider"
    )
    ui.button("同步 A1 最新庫存", on_click=handle_sync).classes(
        "sync-btn px-4 py-2 text-xs rounded-none"
    )

  with ui.column().classes("w-full p-8 max-w-[1600px] mx-auto"):
    with ui.card().classes(
        "w-full p-6 bg-white border border-[#e2e1dc] shadow-none rounded-none"
    ):
      with ui.row().classes(
          "w-full items-center justify-between mb-6 gap-4 flex-wrap"
      ):
        ui.label("倉庫即時庫存總表").classes(
            "text-lg font-bold text-zinc-900 tracking-wide"
        )

        with ui.row().classes("items-center gap-3 flex-wrap"):
          wh_options = ["全部倉庫"] + app_state["warehouses"]
          wh_select = ui.select(options=wh_options, value="全部倉庫").classes(
              "bg-[#f7f6f2] text-zinc-900 rounded-none px-3 py-1 text-xs"
              " font-bold border border-[#e2e1dc]"
          )

          cat_options = ["全部分類"] + app_state["categories"]
          cat_select = ui.select(options=cat_options, value="全部分類").classes(
              "bg-[#f7f6f2] text-zinc-900 rounded-none px-3 py-1 text-xs"
              " font-bold border border-[#e2e1dc]"
          )

          search_input = ui.input(
              placeholder="輸入品號或品名關鍵字..."
          ).classes("w-64 text-xs")

      table_container = ui.column().classes("w-full")

      def update_table():
        table_container.clear()
        df = app_state["df"].copy()

        # 1. 倉庫篩選
        if wh_select.value and wh_select.value != "全部倉庫":
          df = df[df["倉庫名稱"] == wh_select.value]

        # 2. 分類篩選
        if cat_select.value and cat_select.value != "全部分類":
          df = df[df["商品分類"] == cat_select.value]

        # 3. 關鍵字搜尋
        keyword = search_input.value.strip()
        if keyword:
          mask = df["品號"].astype(str).str.contains(
              keyword, case=False, na=False
          ) | df["品名"].astype(str).str.contains(
              keyword, case=False, na=False
          )
          df = df[mask]

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

      wh_select.on_value_change(lambda e: update_table())
      cat_select.on_value_change(lambda e: update_table())
      search_input.on_value_change(lambda e: update_table())

      update_table()


ui.run(port=8080, title=f"{COMPANY_NAME} ERP 庫存系統", host="0.0.0.0")
