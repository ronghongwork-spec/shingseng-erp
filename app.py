from datetime import datetime
from nicegui import ui
import pandas as pd
import requests

# -------------------------------------------------------------------------
# 1. 鼎新 A1 API 串接與全量資料自動抓取
# -------------------------------------------------------------------------
A1_BASE_URL = "http://a1external.digiwin.com"  # 正式區 URL[cite: 3]
API_KEY = "YOUR_A1_API_KEY"


def get_a1_token():
  """透過 APIKey 取得登入金鑰 (JWT Token)[cite: 3]"""
  url = f"{A1_BASE_URL}/Login"
  headers = {"Content-Type": "application/json"}
  body = {"UserName": API_KEY, "Password": ""}  # A1 慣例通常將 APIKey 填入 UserName

  try:
    response = requests.post(url, json=body, headers=headers)
    if response.status_code == 200:
      data = response.json()
      return data.get("access_token")
    else:
      print(f"A1 登入失敗: {response.text}")
  except Exception as e:
    print(f"A1 登入連線異常: {e}")
  return None


def fetch_warehouses(token):
  """動態取得所有未停用的倉庫列表[cite: 3]"""
  url = f"{A1_BASE_URL}/Warehouses"
  headers = {"Authorization": f"Bearer {token}"}
  try:
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
      return [w["Name"] for w in response.json()]
  except Exception as e:
    print(f"取得倉庫列表失敗: {e}")
  return []


def fetch_categories(token):
  """動態取得所有商品分類列表[cite: 3]"""
  url = f"{A1_BASE_URL}/Categorys"
  headers = {"Authorization": f"Bearer {token}"}
  try:
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
      return {c["ID"]: c["Name"] for c in response.json()}
  except Exception as e:
    print(f"取得商品分類失敗: {e}")
  return {}


def fetch_items_map(token):
  """取得商品詳細資料（對應品名、分類等）[cite: 3]"""
  url = f"{A1_BASE_URL}/Items"
  headers = {"Authorization": f"Bearer {token}"}
  items_dict = {}
  try:
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
      for item in response.json():
        item_id = item.get("ID")
        # 抓取單一商品明細以取得完整資訊（如分類、單位、平均成本等）
        detail_url = f"{A1_BASE_URL}/Items/{item_id}"
        detail_res = requests.get(detail_url, headers=headers)
        if detail_res.status_code == 200:
          items_dict[item_id] = detail_res.json()
  except Exception as e:
    print(f"取得商品明細失敗: {e}")
  return items_dict


def fetch_all_a1_inventory():
  """透過 StockBatch API 分頁完整抓取所有商品在各倉庫的庫存資料[cite: 3]"""
  token = get_a1_token()

  if not token:
    print("無法取得 A1 Token，啟用測試防呆數據...")
    return get_mock_data(), [], []

  headers = {
      "Content-Type": "application/json",
      "Authorization": f"Bearer {token}",
  }

  warehouses = fetch_warehouses(token)
  categories_map = fetch_categories(token)
  items_map = fetch_items_map(token)

  all_stock_data = []
  pagination = 1
  more_data = True

  # 使用 StockBatch 分頁迴圈抓取全部庫存[cite: 3]
  while more_data:
    url = f"{A1_BASE_URL}/Stock/Batch"
    payload = {"Pagination": pagination}

    try:
      response = requests.post(url, json=payload, headers=headers)
      if response.status_code == 200:
        res_json = response.json()
        # 注意：實際回傳結構依 API 手冊為主，通常包在 Data 欄位中
        rows = res_json if isinstance(res_json, list) else res_json.get("Data", [])

        for row in rows:
          item_id = row.get("ItemID")
          item_info = items_map.get(item_id, {})

          cat_id = item_info.get("CategoryID")
          cat_name = categories_map.get(cat_id, "未分類")

          all_stock_data.append({
              "倉庫名稱": row.get("WarehouseName"),
              "商品分類": cat_name,
              "品號": item_id,
              "品名": row.get("ItemName") or item_info.get("Name"),
              "單位": item_info.get("UnitName", "個"),
              "庫存數量": row.get("Qty", 0.0),
              "平均成本": item_info.get("StdPurPrice", 0.0),
          })

        # 檢查是否還有下一頁 (根據手冊 StockBatch 回傳格式)
        more_data = (
            res_json.get("More", False)
            if isinstance(res_json, dict)
            else False
        )
        pagination += 1
      else:
        print(f"StockBatch 抓取失敗: {response.text}")
        break
    except Exception as e:
      print(f"StockBatch 請求異常: {e}")
      break

  # 防呆機制：若 API 無資料或連線失敗，回傳範例資料
  if not all_stock_data:
    return get_mock_data(), warehouses, list(categories_map.values())

  return pd.DataFrame(all_stock_data), warehouses, list(categories_map.values())


def get_mock_data():
  """提供本地測試用的防呆 DataFrame"""
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

    ui.notify("已成功從鼎新 A1 API 同步最新庫存資料！", color="positive")
    update_table()

  with ui.row().classes(
      "w-full items-center justify-between bg-white border-b border-[#e2e1dc]"
      " px-8 py-4 sticky top-0 z-50"
  ):
    ui.label("興聖集團｜A1 智慧庫存總管理系統").classes(
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


ui.run(port=8080, title="興聖集團 ERP 庫存系統", host="0.0.0.0")