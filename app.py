from datetime import datetime
from nicegui import ui
import pandas as pd
import requests

# -------------------------------------------------------------------------
# 1. 鼎新 A1 API 串接與全量資料自動抓取
# -------------------------------------------------------------------------
A1_BASE_URL = "http://a1external.digiwin.com"
API_KEY = "YOUR_A1_API_KEY"

WAREHOUSES = [
    {"id": "WH01", "name": "食品廠鳳仁倉"},
    {"id": "WH02", "name": "即期品/報廢倉"},
    {"id": "WH03", "name": "供應商-原料倉"},
    {"id": "WH04", "name": "永福倉"},
    {"id": "WH05", "name": "北仁街辦公室"},
    {"id": "WH06", "name": "小琉球現場"},
    {"id": "WH07", "name": "供應商-耗材倉"},
]

CATEGORIES = [
    "(海濤客)_成品11",
    "(海濤客)_原料21",
    "(海濤客)_物料31",
    "(海濤客)_組合品61",
    "(海濤客)_費用71",
    "(海濤客)_代工含料81",
    "(海濤客)_限定組合99",
]


def fetch_all_a1_inventory():
  """透過迴圈自動向 A1 API 抓取所有倉庫與所有分類的即時庫存資料"""
  headers = {
      "Content-Type": "application/json",
      "Authorization": f"Bearer {API_KEY}",
  }

  all_data = []

  for wh in WAREHOUSES:
    for cat in CATEGORIES:
      try:
        # 實際串接時請在此處發送 API 請求
        pass
      except Exception as e:
        print(f"抓取倉庫 {wh['name']} 分類 {cat} 失敗: {e}")

  # 測試防呆數據
  if not all_data:
    all_data = [
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
    ]

  return pd.DataFrame(all_data)


# 建立全域狀態物件存放 DataFrame
app_state = {"df": fetch_all_a1_inventory()}

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
    app_state["df"] = fetch_all_a1_inventory()
    ui.notify("已成功從鼎新 A1 API 抓取所有倉庫與分類資料！", color="positive")
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
          wh_options = ["全部倉庫"] + [w["name"] for w in WAREHOUSES]
          wh_select = ui.select(options=wh_options, value="全部倉庫").classes(
              "bg-[#f7f6f2] text-zinc-900 rounded-none px-3 py-1 text-xs"
              " font-bold border border-[#e2e1dc]"
          )

          cat_options = ["全部分類"] + CATEGORIES
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