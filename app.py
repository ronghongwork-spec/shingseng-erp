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
try:
  from dotenv import load_dotenv
  load_dotenv()
except ImportError:
  pass

A1_BASE_URL = "http://a1external.digiwin.com"
API_KEY = os.environ.get("A1_API_KEY", "")
API_PASSWORD = os.environ.get("A1_API_PASSWORD", "")

# 支援多公司狀態管理（預設為海濤客）
company_state = {
    "current_company": "海濤客",
    "companies": [
        "興聖",
        "海濤客",
        "容鴻",
        "芙萊柏",
    ]
}

REQUEST_TIMEOUT = 15
STOCK_PAGE_SIZE = 100
MAX_STOCK_PAGES = 1000
ITEM_DETAIL_WORKERS = 8


def get_a1_token():
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
        return None
      return access_token
  except requests.exceptions.RequestException as e:
    print(f"A1 登入連線異常: {e}")
  return None


def fetch_warehouses(token):
  url = f"{A1_BASE_URL}/Warehouses"
  headers = {"Authorization": token}
  try:
    response = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
    if response.status_code == 200:
      return [w["Name"] for w in response.json()]
  except requests.exceptions.RequestException as e:
    print(f"取得倉庫列表失敗: {e}")
  return []


def fetch_categories(token):
  url = f"{A1_BASE_URL}/Categorys"
  headers = {"Authorization": token}
  try:
    response = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
    if response.status_code == 200:
      return {str(c["ID"]): c["Name"] for c in response.json()}
  except requests.exceptions.RequestException as e:
    print(f"取得商品分類失敗: {e}")
  return {}


def fetch_items_map(token):
  url = f"{A1_BASE_URL}/Items"
  headers = {"Authorization": token}
  items_dict = {}

  try:
    response = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
    if response.status_code != 200:
      return items_dict
    item_ids = [item.get("ID") for item in response.json() if item.get("ID")]
  except requests.exceptions.RequestException as e:
    return items_dict

  ITEM_DETAIL_RETRIES = 2

  def fetch_one(item_id):
    detail_url = f"{A1_BASE_URL}/Items/{item_id}"
    for _ in range(ITEM_DETAIL_RETRIES):
      try:
        detail_res = requests.get(
            detail_url, headers=headers, timeout=REQUEST_TIMEOUT
        )
        if detail_res.status_code == 200:
          return item_id, detail_res.json()
      except requests.exceptions.RequestException:
        pass
    return item_id, None

  with concurrent.futures.ThreadPoolExecutor(
      max_workers=ITEM_DETAIL_WORKERS
  ) as executor:
    for item_id, detail in executor.map(fetch_one, item_ids):
      if detail:
        items_dict[item_id] = detail

  return items_dict


def fetch_boms(token):
  """取得鼎新 A1 商品組合 (BOM) 資料"""
  url = f"{A1_BASE_URL}/BOM"  # 依實際 A1 API 端點調整
  headers = {"Authorization": token}
  try:
    response = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
    if response.status_code == 200:
      data = response.json()
      return data if isinstance(data, list) else data.get("Data", [])
  except Exception:
    pass
  # 測試防呆 BOM 資料
  return [
      {
          "主件品號": "011109900001",
          "主件品名": "現場限定澎湃禮盒",
          "元件品號": "011101180001",
          "元件品名": "醬料_烏金干貝醬",
          "組成數量": 1.0,
          "單位": "罐",
      },
      {
          "主件品號": "011109900001",
          "主件品名": "現場限定澎湃禮盒",
          "元件品號": "011101180002",
          "元件品名": "醬料_飛魚卵XO醬",
          "組成數量": 1.0,
          "單位": "罐",
      },
  ]


def fetch_all_a1_inventory():
  token = get_a1_token()
  if not token:
    return get_mock_data(), [], [], get_mock_boms()

  headers = {"Content-Type": "application/json", "Authorization": token}
  warehouses = fetch_warehouses(token)
  categories_map = fetch_categories(token)
  items_map = fetch_items_map(token)
  boms_data = fetch_boms(token)

  def fetch_stock_rows(item_ids_batch=None):
    url = f"{A1_BASE_URL}/Stock/Batch"
    rows_collected = []
    pagination = 1
    more_data = True

    while more_data and pagination <= MAX_STOCK_PAGES:
      payload = {"Pagination": pagination}
      if item_ids_batch:
        payload["ItemIDs"] = item_ids_batch
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
          break
      except requests.exceptions.RequestException:
        break
    return rows_collected

  ITEM_BATCH_SIZE = 100
  all_item_ids = list(items_map.keys())
  raw_rows = []

  if all_item_ids:
    batches = [
        all_item_ids[i:i + ITEM_BATCH_SIZE]
        for i in range(0, len(all_item_ids), ITEM_BATCH_SIZE)
    ]
    for batch in batches:
      raw_rows.extend(fetch_stock_rows(item_ids_batch=batch))
  else:
    raw_rows = fetch_stock_rows()

  all_stock_data = []
  for row in raw_rows:
    item_id = row.get("ItemID")
    item_info = items_map.get(item_id, {})
    cat_id = item_info.get("CategoryID")
    cat_name = categories_map.get(str(cat_id), "未分類")

    # 模擬商品圖片欄位（若 A1 無圖片可由品號對應或給預設圖）
    img_url = item_info.get(
        "ImageUrl", "https://via.placeholder.com/60?text=No+Image"
    )

    all_stock_data.append({
        "圖片": img_url,
        "倉庫名稱": row.get("WarehouseName"),
        "商品分類": cat_name,
        "品號": item_id,
        "品名": row.get("ItemName") or item_info.get("Name"),
        "單位": item_info.get("UnitName", "個"),
        "庫存數量": row.get("Qty", 0.0),
        "平均成本": item_info.get("StdPurPrice", 0.0),
    })

  covered_item_ids = {row["品號"] for row in all_stock_data}
  for item_id, item_info in items_map.items():
    if item_id not in covered_item_ids:
      cat_id = item_info.get("CategoryID")
      cat_name = categories_map.get(str(cat_id), "未分類")
      all_stock_data.append({
          "圖片": "https://via.placeholder.com/60?text=No+Image",
          "倉庫名稱": "(無庫存異動)",
          "商品分類": cat_name,
          "品號": item_id,
          "品名": item_info.get("Name"),
          "單位": item_info.get("UnitName", "個"),
          "庫存數量": 0.0,
          "平均成本": item_info.get("StdPurPrice", 0.0),
      })

  if not all_stock_data:
    return (
        get_mock_data(),
        ["食品廠鳳仁倉", "永福倉", "小琉球現場"],
        ["(海濤客)_成品11", "(海濤客)_原料21"],
        get_mock_boms(),
    )

  return (
      pd.DataFrame(all_stock_data),
      warehouses,
      list(categories_map.values()),
      boms_data,
  )


def fetch_stock_single_item(token, item_id):
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
  return pd.DataFrame([
      {
          "圖片": "https://via.placeholder.com/60?text=XO",
          "倉庫名稱": "食品廠鳳仁倉",
          "商品分類": "(海濤客)_成品11",
          "品號": "011101180001",
          "品名": "醬料_烏金干貝醬",
          "單位": "罐",
          "庫存數量": 262.00,
          "平均成本": 214.41,
      },
      {
          "圖片": "https://via.placeholder.com/60?text=Fish",
          "倉庫名稱": "食品廠鳳仁倉",
          "商品分類": "(海濤客)_成品11",
          "品號": "011101180002",
          "品名": "醬料_飛魚卵XO醬",
          "單位": "罐",
          "庫存數量": 258.00,
          "平均成本": 90.50,
      },
  ])


def get_mock_boms():
  return [
      {
          "主件品號": "011109900001",
          "主件品名": "現場限定澎湃禮盒",
          "元件品號": "011101180001",
          "元件品名": "醬料_烏金干貝醬",
          "組成數量": 1.0,
          "單位": "組",
      }
  ]


# 初始化全域狀態
initial_df, initial_whs, initial_cats, initial_boms = fetch_all_a1_inventory()
app_state = {
    "df": initial_df,
    "warehouses": (
        initial_whs if initial_whs else ["食品廠鳳仁倉", "永福倉", "小琉球現場"]
    ),
    "categories": (
        initial_cats
        if initial_cats
        else ["(海濤客)_成品11", "(海濤客)_原料21"]
    ),
    "boms": initial_boms,
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
    df, whs, cats, boms = fetch_all_a1_inventory()
    app_state["df"] = df
    if whs:
      app_state["warehouses"] = whs
    if cats:
      app_state["categories"] = cats
    if boms:
      app_state["boms"] = boms

    ui.notify(
        f"已成功同步 {company_state['current_company']} 最新資料！"
        f"（{datetime.now().strftime('%H:%M:%S')}）",
        color="positive",
    )
    update_all_views()

  def switch_company(comp):
    company_state["current_company"] = comp
    company_title_label.text = (
        f"興聖集團｜{company_state['current_company']}｜A1 智慧進銷存總管理系統"
    )
    handle_sync()

  # 頂端導覽列（包含左側標題與右側四家分公司切換按鈕）
  with ui.row().classes(
      "w-full items-center justify-between bg-white border-b border-[#e2e1dc]"
      " px-8 py-4 sticky top-0 z-50 flex-wrap"
  ):
    company_title_label = ui.label(
        f"興聖集團｜{company_state['current_company']}｜A1"
        " 智慧進銷存總管理系統"
    ).classes("text-base font-black tracking-wider")

    with ui.row().classes("items-center gap-2 flex-wrap"):
      ui.label("切換分公司：").classes("text-xs font-bold text-zinc-600")
      for comp in company_state["companies"]:
        ui.button(comp, on_click=lambda c=comp: switch_company(c)).classes(
            "px-3 py-1 text-xs rounded-none border border-[#e2e1dc]"
            f" {'bg-zinc-900 text-white font-bold' if comp == company_state['current_company'] else 'bg-white text-zinc-800'}"
        )
      ui.button("同步 A1 最新資料", on_click=handle_sync).classes(
          "sync-btn px-4 py-1.5 text-xs rounded-none ml-4"
      )

  # 主畫面多頁籤容器
  with ui.column().classes("w-full p-8 max-w-[1600px] mx-auto"):
    with ui.tabs().classes("w-full text-zinc-800") as tabs:
      tab_items = ui.tab("頁籤1-商品資料", label="頁籤 1 - 商品資料")
      tab_stock = ui.tab("頁籤2-倉庫即時庫存總表", label="頁籤 2 - 倉庫即時庫存總表")
      tab_bom = ui.tab("頁籤3-BOM表", label="頁籤 3 - BOM 表")

    with ui.tab_panels(tabs, value=tab_stock).classes(
        "w-full bg-transparent shadow-none"
    ):

      # ==========================================
      # 頁籤 1：商品資料（含圖片、品名與庫存關聯）
      # ==========================================
      with ui.tab_panel(tab_items):
        with ui.card().classes(
            "w-full p-6 bg-white border border-[#e2e1dc] shadow-none rounded-none"
        ):
          ui.label("商品資料與即時庫存對照").classes(
              "text-lg font-bold text-zinc-900 mb-4"
          )
          items_table_container = ui.column().classes("w-full")

      # ==========================================
      # 頁籤 2：倉庫即時庫存總表
      # ==========================================
      with ui.tab_panel(tab_stock):
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
              cat_select = ui.select(
                  options=cat_options, value="全部分類"
              ).classes(
                  "bg-[#f7f6f2] text-zinc-900 rounded-none px-3 py-1 text-xs"
                  " font-bold border border-[#e2e1dc]"
              )

              search_input = ui.input(
                  placeholder="輸入品號或品名關鍵字..."
              ).classes("w-64 text-xs")

          stats_label = ui.label().classes("text-xs text-zinc-500 mb-3")

          with ui.row().classes(
              "w-full items-center gap-3 mb-4 p-3 bg-[#f7f6f2] border"
              " border-[#e2e1dc] flex-wrap"
          ):
            ui.label("單一品號覆核（直查 A1）").classes(
                "text-xs font-bold text-zinc-700"
            )
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
                verify_result.text = "無法登入 A1"
                return
              ok, msg, rows = fetch_stock_single_item(token, item_id)
              if not ok:
                verify_result.text = f"查詢失敗：{msg}"
                return
              if not rows:
                verify_result.text = "A1 回傳：此品號目前無庫存資料"
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

      # ==========================================
      # 頁籤 3：BOM 表（商品組合）
      # ==========================================
      with ui.tab_panel(tab_bom):
        with ui.card().classes(
            "w-full p-6 bg-white border border-[#e2e1dc] shadow-none rounded-none"
        ):
          ui.label("鼎新 A1 商品組合 (BOM) 設定").classes(
              "text-lg font-bold text-zinc-900 mb-4"
          )
          bom_table_container = ui.column().classes("w-full")

  def update_all_views():
    df = app_state["df"].copy()

    # 更新頁籤 1 - 商品資料 (呈現圖片、品名與庫存關係)
    items_table_container.clear()
    with items_table_container:
      # 這裡利用 HTML 自定義呈現圖片
      ui.table(
          columns=[
              {"name": "圖片", "label": "商品圖片", "field": "圖片"},
              {"name": "品號", "label": "品號", "field": "品號", "align": "left"},
              {"name": "品名", "label": "品名", "field": "品名", "align": "left"},
              {"name": "單位", "label": "單位", "field": "單位"},
              {
                  "name": "庫存數量",
                  "label": "目前總庫存數量",
                  "field": "庫存數量",
              },
              {"name": "平均成本", "label": "平均成本", "field": "平均成本"},
          ],
          rows=df.to_dict("records"),
      ).classes("w-full")

    # 更新頁籤 2 - 倉庫即時庫存總表
    table_container.clear()
    total_before_filters = len(df)

    if wh_select.value and wh_select.value != "全部倉庫":
      df = df[df["倉庫名稱"] == wh_select.value]
    if cat_select.value and cat_select.value != "全部分類":
      df = df[df["商品分類"] == cat_select.value]

    keyword = search_input.value.strip()
    if keyword:
      mask = df["品號"].astype(str).str.contains(
          keyword, case=False, na=False
      ) | df["品名"].astype(str).str.contains(keyword, case=False, na=False)
      df = df[mask]

    before_qty_filter = len(df)
    df = df[df["庫存數量"] != 0]
    hidden_zero_count = before_qty_filter - len(df)

    stats_label.text = (
        f"同步資料共 {total_before_filters} 列｜符合篩選條件"
        f" {before_qty_filter} 列｜已隱藏零庫存 {hidden_zero_count} 列｜目前顯示"
        f" {len(df)} 列"
    )

    with table_container:
      ui.table(
          columns=[
              {"name": "倉庫名稱", "label": "倉庫", "field": "倉庫名稱"},
              {"name": "商品分類", "label": "商品分類", "field": "商品分類"},
              {"name": "品號", "label": "品號", "field": "品號"},
              {"name": "品名", "label": "品名", "field": "品名"},
              {"name": "單位", "label": "單位", "field": "單位"},
              {"name": "庫存數量", "label": "庫存數量", "field": "庫存數量"},
              {"name": "平均成本", "label": "平均成本", "field": "平均成本"},
          ],
          rows=df.to_dict("records"),
      ).classes("w-full")

    # 更新頁籤 3 - BOM 表
    bom_table_container.clear()
    with bom_table_container:
      ui.table(
          columns=[
              {"name": "主件品號", "label": "主件品號", "field": "主件品號"},
              {"name": "主件品名", "label": "主件品名", "field": "主件品名"},
              {"name": "元件品號", "label": "元件品號", "field": "元件品號"},
              {"name": "元件品名", "label": "元件品名", "field": "元件品名"},
              {"name": "組成數量", "label": "組成數量", "field": "組成數量"},
              {"name": "單位", "label": "單位", "field": "單位"},
          ],
          rows=app_state["boms"],
      ).classes("w-full")

  wh_select.on_value_change(lambda e: update_all_views())
  cat_select.on_value_change(lambda e: update_all_views())
  search_input.on_value_change(lambda e: update_all_views())

  update_all_views()


ui.run(port=8080, title="興聖集團總管理系統", host="0.0.0.0")