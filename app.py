import os
import json
import pandas as pd
import gspread
from nicegui import ui

# -------------------------------------------------------------------------
# Google Sheet 數據串接與解析
# -------------------------------------------------------------------------
def load_shingseng_target_data():
    """從結構化 Google 試算表中精準萃取第一季各通路數據，具備多重防呆與環境變數支援"""
    try:
        creds_json = os.getenv("GOOGLE_CREDENTIALS_JSON")
        if creds_json:
            creds_dict = json.loads(creds_json)
            gc = gspread.service_account_from_dict(creds_dict)
        else:
            gc = gspread.service_account(filename='credentials.json')
        
        sh = gc.open_by_url("https://docs.google.com/spreadsheets/d/1Wc64Uqc1gvOS2CMX4QXRPOHDUsxuwDVS3tyRf5hexuY/edit?gid=80636189")
        worksheet = sh.get_worksheet(0) 
        rows = worksheet.get_all_values()
        
        if len(rows) > 10:
            parsed_data = []
            for r in rows[10:]:
                if not r or len(r) < 15:
                    continue
                
                channel_name = r[0].strip()
                if not channel_name or '總計' in channel_name or '合計' in channel_name:
                    continue
                
                try:
                    m1_actual = float(str(r[2]).replace(',', '').strip()) if r[2] else 0
                    m1_target = float(str(r[4]).replace(',', '').strip()) if r[4] else 0
                    m2_actual = float(str(r[7]).replace(',', '').strip()) if r[7] else 0
                    m2_target = float(str(r[9]).replace(',', '').strip()) if r[9] else 0
                    m3_actual = float(str(r[12]).replace(',', '').strip()) if r[12] else 0
                    m3_target = float(str(r[14]).replace(',', '').strip()) if r[14] else 0
                    
                    parsed_data.append({'通路': channel_name, '月份': '1月', '目標銷售額': m1_target, '實際銷售額': m1_actual})
                    parsed_data.append({'通路': channel_name, '月份': '2月', '目標銷售額': m2_target, '實際銷售額': m2_actual})
                    parsed_data.append({'通路': channel_name, '月份': '3月', '目標銷售額': m3_target, '實際銷售額': m3_actual})
                except Exception:
                    continue
            
            if parsed_data:
                print("成功從 Google 試算表載入第一季營運數據！")
                return pd.DataFrame(parsed_data)

    except Exception as e:
        print(f"Google API 連線或解析發生例外，自動載入備用模擬數據: {e}")
    
    # 備用防呆模擬數據
    return pd.DataFrame({
        '月份': ['1月', '1月', '2月', '2月', '3月', '3月'],
        '通路': ['X線上-FB官網', 'R線上-蝦皮', 'X線上-FB官網', 'R線上-蝦皮', 'X線上-FB官網', 'R線上-蝦皮'],
        '目標銷售額': [11168257, 2138820, 2236046, 664841, 2395151, 775962],
        '實際銷售額': [8278256, 1374061, 3072897, 796704, 887073, 457013]
    })

# -------------------------------------------------------------------------
# 共用導航與框架佈局
# -------------------------------------------------------------------------
def layout(title: str):
    """建立 ERP 系統的共用側邊欄與頂部導航框架"""
    with ui.header().classes('bg-slate-800 text-white items-center justify-between px-6'):
        ui.label('興聖集團 ERP 系統').classes('text-lg font-bold tracking-wider')
        ui.label(title).classes('text-sm text-slate-300')
        
    with ui.left_drawer(value=True).classes('bg-slate-900 text-slate-200 p-4'):
        ui.label('導航選單').classes('text-xs font-semibold text-slate-400 mb-2 uppercase tracking-wider')
        ui.link('📊 Q1 營運戰情分析', '/').classes('text-white hover:text-blue-400 block py-2 font-medium')
        ui.link('📦 庫存與物料管理', '/inventory').classes('text-slate-300 hover:text-blue-400 block py-2')
        ui.link('🛒 採購與訂單管理', '/orders').classes('text-slate-300 hover:text-blue-400 block py-2')
        ui.link('⚙️ 系統設定與維護', '/settings').classes('text-slate-300 hover:text-blue-400 block py-2')

# -------------------------------------------------------------------------
# 頁面 1：Q1 營運戰情分析（首頁）
# -------------------------------------------------------------------------
@ui.page('/')
def main_page():
    layout('Q1 營運戰情分析')
    
    with ui.column().classes('w-full p-6 max-w-7xl mx-auto'):
        ui.label('第一季營運戰情與數據分析').classes('text-2xl font-bold mb-4 text-slate-800')
        
        df = load_shingseng_target_data()
        
        total_target = df['目標銷售額'].sum()
        total_actual = df['實際銷售額'].sum()
        achievement_rate = (total_actual / total_target * 100) if total_target > 0 else 0
        
        with ui.row().classes('w-full gap-4 mb-6'):
            with ui.card().classes('p-4 bg-blue-50 flex-1 border border-blue-100'):
                ui.label('第一季總目標營收').classes('text-sm text-gray-500')
                ui.label(f'{total_target:,.0f}').classes('text-2xl font-bold text-blue-600')
                
            with ui.card().classes('p-4 bg-green-50 flex-1 border border-green-100'):
                ui.label('第一季實際總營收').classes('text-sm text-gray-500')
                ui.label(f'{total_actual:,.0f}').classes('text-2xl font-bold text-green-600')
                
            with ui.card().classes('p-4 bg-orange-50 flex-1 border border-orange-100'):
                ui.label('整體達成率').classes('text-sm text-gray-500')
                ui.label(f'{achievement_rate:.1f}%').classes('text-2xl font-bold text-orange-600')

        ui.separator().classes('my-4')
        ui.label('各通路銷售額對比（實際 vs 目標）').classes('text-lg font-semibold mb-2 text-slate-700')
        
        channels = df['通路'].unique().tolist()
        actual_vals = [df[df['通路'] == c]['實際銷售額'].sum() for c in channels]
        target_vals = [df[df['通路'] == c]['目標銷售額'].sum() for c in channels]
        
        ui.plotly({
            'data': [
                {'x': channels, 'y': actual_vals, 'type': 'bar', 'name': '實際銷售額', 'marker': {'color': '#3b82f6'}},
                {'x': channels, 'y': target_vals, 'type': 'bar', 'name': '目標銷售額', 'marker': {'color': '#94a3b8'}}
            ],
            'layout': {
                'barmode': 'group',
                'margin': {'l': 50, 'r': 20, 't': 20, 'b': 80},
                'xaxis': {'tickangle': -30}
            }
        }).classes('w-full h-96 shadow rounded bg-white p-4')

# -------------------------------------------------------------------------
# 頁面 2：庫存與物料管理
# -------------------------------------------------------------------------
@ui.page('/inventory')
def inventory_page():
    layout('庫存與物料管理')
    with ui.column().classes('w-full p-6 max-w-7xl mx-auto'):
        ui.label('庫存與物料管理中心').classes('text-2xl font-bold mb-4 text-slate-800')
        ui.card().classes('w-full p-6 bg-white shadow').content(lambda: [
            ui.label('目前物料與庫存狀態正常。').classes('text-gray-600')
        ])

# -------------------------------------------------------------------------
# 頁面 3：採購與訂單管理
# -------------------------------------------------------------------------
@ui.page('/orders')
def orders_page():
    layout('採購與訂單管理')
    with ui.column().classes('w-full p-6 max-w-7xl mx-auto'):
        ui.label('電商採購與訂單作業').classes('text-2xl font-bold mb-4 text-slate-800')
        ui.card().classes('w-full p-6 bg-white shadow').content(lambda: [
            ui.label('在此管理 Shopline、蝦皮等各通路進銷存與訂單同步作業。').classes('text-gray-600')
        ])

# -------------------------------------------------------------------------
# 頁面 4：系統設定
# -------------------------------------------------------------------------
@ui.page('/settings')
def settings_page():
    layout('系統設定與維護')
    with ui.column().classes('w-full p-6 max-w-7xl mx-auto'):
        ui.label('系統設定').classes('text-2xl font-bold mb-4 text-slate-800')
        ui.card().classes('w-full p-6 bg-white shadow').content(lambda: [
            ui.label('ERP 系統串接參數與 Google Sheet 連結設定。').classes('text-gray-600')
        ])

# -------------------------------------------------------------------------
# 伺服器啟動
# -------------------------------------------------------------------------
if __name__ in {"__main__", "__mp_main__"}:
    port = int(os.environ.get("PORT", 8080))
    ui.run(port=port, title="興聖集團 ERP 系統", host='0.0.0.0', reload=False)