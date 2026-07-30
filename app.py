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
        # 優先檢查是否有透過 Render 環境變數設定金鑰
        creds_json = os.getenv("GOOGLE_CREDENTIALS_JSON")
        if creds_json:
            creds_dict = json.loads(creds_json)
            gc = gspread.service_account_from_dict(creds_dict)
        else:
            # 本地端或根目錄有放置 credentials.json 檔案時使用
            gc = gspread.service_account(filename='credentials.json')
        
        # 開啟指定的 Google 試算表
        sh = gc.open_by_url("https://docs.google.com/spreadsheets/d/1Wc64Uqc1gvOS2CMX4QXRPOHDUsxuwDVS3tyRf5hexuY/edit?gid=80636189")
        worksheet = sh.get_worksheet(0) 
        
        # 抓取整張表的原始二維資料陣列
        rows = worksheet.get_all_values()
        
        if len(rows) > 10:
            parsed_data = []
            
            # 依照你的表格結構，從第 10 行（索引 9）開始讀取各通路明細
            for r in rows[10:]:
                if not r or len(r) < 15:
                    continue
                
                channel_name = r[0].strip() # 第一欄是通路名稱
                if not channel_name or '總計' in channel_name or '合計' in channel_name:
                    continue
                
                try:
                    # 依據你的試算表欄位索引抓取：
                    # 1月實際: 索引 2, 1月2026目標: 索引 4
                    # 2月實際: 索引 7, 2月2026目標: 索引 9
                    # 3月實際: 索引 12, 3月2026目標: 索引 14
                    m1_actual = float(str(r[2]).replace(',', '').strip()) if r[2] else 0
                    m1_target = float(str(r[4]).replace(',', '').strip()) if r[4] else 0
                    
                    m2_actual = float(str(r[7]).replace(',', '').strip()) if r[7] else 0
                    m2_target = float(str(r[9]).replace(',', '').strip()) if r[9] else 0
                    
                    m3_actual = float(str(r[12]).replace(',', '').strip()) if r[12] else 0
                    m3_target = float(str(r[14]).replace(',', '').strip()) if r[14] else 0
                    
                    # 組合為標準清單
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
    
    # 備用防呆模擬數據（確保 API 異常時網頁依然能正常呈現圖表）
    return pd.DataFrame({
        '月份': ['1月', '1月', '2月', '2月', '3月', '3月'],
        '通路': ['X線上-FB官網', 'R線上-蝦皮', 'X線上-FB官網', 'R線上-蝦皮', 'X線上-FB官網', 'R線上-蝦皮'],
        '目標銷售額': [11168257, 2138820, 2236046, 664841, 2395151, 775962],
        '實際銷售額': [8278256, 1374061, 3072897, 796704, 887073, 457013]
    })

# -------------------------------------------------------------------------
# NiceGUI 網頁介面與圖表建構
# -------------------------------------------------------------------------
@ui.page('/')
def main_page():
    ui.label('興聖集團 - 第一季營運戰情與數據分析').classes('text-2xl font-bold mb-4')
    
    # 載入資料
    df = load_shingseng_target_data()
    
    # 統計看板數值
    total_target = df['目標銷售額'].sum()
    total_actual = df['實際銷售額'].sum()
    achievement_rate = (total_actual / total_target * 100) if total_target > 0 else 0
    
    with ui.row().classes('w-full gap-4'):
        with ui.card().classes('p-4 bg-blue-50'):
            ui.label('第一季總目標營收').classes('text-sm text-gray-500')
            ui.label(f'{total_target:,.0f}').classes('text-xl font-bold text-blue-600')
            
        with ui.card().classes('p-4 bg-green-50'):
            ui.label('第一季實際總營收').classes('text-sm text-gray-500')
            ui.label(f'{total_actual:,.0f}').classes('text-xl font-bold text-green-600')
            
        with ui.card().classes('p-4 bg-orange-50'):
            ui.label('整體達成率').classes('text-sm text-gray-500')
            ui.label(f'{achievement_rate:.1f}%').classes('text-xl font-bold text-orange-600')
        ])

    ui.separator().classes('my-6')
    
    ui.label('各通路銷售額對比（實際 vs 目標）').classes('text-lg font-semibold mb-2')
    
    # 準備 Plotly 圖表資料
    channels = df['通路'].unique().tolist()
    actual_vals = [df[df['通路'] == c]['實際銷售額'].sum() for c in channels]
    target_vals = [df[df['通路'] == c]['目標銷售額'].sum() for c in channels]
    
    chart_data = {
        'тивных': channels,
        'antic': {
            'type': 'bar',
            'title': '通路績效表現',
            'xAxis': {'categories': channels},
            'series': [
                {'name': '實際銷售額', 'data': actual_vals},
                {'name': '目標銷售額', 'data': target_vals}
            ]
        }
    }
    
    # 使用 Plotly 長條圖展示
    ui.plotly({
        'data': [
            {'x': channels, 'y': actual_vals, 'type': 'bar', 'name': '實際銷售額', 'marker': {'color': '#3b82f6'}},
            {'x': channels, 'y': target_vals, 'type': 'bar', 'name': '目標銷售額', 'marker': {'color': '#94a3b8'}}
        ],
        'layout': {
            'barmode': 'group',
            'margin': {'l': 40, 'r': 40, 't': 20, 'b': 80},
            'xaxis': {'tickangle': -45}
        }
    }).classes('w-full h-96')

# -------------------------------------------------------------------------
# 伺服器啟動（對應 Render 雲端環境埠號）
# -------------------------------------------------------------------------
if __name__ in {"__main__", "__mp_main__"}:
    port = int(os.environ.get("PORT", 8080))
    ui.run(port=port, title="興聖集團 ERP 系統", host='0.0.0.0', reload=False)