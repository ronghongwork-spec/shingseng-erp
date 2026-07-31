from nicegui import ui
import pandas as pd
import plotly.express as px
import gspread
from google.oauth2.service_account import Credentials

# -------------------------------------------------------------------------
# Google Sheet 數據串接設定 (透過 Service Account API)
# -------------------------------------------------------------------------
def load_shingseng_target_data():
    """透過 Google Sheets API 直接安全讀取試算表數據"""
    try:
        scope = [
            'https://www.googleapis.com/auth/spreadsheets.readonly',
            'https://www.googleapis.com/auth/drive.readonly'
        ]
        # 如果你的 Render 環境變數或本地是用 credentials.json 檔案，或是可以直接讀取
        # 這裡會優先嘗試以 gspread 預設方式或指定 service_account 載入
        gc = gspread.service_account(filename='credentials.json')
        
        sh = gc.open_by_url("https://docs.google.com/spreadsheets/d/1Wc64Uqc1gvOS2CMX4QXRPOHDUsxuwDVS3tyRf5hexuY/edit?gid=80636189")
        worksheet = sh.get_worksheet(0) # 讀取第一個工作表
        data = worksheet.get_all_records()
        if data:
            df = pd.DataFrame(data)
            return df
    except Exception as e:
        print(f"API 讀取失敗，已自動切換為內建數據: {e}")
    
    # 預設的興聖分公司營業目標與銷售數據（備用）
    return pd.DataFrame({
        '月份': ['1月', '2月', '3月', '4月', '5月', '6月'],
        '通路': ['團購主A', '地區經銷商', '團購主B', '直客官網', '團購主A', '地區經銷商'],
        '目標銷售額': [300000, 350000, 400000, 380000, 450000, 500000],
        '實際銷售額': [280000, 360000, 420000, 350000, 480000, 520000]
    })

# -------------------------------------------------------------------------
# 模擬多公司與多倉庫資料庫
# -------------------------------------------------------------------------
PURCHASE_ORDERS = {
    '海濤客食品工廠': [],
    '興聖分公司': [],
    '容鴻分公司': [],
    '芙萊柏分公司': []
}

COMPANY_DATA = {
    '海濤客食品工廠': {
        'warehouses': ['總廠原料倉', '成品冷凍倉', '電商出貨倉'],
        'inventory_by_wh': {
            '總廠原料倉': pd.DataFrame({
                '品號': ['MAT-01', 'MAT-02', 'MAT-03'],
                '產品名稱': ['干貝原料', '辣椒油', '頂級蝦米'],
                '現有庫存': [500, 300, 150],
                '安全水位': [200, 100, 50],
                '狀態': ['庫存充足', '正常', '正常']
            }),
            '成品冷凍倉': pd.DataFrame({
                '品號': ['A001', 'A002', 'A003', 'A004'],
                '產品名稱': ['海濤客XO醬', '烏魚子禮盒', '干貝醬', '一口烏魚子'],
                '現有庫存': [150, 80, 220, 60],
                '安全水位': [50, 30, 100, 20],
                '狀態': ['正常', '正常', '庫存充足', '注意']
            }),
            '電商出貨倉': pd.DataFrame({
                '品號': ['A001', 'A002', 'A003', 'A004'],
                '產品名稱': ['海濤客XO醬', '烏魚子禮盒', '干貝醬', '一口烏魚子'],
                '現有庫存': [30, 15, 40, 10],
                '安全水位': [40, 20, 50, 15],
                '狀態': ['注意', '注意', '注意', '偏低']
            })
        },
        'orders': pd.DataFrame({
            '訂單編號': ['HTK-2026-01', 'HTK-2026-02'],
            '客戶/通路': ['momo購物網', 'Shopee經銷'],
            '品名': ['海濤客XO醬 x 10', '烏魚子禮盒 x 5'],
            '金額': [4500, 6000],
            '狀態': ['待確認', '已確認']
        }),
        'bom': pd.DataFrame({
            '母品號': ['BOM-XO-01', 'BOM-XO-01'],
            '子品號': ['MAT-01', 'MAT-02'],
            '物料名稱': ['干貝原料', '辣椒油'],
            '需求數量': [2, 1],
            '單位': ['公斤', '公升']
        }),
        'has_factory_modules': True
    },
    '興聖分公司': {
        'warehouses': ['興聖一倉', '興聖二倉 (暫存)'],
        'inventory_by_wh': {
            '興聖一倉': pd.DataFrame({
                '品號': ['HS-001', 'HS-002'],
                '產品名稱': ['興聖特選米', '高級苦茶油'],
                '現有庫存': [500, 120],
                '安全水位': [100, 30],
                '狀態': ['正常', '正常']
            }),
            '興聖二倉 (暫存)': pd.DataFrame({
                '品號': ['HS-001', 'HS-002'],
                '產品名稱': ['興聖特選米', '高級苦茶油'],
                '現有庫存': [50, 10],
                '安全水位': [20, 10],
                '狀態': ['正常', '注意']
            })
        },
        'orders': pd.DataFrame({
            '訂單編號': ['HS-ORD-01', 'HS-ORD-02'],
            '客戶/通路': ['團購主A', '地區經銷商'],
            '品名': ['興聖特選米 x 50', '高級苦茶油 x 10'],
            '金額': [15000, 8000],
            '狀態': ['已確認', '備貨中']
        }),
        'has_factory_modules': False
    },
    '容鴻分公司': {
        'warehouses': ['容鴻北區倉', '容鴻南區倉'],
        'inventory_by_wh': {
            '容鴻北區倉': pd.DataFrame({
                '品號': ['RH-001', 'RH-002'],
                '產品名稱': ['容鴻禮盒A', '容鴻禮盒B'],
                '現有庫存': [60, 25],
                '安全水位': [20, 10],
                '狀態': ['正常', '正常']
            }),
            '容鴻南區倉': pd.DataFrame({
                '品號': ['RH-001', 'RH-002'],
                '產品名稱': ['容鴻禮盒A', '容鴻禮盒B'],
                '現有庫存': [25, 15],
                '安全水位': [10, 10],
                '狀態': ['正常', '注意']
            })
        },
        'orders': pd.DataFrame({
            '訂單編號': ['RH-ORD-01', 'RH-ORD-02'],
            '客戶/通路': ['PChome', '直客'],
            '品名': ['容鴻禮盒A x 5', '容鴻禮盒B x 2'],
            '金額': [6200, 3100],
            '狀態': ['待確認', '已確認']
        }),
        'has_factory_modules': False
    },
    '芙萊柏分公司': {
        'warehouses': ['芙萊柏主倉'],
        'inventory_by_wh': {
            '芙萊柏主倉': pd.DataFrame({
                '品號': ['FB-001', 'FB-002'],
                '產品名稱': ['芙萊柏進口調味粉', '專用醬料'],
                '現有庫存': [310, 190],
                '安全水位': [80, 50],
                '狀態': ['正常', '正常']
            })
        },
        'orders': pd.DataFrame({
            '訂單編號': ['FB-ORD-01', 'FB-ORD-02'],
            '客戶/通路': ['餐飲通路', '零售商'],
            '品名': ['芙萊柏進口調味粉 x 20', '專用醬料 x 15'],
            '金額': [12000, 4500],
            '狀態': ['已確認', '已出貨']
        }),
        'has_factory_modules': False
    }
}

# -------------------------------------------------------------------------
# 主頁面佈局
# -------------------------------------------------------------------------
@ui.page('/')
def main_dashboard():
    ui.add_head_html('''
        <style>
            body { background-color: #f7f6f2; color: #1a1a1a; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }
            .q-table__container { background-color: #ffffff !important; color: #1a1a1a !important; border: 1px solid #e2e1dc; border-radius: 0px; box-shadow: none !important; }
            .q-table th { color: #555555 !important; font-weight: 700 !important; font-size: 13px; border-bottom: 2px solid #1a1a1a !important; }
            .q-table td { color: #1a1a1a !important; border-bottom: 1px solid #eeede8 !important; }
            .awwwards-btn { background-color: #5bc0be !important; color: #ffffff !important; font-weight: 700; }
        </style>
    ''')

    @ui.refreshable
    def render_content(selected_co: str):
        data = COMPANY_DATA[selected_co]
        is_factory = data['has_factory_modules']
        warehouses = data['warehouses']

        with ui.row().classes('w-full items-center justify-between bg-white border border-[#e2e1dc] p-4 mb-6 shadow-sm'):
            with ui.row().classes('items-center gap-3'):
                ui.icon('business').classes('text-xl text-zinc-800')
                ui.label(f'目前檢視單位：{selected_co}').classes('font-bold text-zinc-900 text-sm tracking-wide')
            badge_text = '● 啟用完整工廠模組' if is_factory else '● 標準商貿營運模式 (含營業目標計劃分析)'
            ui.label(badge_text).classes('text-xs font-bold px-3 py-1 bg-emerald-50 text-emerald-700 border border-emerald-200')

        with ui.tabs().classes('w-full bg-[#f7f6f2] px-0 text-zinc-500 border-b border-[#e2e1dc]') as tabs:
            t_dash = ui.tab('📊 總覽與數據分析', icon='dashboard')
            t_inv = ui.tab('📦 即時庫存清單', icon='inventory')
            t_ord = ui.tab('📋 訂單管理與確認', icon='shopping_cart')
            t_pur = ui.tab('🛒 採購單建立與管理', icon='add_shopping_cart')
            
            if selected_co == '興聖分公司':
                t_target = ui.tab('🎯 營業目標與銷售分析', icon='trending_up')

            if is_factory:
                t_bom = ui.tab('🌳 產品用料結構 (BOM)', icon='account_tree')
                t_sched = ui.tab('📅 生產與包裝排程', icon='calendar_month')

        with ui.tab_panels(tabs, value=t_dash).classes('w-full bg-[#f7f6f2] pt-6'):
            
            # 1. 總覽
            with ui.tab_panel(t_dash):
                default_wh_df = list(data['inventory_by_wh'].values())[0]
                df_ord = data['orders']
                
                with ui.row().classes('w-full gap-5 mb-8'):
                    with ui.card().classes('flex-1 p-6 bg-white border border-[#e2e1dc] shadow-none rounded-none'):
                        ui.label('總品項數量').classes('text-zinc-400 text-xs font-bold tracking-wider')
                        ui.label(str(len(default_wh_df))).classes('text-4xl font-black text-zinc-900 mt-2')
                    with ui.card().classes('flex-1 p-6 bg-white border border-[#e2e1dc] shadow-none rounded-none'):
                        ui.label('總庫存件數').classes('text-zinc-400 text-xs font-bold tracking-wider')
                        ui.label(str(default_wh_df['現有庫存'].sum())).classes('text-4xl font-black text-emerald-600 mt-2')
                    with ui.card().classes('flex-1 p-6 bg-white border border-[#e2e1dc] shadow-none rounded-none'):
                        ui.label('待處理訂單數').classes('text-zinc-400 text-xs font-bold tracking-wider')
                        ui.label(str(len(df_ord))).classes('text-4xl font-black text-amber-600 mt-2')

                fig = px.bar(default_wh_df, x='產品名稱', y='現有庫存', title=f'【{selected_co}】主倉庫存水位深度分析', color='現有庫存', color_continuous_scale='Tealgrn')
                fig.update_layout(
                    margin=dict(l=20, r=20, t=50, b=20), height=380, 
                    plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                    font=dict(color='#1a1a1a', family='sans-serif'),
                    title_font=dict(color='#1a1a1a', size=15)
                )

                with ui.row().classes('w-full gap-6'):
                    with ui.card().classes('flex-2 p-6 bg-white border border-[#e2e1dc] shadow-none rounded-none'):
                        ui.plotly(fig).classes('w-full')
                    with ui.card().classes('flex-1 p-6 bg-white border border-[#e2e1dc] shadow-none rounded-none'):
                        ui.label('數據洞察與摘要').classes('font-bold text-zinc-900 text-sm tracking-wide mb-3')
                        ui.markdown(f'• **{selected_co}** 目前多倉庫存連線正常。<br>• 興聖分公司已開通 Google Sheets API 營業目標與多維度銷售分析模組。').classes('text-zinc-600 text-sm leading-relaxed')

            # 2. 庫存
            with ui.tab_panel(t_inv):
                with ui.card().classes('w-full p-6 bg-white border border-[#e2e1dc] shadow-none rounded-none'):
                    with ui.row().classes('w-full items-center justify-between mb-4'):
                        ui.label(f'{selected_co} — 即時庫存明細表').classes('text-lg font-bold text-zinc-900 tracking-wide')
                        
                        with ui.row().classes('items-center gap-2'):
                            ui.label('選擇倉庫：').classes('text-zinc-500 text-xs font-bold')
                            wh_select = ui.select(options=warehouses, value=warehouses[0]).classes('bg-[#f7f6f2] text-zinc-900 rounded-none px-3 py-1 text-xs font-bold border border-[#e2e1dc]')

                    table_container = ui.column().classes('w-full')

                    def update_inventory_table(wh_name):
                        table_container.clear()
                        df_target = data['inventory_by_wh'][wh_name]
                        with table_container:
                            ui.table(
                                columns=[
                                    {'name': '品號', 'label': '品號', 'field': '品號', 'align': 'left'},
                                    {'name': '產品名稱', 'label': '產品名稱', 'field': '產品名稱', 'align': 'left'},
                                    {'name': '現有庫存', 'label': '現有庫存', 'field': '現有庫存'},
                                    {'name': '安全水位', 'label': '安全水位', 'field': '安全水位'},
                                    {'name': '狀態', 'label': '庫存狀態', 'field': '狀態'},
                                ],
                                rows=df_target.to_dict('records')
                            ).classes('w-full')

                    wh_select.on_value_change(lambda e: update_inventory_table(e.value))
                    update_inventory_table(warehouses[0])

            # 3. 訂單
            with ui.tab_panel(t_ord):
                with ui.card().classes('w-full p-6 bg-white border border-[#e2e1dc] shadow-none rounded-none'):
                    ui.label(f'{selected_co} — 訂單確認與審核').classes('text-lg font-bold text-zinc-900 mb-4 tracking-wide')
                    ui.table(
                        columns=[
                            {'name': '訂單編號', 'label': '訂單編號', 'field': '訂單編號'},
                            {'name': '通路', 'label': '客戶與通路', 'field': '客戶/通路'},
                            {'name': '品名', 'label': '訂購內容', 'field': '品名'},
                            {'name': '金額', 'label': '金額 (NTD)', 'field': '金額'},
                            {'name': '狀態', 'label': '處理狀態', 'field': '狀態'},
                        ],
                        rows=data['orders'].to_dict('records')
                    ).classes('w-full')

            # 4. 採購
            with ui.tab_panel(t_pur):
                with ui.card().classes('w-full p-6 bg-white border border-[#e2e1dc] shadow-none rounded-none mb-6'):
                    ui.label(f'{selected_co} — 建立新採購單').classes('text-lg font-bold text-zinc-900 mb-4 tracking-wide')
                    
                    with ui.row().classes('w-full gap-4'):
                        item_name_input = ui.input(label='品名').classes('flex-1')
                        item_no_input = ui.input(label='品號').classes('flex-1')
                        quantity_input = ui.number(label='採購數量', value=1).classes('w-32')
                    
                    with ui.row().classes('w-full gap-4 mt-2'):
                        supplier_input = ui.input(label='供應商').classes('flex-1')
                        remark_input = ui.input(label='備註').classes('flex-1')

                    def submit_order():
                        if not item_name_input.value or not item_no_input.value:
                            ui.notify('請填寫品名與品號！', color='negative')
                            return
                        
                        order_data = {
                            "品名": item_name_input.value,
                            "品號": item_no_input.value,
                            "數量": quantity_input.value,
                            "供應商": supplier_input.value,
                            "備註": remark_input.value
                        }
                        
                        PURCHASE_ORDERS[selected_co].append(order_data)
                        ui.notify(f'成功為 {selected_co} 建立採購單！', color='positive')
                        
                        item_name_input.value = ''
                        item_no_input.value = ''
                        quantity_input.value = 1
                        supplier_input.value = ''
                        remark_input.value = ''
                        
                        purchase_table.rows = PURCHASE_ORDERS[selected_co]
                        purchase_table.update()

                    ui.button('送出採購單', on_click=submit_order).classes('mt-4 awwwards-btn px-6 py-2 text-xs rounded-none')

                with ui.card().classes('w-full p-6 bg-white border border-[#e2e1dc] shadow-none rounded-none'):
                    ui.label(f'{selected_co} — 歷史採購單列表').classes('text-lg font-bold text-zinc-900 mb-4 tracking-wide')
                    
                    purchase_columns = [
                        {'name': '品名', 'label': '品名', 'field': '品名', 'required': True},
                        {'name': '品號', 'label': '品號', 'field': '品號'},
                        {'name': '數量', 'label': '數量', 'field': '數量'},
                        {'name': '供應商', 'label': '供應商', 'field': '供應商'},
                        {'name': '備註', 'label': '備註', 'field': '備註'},
                    ]
                    
                    purchase_table = ui.table(columns=purchase_columns, rows=PURCHASE_ORDERS[selected_co], row_key='品號').classes('w-full')

            # 5. 興聖分公司專屬：營業目標與銷售分析模組（串接 Google Sheets API）
            if selected_co == '興聖分公司':
                with ui.tab_panel(t_target):
                    target_df = load_shingseng_target_data()
                    
                    total_target = target_df['目標銷售額'].sum() if '目標銷售額' in target_df.columns else 0
                    total_actual = target_df['實際銷售額'].sum() if '實際銷售額' in target_df.columns else 0
                    achievement_rate = round((total_actual / total_target) * 100, 1) if total_target > 0 else 0

                    with ui.row().classes('w-full gap-5 mb-6'):
                        with ui.card().classes('flex-1 p-6 bg-white border border-[#e2e1dc] shadow-none rounded-none'):
                            ui.label('年度累計目標銷售額').classes('text-zinc-400 text-xs font-bold tracking-wider')
                            ui.label(f'NT$ {total_target:,}').classes('text-3xl font-black text-zinc-900 mt-2')
                        with ui.card().classes('flex-1 p-6 bg-white border border-[#e2e1dc] shadow-none rounded-none'):
                            ui.label('年度累計實際銷售額').classes('text-zinc-400 text-xs font-bold tracking-wider')
                            ui.label(f'NT$ {total_actual:,}').classes('text-3xl font-black text-emerald-600 mt-2')
                        with ui.card().classes('flex-1 p-6 bg-white border border-[#e2e1dc] shadow-none rounded-none'):
                            ui.label('全年度目標達成率').classes('text-zinc-400 text-xs font-bold tracking-wider')
                            ui.label(f'{achievement_rate} %').classes('text-3xl font-black text-blue-600 mt-2')

                    fig_monthly = px.bar(
                        target_df, x='月份', y=['目標銷售額', '實際銷售額'] if '目標銷售額' in target_df.columns else target_df.columns[1], 
                        barmode='group', 
                        title='【興聖】各月份銷售額與目標多維度對比',
                        color_discrete_sequence=['#d1d5db', '#0d9488']
                    )
                    fig_monthly.update_layout(
                        margin=dict(l=20, r=20, t=50, b=20), height=380,
                        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                        font=dict(color='#1a1a1a', family='sans-serif'),
                        title_font=dict(color='#1a1a1a', size=15)
                    )

                    channel_df = target_df.groupby('通路')[['實際銷售額', '目標銷售額']].sum().reset_index() if '通路' in target_df.columns else target_df
                    fig_channel = px.pie(
                        channel_df, names='通路' if '通路' in channel_df.columns else channel_df.columns[0], 
                        values='實際銷售額' if '實際銷售額' in channel_df.columns else channel_df.columns[1], 
                        title='【興聖】各銷售通路實際營業額佔比分佈',
                        hole=0.4,
                        color_discrete_sequence=px.colors.sequential.Tealgrn
                    )
                    fig_channel.update_layout(
                        margin=dict(l=20, r=20, t=50, b=20), height=380,
                        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                        font=dict(color='#1a1a1a', family='sans-serif'),
                        title_font=dict(color='#1a1a1a', size=15)
                    )

                    with ui.row().classes('w-full gap-6 mb-6'):
                        with ui.card().classes('flex-2 p-6 bg-white border border-[#e2e1dc] shadow-none rounded-none'):
                            ui.plotly(fig_monthly).classes('w-full')
                        with ui.card().classes('flex-1 p-6 bg-white border border-[#e2e1dc] shadow-none rounded-none'):
                            ui.plotly(fig_channel).classes('w-full')

                    with ui.card().classes('w-full p-6 bg-white border border-[#e2e1dc] shadow-none rounded-none'):
                        with ui.row().classes('w-full items-center justify-between mb-4'):
                            ui.label('Google Sheets API 遠端來源明細數據').classes('text-lg font-bold text-zinc-900 tracking-wide')
                            ui.button('重新載入遠端數據', on_click=lambda: render_content.refresh(selected_co)).classes('awwwards-btn px-4 py-1 text-xs rounded-none')
                        
                        ui.table(
                            columns=[{'name': col, 'label': col, 'field': col, 'align': 'left'} for col in target_df.columns],
                            rows=target_df.to_dict('records')
                        ).classes('w-full')

            # 6. 工廠模組
            if is_factory:
                with ui.tab_panel(t_bom):
                    with ui.card().classes('w-full p-6 bg-white border border-[#e2e1dc] shadow-none rounded-none'):
                        ui.label('海濤客食品工廠 — 產品用料清單 (BOM)').classes('text-lg font-bold text-zinc-900 mb-4 tracking-wide')
                        ui.table(
                            columns=[
                                {'name': '母品號', 'label': '主產品品號', 'field': '母品號'},
                                {'name': '子品號', 'label': '物料品號', 'field': '子品號'},
                                {'name': '物料名稱', 'label': '物料名稱', 'field': '物料名稱'},
                                {'name': '需求數量', 'label': '單位用量', 'field': '需求數量'},
                                {'name': '單位', 'label': '單位', 'field': '單位'},
                            ],
                            rows=data['bom'].to_dict('records')
                        ).classes('w-full')

                with ui.tab_panel(t_sched):
                    with ui.row().classes('w-full gap-6'):
                        with ui.card().classes('flex-2 p-6 bg-white border border-[#e2e1dc] shadow-none rounded-none'):
                            ui.label('本週包裝與生產排程').classes('text-lg font-bold text-zinc-900 mb-4 tracking-wide')
                            ui.markdown('''
                            * **週一 (08:00 - 12:00)**：海濤客 XO 醬禮盒批次包裝作業
                            * **週二 (13:00 - 17:00)**：干貝醬真空封口與品管作業
                            * **週四 (全天)**：一口烏魚子真空包裝與精美禮盒裝箱
                            ''').classes('text-zinc-600 text-sm leading-loose')
                        with ui.card().classes('flex-1 p-6 bg-white border border-[#e2e1dc] shadow-none rounded-none'):
                            ui.label('廠區行事曆備忘').classes('text-lg font-bold text-zinc-900 mb-4 tracking-wide')
                            ui.date().classes('w-full border border-[#e2e1dc] bg-white shadow-none text-zinc-900')

    with ui.row().classes('w-full items-center justify-between bg-[#f7f6f2] text-zinc-900 px-8 py-4 border-b border-[#e2e1dc] sticky top-0 z-50'):
        with ui.row().classes('items-center gap-3'):
            ui.icon('domain', size='sm').classes('text-zinc-800')
            ui.label('興聖集團｜智慧 ERP 總管理中樞').classes('text-sm font-black tracking-wider')
        
        with ui.row().classes('items-center gap-4'):
            ui.label('檢視單位：').classes('text-zinc-500 text-xs font-bold')
            company_select = ui.select(options=list(COMPANY_DATA.keys()), value='興聖分公司').classes('bg-white text-zinc-900 rounded-none px-3 py-1 text-xs font-bold border border-[#e2e1dc]')
            ui.button('立即同步', on_click=lambda: ui.notify('資料同步觸發成功')).classes('awwwards-btn px-4 py-2 text-xs rounded-none')

    with ui.column().classes('w-full p-8 max-w-[1600px] mx-auto bg-[#f7f6f2]'):
        company_select.on_value_change(lambda e: render_content.refresh(e.value))
        render_content(company_select.value)

ui.run(port=8080, title="興聖集團 ERP 系統", host='0.0.0.0')