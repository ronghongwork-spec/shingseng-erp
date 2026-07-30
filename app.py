from nicegui import ui

# 模擬儲存採購單的資料庫或列表
purchase_orders = {
    "海濤客": [],
    "容鴻": [],
    "芙萊柏": []
}

def create_purchase_page(company_name: str):
    """建立各公司的採購頁面與採購單介面"""
    with ui.column().classes('w-full p-4'):
        ui.label(f'{company_name} - 採購管理系統').classes('text-2xl font-bold mb-4')
        
        # 採購單輸入表單
        with ui.card().classes('w-full p-4 mb-4'):
            ui.label('建立新採購單').classes('text-lg font-semibold mb-2')
            
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
                
                purchase_orders[company_name].append(order_data)
                ui.notify(f'成功為 {company_name} 建立採購單！', color='positive')
                
                # 清空欄位
                item_name_input.value = ''
                item_no_input.value = ''
                quantity_input.value = 1
                supplier_input.value = ''
                remark_input.value = ''
                
                # 重新整理表格
                order_table.rows = purchase_orders[company_name]
                order_table.update()

            ui.button('送出採購單', on_click=submit_order).classes('mt-4 bg-blue-500 text-white')

        # 採購單清單表格
        ui.label('歷史採購單列表').classes('text-lg font-semibold mb-2')
        
        columns = [
            {'name': '品名', 'label': '品名', 'field': '品名', 'required': True},
            {'name': '品號', 'label': '品號', 'field': '品號'},
            {'name': '數量', 'label': '數量', 'field': '數量'},
            {'name': '供應商', 'label': '供應商', 'field': '供應商'},
            {'name': '備註', 'label': '備註', 'field': '備註'},
        ]
        
        order_table = ui.table(columns=columns, rows=purchase_orders[company_name], row_key='品號').classes('w-full')

# 註冊各公司的頁面路由
@ui.page('/haitaoke_purchase')
def haitaoke_page():
    create_purchase_page("海濤客")

@ui.page('/ronghong_purchase')
def ronghong_page():
    create_purchase_page("容鴻")

@ui.page('/freiber_purchase')
def freiber_page():
    create_purchase_page("芙萊柏")

# 主頁導覽連結
@ui.page('/')
def main_index():
    ui.label('多公司 ERP 採購系統').classes('text-3xl font-bold mb-6')
    ui.link('前往 海濤客 採購頁面', '/haitaoke_purchase').classes('text-blue-600 block mb-2')
    ui.link('前往 容鴻 採購頁面', '/ronghong_purchase').classes('text-blue-600 block mb-2')
    ui.link('前往 芙萊柏 採購頁面', '/freiber_purchase').classes('text-blue-600 block mb-2')

ui.run(port=8080, host='0.0.0.0')