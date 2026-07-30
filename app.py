import gspread
import pandas as pd

def load_shingseng_target_data():
    """從結構化 Google 試算表中精準萃取第一季各通路數據"""
    try:
        gc = gspread.service_account(filename='credentials.json')
        sh = gc.open_by_url("https://docs.google.com/spreadsheets/d/1Wc64Uqc1gvOS2CMX4QXRPOHDUsxuwDVS3tyRf5hexuY/edit?gid=80636189")
        worksheet = sh.get_worksheet(0) 
        
        # 抓取整張表的原始二維資料陣列 (Row x Col)
        rows = worksheet.get_all_values()
        
        if len(rows) > 10:
            parsed_data = []
            
            # 根據你的截圖位置，明細資料大約從第 10 行開始 (索引 9 開始)
            # 每一列包含：銷貨收入(通路) | 1月(2025, 實際, 佔比, 2026目標) | 2月(...) | 3月(...)
            for r in rows[10:]:
                channel_name = r[0].strip() # 第一欄是通路名稱 (例如 X線上 - FB官網)
                if not channel_name or '總計' in channel_name or '合計' in channel_name:
                    continue
                
                # 萃取各月份的「實際營收」與「目標營收」(根據圖片欄位對應索引)
                # 1月實際: 索引 2, 1月2026目標: 索引 4
                # 2月實際: 索引 7, 2月2026目標: 索引 9
                # 3月實際: 索引 12, 3月2026目標: 索引 14
                try:
                    m1_actual = float(r[2].replace(',', '')) if r[2] else 0
                    m1_target = float(r[4].replace(',', '')) if r[4] else 0
                    
                    m2_actual = float(r[7].replace(',', '')) if r[7] else 0
                    m2_target = float(r[9].replace(',', '')) if r[9] else 0
                    
                    m3_actual = float(r[12].replace(',', '')) if r[12] else 0
                    m3_target = float(r[14].replace(',', '')) if r[14] else 0
                    
                    # 加總成為第一季 Q1 數據，或按月份展開
                    parsed_data.append({
                        '通路': channel_name,
                        '月份': '1月',
                        '目標銷售額': m1_target,
                        '實際銷售額': m1_actual
                    })
                    parsed_data.append({
                        '通路': channel_name,
                        '月份': '2月',
                        '目標銷售額': m2_target,
                        '實際銷售額': m2_actual
                    })
                    parsed_data.append({
                        '通路': channel_name,
                        '月份': '3月',
                        '目標銷售額': m3_target,
                        '實際銷售額': m3_actual
                    })
                except Exception as inner_e:
                    continue
            
            if parsed_data:
                return pd.DataFrame(parsed_data)

    except Exception as e:
        print(f"API 解析失敗，已切換為備用數據: {e}")
    
    # 備用模擬數據
    return pd.DataFrame({
        '月份': ['1月', '1月', '2月', '2月', '3月', '3月'],
        '通路': ['X線上-FB', 'R線上-蝦皮', 'X線上-FB', 'R線上-蝦皮', 'X線上-FB', 'R線上-蝦皮'],
        '目標銷售額': [11168257, 2138820, 2236046, 664841, 2395151, 775962],
        '實際銷售額': [8278256, 1374061, 3072897, 796704, 887073, 457013]
    })