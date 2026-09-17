import requests
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

def fetch_tainex_events():
    # 使用正確的網頁資料請求 API
    api_url = "https://www.tainex.com.tw/api/event/getEventList"
    
    # 模擬真實瀏覽器的完整 Header
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
        'Accept': 'application/json, text/plain, */*',
        'Referer': 'https://www.tainex.com.tw/event?hall=1',
        'Content-Type': 'application/json;charset=UTF-8',
        'Origin': 'https://www.tainex.com.tw'
    }
    
    # POST Payload 參數（查詢 1 館展覽）
    payload = {
        "hall": "1",
        "language": "zh-TW"
    }

    events = []
    try:
        # 先以 POST 嘗試，若非 POST 則退回 GET 請求
        response = requests.post(api_url, headers=headers, json=payload, timeout=15)
        
        if response.status_code != 200:
            # 備用端點 GET
            api_url_get = "https://www.tainex.com.tw/api/v1/event/list?hall=1"
            response = requests.get(api_url_get, headers=headers, timeout=15)

        if response.status_code == 200:
            res_data = response.json()
            # 兼容不同回傳層級結構
            items = []
            if isinstance(res_data, list):
                items = res_data
            elif isinstance(res_data, dict):
                items = res_data.get('data', res_data.get('result', res_data.get('events', [])))
            
            for item in items:
                if isinstance(item, dict):
                    title = item.get('title', item.get('eventName', item.get('name', '')))
                    start = item.get('startDate', item.get('startDateStr', item.get('start', '')))
                    end = item.get('endDate', item.get('endDateStr', item.get('end', '')))
                    hall = item.get('hallName', item.get('hall', '1館'))

                    if title:
                        events.append({
                            "展覽名稱": str(title).strip(),
                            "時間起": str(start).strip(),
                            "時間迄": str(end).strip(),
                            "展館": str(hall).strip()
                        })
        else:
            print(f"HTTP 請求回應代碼: {response.status_code}")

    except Exception as e:
        print(f"抓取過程發生例外狀況: {e}")

    return events

def export_to_excel(events, filename="tainex_events.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TaiNEX展覽檔期"

    # 大標題
    ws.merge_cells("A1:D1")
    ws["A1"] = "台北南港展覽館 展覽活動檔期表"
    ws["A1"].font = Font(name="微軟正黑體", size=15, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 35

    # 欄位表頭
    headers = ["展覽名稱", "時間起", "時間迄", "展館"]
    ws.row_dimensions[3].height = 25
    header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    header_font = Font(name="微軟正黑體", size=11, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    data_font = Font(name="微軟正黑體", size=10)
    zebra_fill = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")

    for row_idx, item in enumerate(events, 4):
        ws.row_dimensions[row_idx].height = 22
        c1 = ws.cell(row=row_idx, column=1, value=item["展覽名稱"])
        c2 = ws.cell(row=row_idx, column=2, value=item["時間起"])
        c3 = ws.cell(row=row_idx, column=3, value=item["時間迄"])
        c4 = ws.cell(row=row_idx, column=4, value=item["展館"])

        c1.alignment = Alignment(horizontal="left", vertical="center")
        c2.alignment = Alignment(horizontal="center", vertical="center")
        c3.alignment = Alignment(horizontal="center", vertical="center")
        c4.alignment = Alignment(horizontal="center", vertical="center")

        for c in [c1, c2, c3, c4]:
            c.font = data_font
            c.border = thin_border
            if row_idx % 2 == 0:
                c.fill = zebra_fill

    ws.column_dimensions['A'].width = 45
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 12

    wb.save(filename)
    print(f"成功儲存檔案至: {filename}")

if __name__ == "__main__":
    data = fetch_tainex_events()
    
    # 保底機制：若真的因為網頁擋海外 IP，產出寫有原因的表格，確保 Excel 檔案 100% 存在
    if not data:
        print("未抓取到資料，產生備用 Excel 檔案...")
        data = [{
            "展覽名稱": "暫無資料或遭受防護牆阻擋（請檢查 API）", 
            "時間起": "-", 
            "時間迄": "-", 
            "展館": "1館"
        }]
        
    export_to_excel(data, "tainex_events.xlsx")
