import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

def fetch_tainex_events():
    # 直接存取 TaiNEX 後台 API
    api_url = "https://www.tainex.com.tw/api/v1/events?hall=1"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'application/json, text/plain, */*'
    }
    
    events = []
    try:
        response = requests.get(api_url, headers=headers, timeout=15)
        if response.status_code == 200:
            data = response.json()
            # 依據 API 回傳結構解析（假設包含 events 或 data 陣列）
            items = data.get('events', data.get('data', [])) if isinstance(data, dict) else data
            
            for item in items:
                title = item.get('title', item.get('name', ''))
                start = item.get('startDate', item.get('start_date', ''))
                end = item.get('endDate', item.get('end_date', ''))
                hall = item.get('hallName', item.get('hall', '1館'))
                
                if title:
                    events.append({
                        "展覽名稱": title,
                        "時間起": start,
                        "時間迄": end,
                        "展館": hall
                    })
        else:
            print(f"API 回應異常，Status code: {response.status_code}")
    except Exception as e:
        print(f"請求失敗: {e}")

    return events

def export_to_excel(events, filename="tainex_events.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TaiNEX展覽檔期"

    ws.merge_cells("A1:D1")
    ws["A1"] = "台北南港展覽館 展覽活動檔期表"
    ws["A1"].font = Font(name="微軟正黑體", size=15, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 35

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
    print(f"成功產生並儲存檔案至: {filename}")

if __name__ == "__main__":
    data = fetch_tainex_events()
    
    # 【關鍵防呆】若無抓到資料，產出預設提示行，確保產出 tainex_events.xlsx
    if not data:
        print("未抓取到線上資料，產生備用 Excel 檔...")
        data = [{
            "展覽名稱": "尚未取得最新展覽資料或 API 回應結構異動", 
            "時間起": "-", 
            "時間迄": "-", 
            "展館": "1館"
        }]
        
    export_to_excel(data, "tainex_events.xlsx")
