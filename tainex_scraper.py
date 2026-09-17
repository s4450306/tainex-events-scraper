import urllib.request
import re
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

def fetch_tainex_events():
    url = "https://www.tainex.com.tw/event?hall=1"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36'
    }
    
    req = urllib.request.Request(url, headers=headers)
    try:
        with urllib.request.urlopen(req) as response:
            html = response.read().decode('utf-8')
    except Exception as e:
        print(f"網絡請求失敗: {e}")
        return []

    soup = BeautifulSoup(html, 'html.parser')
    events = []

    # 解析展覽卡片
    event_cards = soup.find_all(['div', 'li'], class_=re.compile(r'event|card|item', re.I))
    
    for card in event_cards:
        try:
            title_el = card.find(['h3', 'h4', 'a', 'div'], class_=re.compile(r'title|name', re.I))
            title = title_el.get_text(strip=True) if title_el else ""

            date_el = card.find(class_=re.compile(r'date|time', re.I))
            date_str = date_el.get_text(strip=True) if date_el else ""
            
            start_date, end_date = "", ""
            if "-" in date_str or "~" in date_str or "〜" in date_str:
                dates = re.split(r'[-~〜]', date_str)
                start_date = dates[0].strip()
                end_date = dates[1].strip() if len(dates) > 1 else start_date
            else:
                start_date = date_str

            hall_el = card.find(text=re.compile(r'[12一二]館'))
            hall = hall_el.strip() if hall_el else "1館"

            if title and len(title) > 2:
                events.append({
                    "展覽名稱": title,
                    "時間起": start_date,
                    "時間迄": end_date,
                    "展館": hall
                })
        except Exception:
            continue

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
    print(f"成功儲存檔案至: {filename}")

if __name__ == "__main__":
    data = fetch_tainex_events()
    # 就算抓不到資料也產生範例 Excel，確保檔案一定會建立
    if not data:
        data = [{"展覽名稱": "暫無展覽資料或網頁結構變更", "時間起": "-", "時間迄": "-", "展館": "1館"}]
    export_to_excel(data)
