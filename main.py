import csv
import os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, colors, Border, Side

# 讀取CSV檔案
csv_path = 'assets/112年中華民國政府行政機關辦公日曆表.csv'
out_path = 'out/'
data = []
if not os.path.exists(out_path):
    os.makedirs(out_path)
out_path += '112_傳輸機房溫度檢查表.xlsx'

# 欄、列寬
column_widths = {
    'A': 15,
    'B': 10,
    'C': 25,
    'D': 25,
    'E': 10,
    'F': 25
}
row_heights = {
    1: 35,
    2: 25
}

# 開始處理
with open(csv_path, 'r', encoding='big5') as file:
    reader = csv.DictReader(file)
    for row in reader:
        data.append(row)

# 建立Excel工作簿
workbook = Workbook()

# 解析CSV資料並生成Excel檔案
for month in range(1, 13):
    # 過濾當月的資料
    month_data = [row for row in data if int(row['西元日期'][4:6]) == month]
    if not month_data:
        continue

    # 創建工作表
    sheet_name = f'{month}月'
    sheet = workbook.create_sheet(title=sheet_name)
    sheet.merge_cells('A1:F1')
    title = f'第一網維二股 水湳傳輸機房 {month_data[0]["西元日期"][:4]}年{month}月'
    sheet['A1'] = title
    title_font = Font(size=30, bold=True)
    sheet['A1'].font = title_font
    sheet['A1'].alignment = Alignment(horizontal='center')

    # 調整列寬
    for column, width in column_widths.items():
        sheet.column_dimensions[column].width = width

    # 寫入標題
    header = ['日期', '星期', '量測點E12溫度', '量測點J09溫度', '量測者', '備註']
    header_font = Font(size=18, bold=True)
    sheet.append(header)
    for cell in sheet[2]:
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center')

    # 設定表格邊框
    border_style = Side(border_style="thin", color="000000")
    border = Border(top=border_style, bottom=border_style, left=border_style, right=border_style)

    # 設定標題和標頭行的邊框
    title_cell = sheet['A1']
    title_cell.border = border

    for cell in sheet[2]:
        cell.border = border

    # 調整行高
    for row, height in row_heights.items():
        sheet.row_dimensions[row].height = height

    # 寫入資料
    for row in month_data:
        date = row['西元日期']
        day = int(date[6:8])
        weekday = row['星期']
        holiday = int(row['是否放假'])
        remark = row['備註']

        # 判斷是否放假，並將整行文字改為紅色
        if holiday == 2:
            font = Font(color='FF0000', size=15, bold=True)
            remark = remark.upper()
        else:
            font = Font(size=15)

        # 修改日期格式
        date = f'{month}月{day}日'

        # 判斷備註是否為「補行上班」，並將整行文字改為藍色
        if remark == '補行上班':
            font.color = colors.BLUE

        # 寫入資料到Excel
        sheet.append([date, weekday, None, None, None, remark])

        # 設定每個儲存格的邊框
        for cell in sheet[sheet.max_row]:
            cell.font = font
            cell.border = border
            cell.alignment = Alignment(horizontal='center')


# 刪除預設的工作表
workbook.remove(workbook['Sheet'])

# 儲存Excel檔案
workbook.save(out_path)
