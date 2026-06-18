"""Очищает col 7 (OEM) и col 9 (Альтернативные) от старых данных."""
import sys, io, openpyxl
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

EXCEL_FILE = r'C:\Users\Admin\Desktop\Топ-500 ВБ\Топ-500 ВБ.xlsx'
wb = openpyxl.load_workbook(EXCEL_FILE)
ws = wb.worksheets[0]

c7 = c9 = 0
for r in range(2, 502):
    if ws.cell(r, 7).value:
        ws.cell(r, 7).value = None
        c7 += 1
    if ws.cell(r, 9).value:
        ws.cell(r, 9).value = None
        c9 += 1

wb.save(EXCEL_FILE)
print(f'Очищено OEM (col7): {c7}  |  Альтернативные (col9): {c9}')
