"""Очищает столбец Применяемость (col 8) в Excel."""
import sys, io, openpyxl
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

EXCEL_FILE = r'C:\Users\Admin\Desktop\Топ-500 ВБ\Топ-500 ВБ.xlsx'
wb = openpyxl.load_workbook(EXCEL_FILE)
ws = wb.worksheets[0]

cleared = 0
for r in range(2, 502):
    if ws.cell(r, 8).value:
        ws.cell(r, 8).value = None
        cleared += 1

wb.save(EXCEL_FILE)
print(f'Очищено строк Применяемость: {cleared}')
