"""Проверяет текущее состояние столбцов в Топ-500 ВБ.xlsx."""
import sys, io, openpyxl
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

wb = openpyxl.load_workbook(r'C:\Users\Admin\Desktop\Топ-500 ВБ\Топ-500 ВБ.xlsx')
ws = wb.worksheets[0]

# Считаем заполненность по столбцам
print('Колонка                              | Заполнено | Из 500')
print('-' * 60)
for col in range(1, 16):
    header = ws.cell(1, col).value or f'Col{col}'
    filled = sum(1 for r in range(2, 502) if ws.cell(r, col).value not in (None, '', 'None'))
    print(f'{str(header)[:35]:35s} | {filled:5d}     | 500')

print()
print('=== Примеры данных первых 10 строк ===')
for row_idx in range(2, 12):
    article  = ws.cell(row_idx, 1).value
    name     = ws.cell(row_idx, 2).value
    brand    = ws.cell(row_idx, 3).value
    buy      = ws.cell(row_idx, 4).value
    supplier = ws.cell(row_idx, 5).value
    params   = ws.cell(row_idx, 6).value
    oem      = ws.cell(row_idx, 7).value
    applic   = ws.cell(row_idx, 8).value
    alts     = ws.cell(row_idx, 9).value
    desc     = ws.cell(row_idx, 10).value
    print(f'#{row_idx-1:3d} {str(article)[:20]:20s} | {str(brand)[:15]:15s} | OEM={str(oem)[:20] if oem else "—":20s} | Applic={str(applic)[:30] if applic else "—"}')
