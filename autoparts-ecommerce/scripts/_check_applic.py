"""Проверяет качество данных в столбце Применяемость."""
import sys, io, openpyxl
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

wb = openpyxl.load_workbook(r'C:\Users\Admin\Desktop\Топ-500 ВБ\Топ-500 ВБ.xlsx')
ws = wb.worksheets[0]

# Примеры пустых строк
print('=== 20 строк без Применяемость ===')
empty_cnt = 0
for r in range(2, 502):
    applic = ws.cell(r, 8).value
    if not applic:
        article = ws.cell(r, 1).value
        name    = ws.cell(r, 2).value
        brand   = ws.cell(r, 3).value
        print(f'  #{r-1:3d}  {str(article)[:20]:20s}  {str(brand)[:15]:15s}  {str(name)[:60]}')
        empty_cnt += 1
        if empty_cnt >= 20:
            break

print(f'\nВсего без Применяемость: {sum(1 for r in range(2,502) if not ws.cell(r,8).value)}')

# Примеры с Применяемостью — насколько детальные?
print('\n=== 10 примеров с Применяемостью ===')
filled = [(r, ws.cell(r,8).value) for r in range(2,502) if ws.cell(r,8).value]
for r, applic in filled[:10]:
    name = ws.cell(r, 2).value
    print(f'  #{r-1:3d}  {str(applic)[:80]}')
    print(f'       name: {str(name)[:80]}')
