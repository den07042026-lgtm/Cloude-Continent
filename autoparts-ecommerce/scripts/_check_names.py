import sys, io, openpyxl
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

wb = openpyxl.load_workbook(r'C:\Users\Admin\Desktop\Топ-500 ВБ\Топ-500 ВБ_new.xlsx')
ws = wb.worksheets[0]

TARGET_ROWS = set(range(19, 23)) | set(range(24, 32))

print(f'{"Строка":<8} {"Артикул":<22} {"Бренд":<18} {"Наименование"}')
print('-' * 110)
for row in ws.iter_rows(min_row=2, max_row=502, values_only=True):
    # строка в файле = row[0] это данные, нам нужен номер строки Excel
    pass

# Перечитаем с номерами строк
for row_idx in range(2, 502):
    excel_row = row_idx  # строка в Excel
    data_row  = row_idx - 1  # номер строки данных (1=первый товар)
    if data_row not in TARGET_ROWS:
        continue
    article = ws.cell(row_idx, 1).value
    name    = ws.cell(row_idx, 2).value
    brand   = ws.cell(row_idx, 3).value
    buy     = ws.cell(row_idx, 4).value
    supplier= ws.cell(row_idx, 5).value
    print(f'  #{data_row:<5} | {str(article):<22} | {str(brand):<18} | {str(name)[:70]}')
