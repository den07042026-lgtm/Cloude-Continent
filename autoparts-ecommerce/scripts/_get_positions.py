import sys, io, openpyxl
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

wb = openpyxl.load_workbook('data/analytics/wb_top500_combined.xlsx')
ws = wb.worksheets[0]

TARGETS = {34, 255, 256, 380, 381, 383, 386}

print('Позиция | Артикул              | Бренд              | Название')
print('-' * 100)
for row in ws.iter_rows(min_row=2, max_row=501, values_only=True):
    rank = row[0]
    if rank in TARGETS:
        article = str(row[1])
        brand   = str(row[2])
        name    = str(row[3])
        src     = str(row[4])
        buy     = row[5]
        print(f'  {rank:>3}   | {article:<20} | {brand:<18} | {name[:55]}  ({src}, {buy} руб)')
