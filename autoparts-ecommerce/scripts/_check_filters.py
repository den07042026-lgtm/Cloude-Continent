"""Быстрая проверка фильтров в пуле Автолига + Микадо."""
import sys, io, xlrd, openpyxl, re
from collections import Counter
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

def norm(s):
    return re.sub(r'[\s\-\.\/\\]', '', str(s)).upper()

def is_filter(name):
    n = name.lower()
    return 'фильтр' in n or 'filter' in n

# --- Автолига ---
book = xlrd.open_workbook('data/suppliers/autoliga/PriceALVLG0411.xls', encoding_override='cp1251')
ws = book.sheet_by_index(0)
al_all, al_filters = [], []
for r in range(9, ws.nrows):
    try:
        brand   = str(ws.cell_value(r, 1)).strip()
        article = str(ws.cell_value(r, 2)).strip()
        name    = str(ws.cell_value(r, 4)).strip()
        stock   = float(str(ws.cell_value(r, 6)).replace(',', '.') or 0)
        price   = float(str(ws.cell_value(r, 7)).replace(',', '.') or 0)
    except:
        continue
    if not brand or not article or stock <= 0 or price <= 0:
        continue
    key = (norm(brand), norm(article))
    al_all.append(key)
    if is_filter(name):
        al_filters.append({'brand': brand, 'article': article, 'name': name,
                           'buy': price, 'stock': int(stock), 'key': key})

# --- Микадо ---
wb = openpyxl.load_workbook('data/suppliers/mikado/mikado_price_live.xlsx', read_only=True, data_only=True)
wm = wb.active
mk_all, mk_filters = [], []
for row in wm.iter_rows(min_row=2, values_only=True):
    try:
        article = str(row[1]).strip() if row[1] else ''
        brand   = str(row[2]).strip() if row[2] else ''
        name    = str(row[3]).strip() if row[3] else ''
        price   = float(str(row[4]).replace(',', '.')) if row[4] else 0.0
        qty_raw = str(row[5]).strip() if row[5] else '0'
        stock   = 10 if qty_raw.startswith('>') else int(float(qty_raw))
    except:
        continue
    if not brand or not article or stock <= 0 or price <= 0:
        continue
    key = (norm(brand), norm(article))
    mk_all.append(key)
    if is_filter(name):
        mk_filters.append({'brand': brand, 'article': article, 'name': name,
                           'buy': price, 'stock': stock, 'key': key})
wb.close()

al_keys = set(it['key'] for it in al_filters)
mk_keys = set(it['key'] for it in mk_filters)
both_keys = al_keys & mk_keys

# Категории фильтров
def filter_cat(name):
    n = name.lower()
    if 'масл' in n or 'oil' in n:    return 'масляный'
    if 'воздуш' in n or 'air' in n:  return 'воздушный'
    if 'салон' in n or 'cabin' in n: return 'салонный'
    if 'топлив' in n or 'fuel' in n: return 'топливный'
    return 'прочий'

al_cats  = Counter(filter_cat(it['name']) for it in al_filters)
mk_cats  = Counter(filter_cat(it['name']) for it in mk_filters)

print('=== ФИЛЬТРЫ В ПРАЙСАХ ===')
print(f'Автолига: {len(al_filters)} фильтров ({len(al_keys)} уник. ключей)')
print(f'  По типам: {dict(al_cats)}')
print(f'Микадо:   {len(mk_filters)} фильтров ({len(mk_keys)} уник. ключей)')
print(f'  По типам: {dict(mk_cats)}')
print()
print(f'Совпадают (brand+article) у обоих: {len(both_keys)}')
print(f'Только Автолига: {len(al_keys - mk_keys)}')
print(f'Только Микадо:   {len(mk_keys - al_keys)}')
print(f'Всего уник. фильтров в объед. пуле: {len(al_keys | mk_keys)}')
