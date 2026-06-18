"""Анализирует структуру __NEXT_DATA__ из сохранённого HTML emex.ru"""
import json, re, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

with open(r'data/analytics/emex_debug/products_CAF100493C_CHAMPION.html', encoding='utf-8') as f:
    html = f.read()

m = re.search(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', html, re.DOTALL)
if not m:
    print('__NEXT_DATA__ NOT FOUND')
    exit(1)

data = json.loads(m.group(1))
print('TOP KEYS:', list(data.keys()))
pp = data.get('props', {}).get('pageProps', {})
print('pageProps KEYS:', list(pp.keys())[:30])
print()

# Ищем OEM / cross-references / applicability
def show_key(d, key, max_items=5):
    val = d.get(key)
    if val is None:
        return
    print(f'  [{key}]:')
    if isinstance(val, list):
        print(f'    list len={len(val)}, first items:')
        for item in val[:max_items]:
            print(f'      {item}')
    elif isinstance(val, dict):
        print(f'    dict keys: {list(val.keys())[:15]}')
    else:
        print(f'    {str(val)[:200]}')

interesting = ['detail', 'product', 'oem', 'analog', 'cross', 'applicability',
               'vehicles', 'cars', 'makes', 'models', 'references', 'oems',
               'crossReferences', 'oemNumbers', 'analogs', 'details']

for k in pp:
    print(f'KEY: {k}  type={type(pp[k]).__name__}', end='')
    if isinstance(pp[k], (list, dict)):
        if isinstance(pp[k], list):
            print(f'  len={len(pp[k])}', end='')
        else:
            print(f'  subkeys={list(pp[k].keys())[:8]}', end='')
    print()

print()
print('=== detail/product info ===')
for k in ['detail', 'product', 'detailInfo', 'productInfo']:
    if k in pp:
        show_key(pp, k)
        d = pp[k]
        if isinstance(d, dict):
            for sk in d:
                print(f'    sub[{sk}] = {str(d[sk])[:150]}')
