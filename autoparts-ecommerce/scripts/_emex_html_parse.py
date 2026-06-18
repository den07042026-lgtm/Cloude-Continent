"""Ищет OEM/аналоги/применяемость прямо в HTML emex.ru."""
import re, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

with open(r'data/analytics/emex_debug/products_CAF100493C_CHAMPION.html', encoding='utf-8') as f:
    html = f.read()

print(f'HTML размер: {len(html)} символов')
print()

# 1. Ищем упоминания брендов и артикулов в контексте "аналоги"
patterns = [
    ('OEM/Original',    r'(?i)(OEM|original|оригинал)[^\n]{0,200}'),
    ('FORD в тексте',   r'FORD[^\n<]{0,100}'),
    ('Ford Focus',      r'[Ff]ord\s+[Ff]ocus[^\n<]{0,150}'),
    ('MANN',            r'MANN[^\n<]{0,100}'),
    ('аналог',          r'(?i)аналог[^\n<]{0,200}'),
    ('применяем',       r'(?i)применяем[^\n<]{0,200}'),
    ('CHAMPION',        r'CHAMPION[^\n<]{0,100}'),
]
for name, pat in patterns:
    matches = re.findall(pat, html)
    print(f'[{name}] найдено {len(matches)} совпадений')
    for m in matches[:3]:
        print(f'  {m[:200]}')
    print()

# 2. Ищем JSON-блоки с артикулами
print('=== JSON-блоки с артикулами ===')
json_items = re.findall(r'\{[^{}]{0,500}"(?:number|article|partNumber|detailNum|num)"[^{}]{0,500}\}', html)
print(f'Найдено JSON-фрагментов: {len(json_items)}')
for item in json_items[:5]:
    print(f'  {item[:300]}')
print()

# 3. window.__INITIAL_STATE__ или подобное
for var in ['__INITIAL_STATE__', '__PRELOADED_STATE__', '__APP_STATE__', 'window.pageData']:
    if var in html:
        idx = html.index(var)
        print(f'FOUND {var} at position {idx}')
        print(html[idx:idx+500])
        print()

# 4. Ищем блоки data-* с артикулами
data_attrs = re.findall(r'data-[a-z]+="[A-Z0-9]{5,20}"', html)
print(f'data-attrs с артикулами: {len(data_attrs)}')
for a in data_attrs[:10]:
    print(f'  {a}')
