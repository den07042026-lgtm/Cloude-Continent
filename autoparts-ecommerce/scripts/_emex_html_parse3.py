"""Извлекает продуктовые данные из большого скрипт-блока HTML emex.ru."""
import re, sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

with open(r'data/analytics/emex_debug/products_CAF100493C_CHAMPION.html', encoding='utf-8') as f:
    html = f.read()

# Ищем крупный JSON-массив с продуктами — вокруг "bestPrice"
idx = html.find('"bestPrice"')
if idx < 0:
    print('bestPrice not found')
    exit(1)

# Идём назад, ищем начало массива/объекта
start = idx
depth = 0
for i in range(idx, max(0, idx - 50000), -1):
    c = html[i]
    if c == '}': depth += 1
    elif c == '{':
        depth -= 1
        if depth < 0:
            start = i
            break

# Идём вперёд, ищем конец массива
end = idx
depth = 0
in_str = False
for i in range(start, min(len(html), start + 100000)):
    c = html[i]
    if c == '"' and (i == 0 or html[i-1] != '\\'):
        in_str = not in_str
    if not in_str:
        if c == '[': depth += 1
        elif c == ']': depth -= 1
        if depth < 0 and i > start + 1000:
            end = i
            break

print(f'Контекст вокруг bestPrice: позиции {start}..{end}')
chunk = html[start:start+3000]
print(chunk[:2000])
print('...')

# Пробуем найти полный JSON-массив с productGroups
# Ищем все объекты с "make" и "num"
pattern = r'\{[^{}]*"make"\s*:\s*"([^"]+)"[^{}]*"num"\s*:\s*"([^"]+)"[^{}]*\}'
matches = re.findall(pattern, html)
print(f'\nНайдено пар make+num: {len(matches)}')
seen = set()
for make, num in matches[:30]:
    key = (make, num)
    if key not in seen:
        seen.add(key)
        print(f'  {make:25} {num}')

# Ищем теги с применяемостью
pattern2 = r'\{[^{}]*"tags"\s*:\s*\[([^\]]+)\][^{}]*\}'
tag_matches = re.findall(pattern2, html)
print(f'\nОбъекты с tags: {len(tag_matches)}')
for t in tag_matches[:5]:
    print(f'  tags: {t[:100]}')
