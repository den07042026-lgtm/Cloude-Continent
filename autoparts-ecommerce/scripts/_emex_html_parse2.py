"""Углублённый анализ HTML emex.ru — ищем JSON с данными о запчасти."""
import re, sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

with open(r'data/analytics/emex_debug/products_CAF100493C_CHAMPION.html', encoding='utf-8') as f:
    html = f.read()

# Ищем крупные inline-скрипты с данными
scripts = re.findall(r'<script[^>]*>([\s\S]{200,}?)</script>', html)
print(f'Скрипт-блоков >{200} символов: {len(scripts)}')
for i, s in enumerate(scripts):
    print(f'\n  [{i}] {len(s)} chars  начало: {s[:200]}')

print()
# Ищем self.__next_f или другие hydration-форматы
hydration = re.findall(r'self\.__next_f\.push\(\[([^\]]+)\]\)', html)
print(f'next_f hydration блоков: {len(hydration)}')
for h in hydration[:3]:
    print(f'  {h[:300]}')

# Ищем конкретные строки с данными о товаре
idx = html.find('"аналоги"')
if idx == -1:
    idx = html.find('\\u0430\\u043d\\u0430\\u043b\\u043e\\u0433\\u0438')
print(f'\nПозиция "аналоги": {idx}')
if idx > 0:
    print(html[max(0,idx-200):idx+500])

# Ищем productGroups / crossReferences / applies
for keyword in ['"productGroups"', '"crossReferences"', '"applies"', '"vehicles"',
                '"oemNumbers"', '"makes"', '"brands"', '"details"', '"group"',
                'detailNum', 'bestPrice', 'parsedUrl']:
    cnt = html.count(keyword)
    if cnt > 0:
        idx = html.index(keyword)
        print(f'\n[{keyword}] x{cnt} — контекст:')
        print(html[max(0,idx-50):idx+300])
