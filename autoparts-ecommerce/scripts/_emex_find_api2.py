"""Ищет API эндпоинты в скачанном JS бандле."""
import re, sys, io, requests, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
}
session = requests.Session()
session.headers.update(HEADERS)

# Загружаем страницу-продукт и находим все JS чанки
js_url = 'https://emex.ru/_next/static/chunks/pages/products/%5BdetailNum%5D/%5Bmake%5D-7ff66ceb0b03312d.js'
r = session.get(js_url, timeout=20)
js = r.text

# Сохраним для анализа
with open('data/analytics/emex_debug/page_js.js', 'w', encoding='utf-8') as f:
    f.write(js)

# Ищем строковые литералы с /
string_literals = re.findall(r'"(/[a-zA-Z][a-zA-Z0-9/_-]{3,50})"', js)
print('String path literals:')
for s in sorted(set(string_literals))[:40]:
    print(f'  {s}')

print()
# Ищем ключевые паттерны в JS
for kw in ['originals', 'analogs', 'replacements', 'oem', 'applicability',
           'getDetail', 'fetchDetail', 'loadDetail', 'getOriginals', 'getCross',
           'dispatch', 'getAnalog']:
    matches = re.findall(f'[^\n]{{0,30}}{kw}[^\n]{{0,100}}', js, re.IGNORECASE)
    if matches:
        print(f'\n[{kw}] x{len(matches)}:')
        for m in matches[:3]:
            print(f'  {m.strip()[:200]}')

# Ищем URL паттерны с параметрами
print('\nTemplate literals (все):')
template_urls = re.findall(r'`[^`\n]{10,200}`', js)
for t in template_urls[:30]:
    print(f'  {t[:150]}')
