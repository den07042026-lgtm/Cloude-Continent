"""Ищет API эндпоинты в JS-бандле страницы продукта emex.ru."""
import re, sys, io, requests, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    'Accept': '*/*',
    'Referer': 'https://emex.ru/',
}
session = requests.Session()
session.headers.update(HEADERS)

# JS файл для страницы /products/[detailNum]/[make]
js_url = 'https://emex.ru/_next/static/chunks/pages/products/%5BdetailNum%5D/%5Bmake%5D-7ff66ceb0b03312d.js'
print(f'Fetching: {js_url}')

try:
    r = session.get(js_url, timeout=20)
    print(f'Status: {r.status_code}  size: {len(r.text)}')
    js = r.text

    # Ищем API пути
    api_patterns = re.findall(r'["\`](/api/[^"\'`\s]{3,80})["\`]', js)
    print(f'\nAPI paths found: {len(api_patterns)}')
    for p in sorted(set(api_patterns))[:30]:
        print(f'  {p}')

    # Ищем fetch/axios вызовы
    fetch_calls = re.findall(r'fetch\(["\`]([^"\'`]+)["\`]', js)
    print(f'\nfetch calls: {len(fetch_calls)}')
    for p in fetch_calls[:20]:
        print(f'  {p}')

    # Ищем URL patterns с параметрами
    url_patterns = re.findall(r'`[^`]*\$\{[^`]*\}[^`]*`', js)
    print(f'\nTemplate literal URLs: {len(url_patterns)}')
    for p in url_patterns[:20]:
        if 'api' in p.lower() or 'product' in p.lower() or 'detail' in p.lower():
            print(f'  {p[:150]}')

    # Ищем строки с "original", "analog", "cross"
    for kw in ['original', 'analog', 'cross', 'oem', 'applicab', 'applies']:
        matches = re.findall(f'.{{0,50}}{kw}.{{0,100}}', js, re.IGNORECASE)
        if matches:
            print(f'\n[{kw}]:')
            for m in matches[:3]:
                print(f'  {m[:200]}')

except Exception as e:
    print(f'ERROR: {e}')
    print('Trying to fetch from saved HTML instead...')
    # Попробуем другой подход — посмотрим на main JS чанки
    with open('data/analytics/emex_debug/products_CAF100493C_CHAMPION.html', encoding='utf-8') as f:
        html = f.read()
    # Найдём все /api/ URLs в HTML
    api_in_html = re.findall(r'["\'/]api/[a-zA-Z0-9/_?&=.%-]{3,80}', html)
    print(f'API URLs in HTML: {len(api_in_html)}')
    for u in sorted(set(api_in_html))[:20]:
        print(f'  {u}')
