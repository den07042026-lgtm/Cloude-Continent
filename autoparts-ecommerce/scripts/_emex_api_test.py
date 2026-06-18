"""Ищет API-эндпоинты emex.ru для получения данных о запчасти."""
import json, re, sys, io, requests
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    'Accept': 'application/json, text/plain, */*',
    'Accept-Language': 'ru-RU,ru;q=0.9',
    'Referer': 'https://emex.ru/products/CAF100493C/CHAMPION',
}

# buildId из HTML (строка iL4MsGAEM_6ZRH-ENNYPM)
with open(r'data/analytics/emex_debug/products_CAF100493C_CHAMPION.html', encoding='utf-8') as f:
    html = f.read()
m = re.search(r'"buildId"\s*:\s*"([^"]+)"', html)
build_id = m.group(1) if m else 'unknown'
print(f'buildId: {build_id}')

article = 'CAF100493C'
brand   = 'CHAMPION'

# Next.js data endpoints
urls_to_try = [
    f'https://emex.ru/_next/data/{build_id}/products/{article}/{brand}.json',
    f'https://emex.ru/api/getDetailInfo?detailNum={article}&make={brand}',
    f'https://emex.ru/api/products/{article}/{brand}',
    f'https://emex.ru/api/detail?number={article}&brand={brand}',
    f'https://emex.ru/api/cross?number={article}&brand={brand}',
    f'https://emex.ru/api/v1/detail/{article}?brand={brand}',
]

session = requests.Session()
session.headers.update(HEADERS)

for url in urls_to_try:
    print(f'\n--- {url}')
    try:
        r = session.get(url, timeout=15)
        print(f'    Status: {r.status_code}  size={len(r.text)}')
        if r.status_code == 200 and len(r.text) > 100:
            # попытка распарсить JSON
            try:
                data = r.json()
                print(f'    JSON keys: {list(data.keys())[:10]}')
                # сохраним
                slug = url.replace('https://emex.ru/', '').replace('/', '_')[:60]
                with open(f'data/analytics/emex_debug/api_{slug}.json', 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                print(f'    Сохранено: api_{slug}.json')
            except:
                print(f'    HTML: {r.text[:200]}')
    except Exception as e:
        print(f'    ERROR: {e}')
