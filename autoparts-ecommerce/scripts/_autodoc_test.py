"""Тестирует API autodoc.ru для получения OEM/применяемости/аналогов."""
import requests, sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    'Accept': 'application/json, text/plain, */*',
    'Accept-Language': 'ru-RU,ru;q=0.9',
    'Origin': 'https://www.autodoc.ru',
    'Referer': 'https://www.autodoc.ru/',
}

session = requests.Session()
session.headers.update(HEADERS)

# Тестовый артикул: Champion CAF100493C
article = 'CAF100493C'
brand = 'CHAMPION'

# API autodoc.ru (catalog.autodoc.ru)
endpoints = [
    f'https://catalog.autodoc.ru/api/references?number={article}&lang=ru',
    f'https://catalog.autodoc.ru/api/brands?number={article}&lang=ru',
    f'https://catalog.autodoc.ru/api/references/{article}?lang=ru',
    f'https://webapi.autodoc.ru/api/manufacturers',
    f'https://webapi.autodoc.ru/api/catalogs',
    # Поиск по артикулу
    f'https://www.autodoc.ru/api/search?query={article}',
    f'https://www.autodoc.ru/api/parts?number={article}&brand={brand}',
    # Основной каталог
    f'https://catalog.autodoc.ru/api/v1/products?number={article}&lang=ru',
]

for url in endpoints:
    print(f'\n--- {url}')
    try:
        r = session.get(url, timeout=15)
        ct = r.headers.get('Content-Type', '')[:50]
        print(f'    {r.status_code}  {len(r.text):,}  {ct}')
        if r.status_code == 200 and len(r.text) > 50:
            try:
                d = r.json()
                if isinstance(d, list):
                    print(f'    list[{len(d)}] first: {json.dumps(d[0], ensure_ascii=False)[:200] if d else "empty"}')
                elif isinstance(d, dict):
                    print(f'    keys: {list(d.keys())[:10]}')
                # Сохраним первый успешный
                fname = url.replace('https://', '').replace('/', '_').replace('?', '_').replace('&', '_')[:60]
                with open(f'data/analytics/emex_debug/autodoc_{fname}.json', 'w', encoding='utf-8') as f:
                    json.dump(d, f, ensure_ascii=False, indent=2)
                print(f'    Saved!')
            except:
                print(f'    Not JSON: {r.text[:300]}')
    except Exception as e:
        print(f'    ERROR: {e}')
