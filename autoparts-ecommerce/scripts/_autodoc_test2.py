"""Тестирует autodoc.ru API с отключённой SSL верификацией."""
import requests, sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import warnings
warnings.filterwarnings('ignore')

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    'Accept': 'application/json, text/plain, */*',
    'Accept-Language': 'ru-RU,ru;q=0.9',
    'Origin': 'https://www.autodoc.ru',
    'Referer': 'https://www.autodoc.ru/catalog',
}
session = requests.Session()
session.headers.update(HEADERS)

article = 'CAF100493C'
brand   = 'CHAMPION'

endpoints = [
    # catalog.autodoc.ru
    (f'https://catalog.autodoc.ru/api/references?number={article}&lang=ru', False),
    (f'https://catalog.autodoc.ru/api/brands?number={article}&lang=ru', False),
    (f'https://catalog.autodoc.ru/api/references/{article}?lang=ru', False),
    # webapi.autodoc.ru
    ('https://webapi.autodoc.ru/api/manufacturers', True),
    (f'https://webapi.autodoc.ru/api/spares/228/{article}', True),   # 228 = Champion brand id ?
    (f'https://webapi.autodoc.ru/api/spares/{article}', True),
    # autodoc search
    (f'https://www.autodoc.ru/catalog/number/{article}', True),
    # parts.autodoc
    (f'https://parts.autodoc.ru/api/search?number={article}&lang=ru', False),
    # Прямой API для кросс-ссылок
    (f'https://catalog.autodoc.ru/api/autoproducers?lang=ru', False),
]

for url, verify in endpoints:
    print(f'\n--- {url}')
    try:
        r = session.get(url, timeout=15, verify=verify)
        ct = r.headers.get('Content-Type', '')[:60]
        print(f'    {r.status_code}  {len(r.text):,}  {ct}')
        if r.status_code == 200 and len(r.text) > 50:
            try:
                d = r.json()
                if isinstance(d, list):
                    print(f'    list[{len(d)}]')
                    if d: print(f'    [0]: {json.dumps(d[0], ensure_ascii=False)[:200]}')
                elif isinstance(d, dict):
                    print(f'    keys: {list(d.keys())[:10]}')
                    print(f'    preview: {json.dumps(d, ensure_ascii=False)[:300]}')
                fname = url.replace('https://', '').replace('/', '_').replace('?', '_').replace('&', '_')[:70]
                with open(f'data/analytics/emex_debug/autodoc2_{fname}.json', 'w', encoding='utf-8') as f:
                    json.dump(d, f, ensure_ascii=False, indent=2)
                print(f'    Saved!')
            except Exception as e:
                print(f'    {r.text[:300]}')
    except Exception as e:
        print(f'    ERROR: {str(e)[:100]}')
