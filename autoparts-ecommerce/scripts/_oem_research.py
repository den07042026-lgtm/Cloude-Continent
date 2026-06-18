"""Исследует дополнительные источники OEM-номеров."""
import sys, io, requests, json, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import warnings; warnings.filterwarnings('ignore')

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    'Accept': 'application/json',
    'Accept-Language': 'ru-RU,ru;q=0.9',
}
s = requests.Session()
s.headers.update(HEADERS)

article = 'CAF100493C'
brand   = 'CHAMPION'

# Различные источники OEM
tests = [
    # TecDoc-based
    ('GET', 'https://www.autodoc.ru/api/spares/search', {'number': article, 'lang': 'ru'}, True),
    ('GET', 'https://www.autodoc.ru/catalog', {'part_number': article}, True),
    # Parts.ru
    ('GET', f'https://www.parts.ru/setApi.html?item_id={article}&format=json', {}, True),
    ('GET', f'https://www.parts.ru/api/v1/parts?number={article}&brand={brand}', {}, True),
    # ZipNavia
    ('GET', f'https://zipnavia.ru/catalog/search/?q={article}', {}, True),
    # PartsCatalog
    ('GET', f'https://partscatalog.ru/api/search?q={article}', {}, True),
    # Tecdoc open
    ('GET', f'https://api.tecdoc.net/v1/articles?number={article}', {}, True),
    # ixat.ru
    ('GET', f'https://ixat.ru/search?article={article}&brand={brand}', {}, True),
]

for method, url, params, verify in tests:
    try:
        r = s.get(url, params=params, timeout=8, verify=verify)
        ct = r.headers.get('Content-Type', '')[:40]
        sz = len(r.text)
        if r.status_code == 200 and sz > 100:
            print(f'OK {r.status_code} {sz:8,}  {ct:35}  {url}')
            try:
                d = r.json()
                k = list(d.keys())[:8] if isinstance(d, dict) else f'list[{len(d)}]'
                print(f'   JSON: {k}')
            except:
                print(f'   HTML/text: {r.text[:200]}')
        else:
            print(f'   {r.status_code} {sz:5}  {url}')
    except Exception as e:
        print(f'   ERR  {url[:60]}: {str(e)[:60]}')

print()
# Попробуем laximo.ru (очень популярный в России TecDoc-каталог)
laximo_tests = [
    f'https://laximo.ru/catalog/search?q={article}',
    f'https://laximo.ru/api/cross?article={article}&brand={brand}',
]
for url in laximo_tests:
    try:
        r = s.get(url, timeout=8, verify=False)
        print(f'{r.status_code} {len(r.text):,}  {url}')
    except Exception as e:
        print(f'ERR {url}: {e}')
