"""
Тестирует autodoc.ru API для получения OEM/аналогов.
"""
import sys, io, requests, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import warnings; warnings.filterwarnings('ignore')

s = requests.Session()
s.headers.update({
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36',
    'Accept': 'application/json, text/plain, */*',
    'Accept-Language': 'ru-RU,ru;q=0.9',
})

# Тестируем несколько хорошо известных деталей
tests = [
    ('AMDFA475', 'AMD'),
    ('W81180', 'MANN-FILTER'),
    ('CAF100493C', 'CHAMPION'),
]

for article, brand in tests:
    print(f'\n=== {article} / {brand} ===')

    # autodoc.ru API
    urls = [
        f'https://www.autodoc.ru/api/catalog/article/{article}?brand={brand}&lang=ru',
        f'https://www.autodoc.ru/api/spares/article?number={article}&brand={brand}',
        f'https://www.autodoc.ru/api/v1/articles/{article}/cross-references?brand={brand}',
        f'https://www.autodoc.ru/api/catalog/search?number={article}&lang=ru',
        # Основная страница
        f'https://www.autodoc.ru/auto-parts/analog/{article}',
    ]

    for url in urls:
        try:
            r = s.get(url, timeout=8)
            sz = len(r.text)
            ct = r.headers.get('Content-Type', '')[:30]
            print(f'  {r.status_code:3d}  {sz:7,}  {ct:28}  {url[:80]}')
            if r.status_code == 200 and sz > 100:
                try:
                    d = r.json()
                    if isinstance(d, dict):
                        print(f'    keys: {list(d.keys())[:8]}')
                    elif isinstance(d, list) and d:
                        print(f'    list[{len(d)}] first: {str(d[0])[:100]}')
                except Exception:
                    print(f'    html: {r.text[:100]}')
        except Exception as e:
            print(f'  ERR  {url[:70]}: {str(e)[:40]}')
