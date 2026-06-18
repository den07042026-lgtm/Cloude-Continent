"""
Пробует прямые HTTP-запросы к api1.emex.ru для получения originals/analogs.
"""
import sys, io, requests, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import warnings; warnings.filterwarnings('ignore')

ARTICLE = 'AMDFA475'
BRAND   = 'AMD'

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36',
    'Accept': 'application/json, text/plain, */*',
    'Accept-Language': 'ru-RU,ru;q=0.9',
    'Referer': f'https://emex.ru/products/{ARTICLE}/{BRAND}',
    'Origin': 'https://emex.ru',
}
s = requests.Session()
s.headers.update(HEADERS)

candidates = [
    # Разные варианты эндпоинтов для originals/analogs
    f'https://api1.emex.ru/details/originals?detailNum={ARTICLE}&brand={BRAND}',
    f'https://api1.emex.ru/details/analogs?detailNum={ARTICLE}&brand={BRAND}',
    f'https://api1.emex.ru/details?detailNum={ARTICLE}&brand={BRAND}',
    f'https://api1.emex.ru/product/originals?num={ARTICLE}&make={BRAND}',
    f'https://api1.emex.ru/product/analogs?num={ARTICLE}&make={BRAND}',
    f'https://api1.emex.ru/cross?num={ARTICLE}&brand={BRAND}',
    f'https://api1.emex.ru/products/{ARTICLE}/{BRAND}/originals',
    f'https://api1.emex.ru/products/{ARTICLE}/{BRAND}/analogs',
    f'https://api1.emex.ru/search/originals?num={ARTICLE}',
    f'https://api1.emex.ru/oe?num={ARTICLE}&brand={BRAND}',
    # Формат /f
    f'https://emex.ru/api/details?detailNum={ARTICLE}&packet=-1',
    f'https://emex.ru/api/originals?detailNum={ARTICLE}',
    f'https://emex.ru/api/analogs?detailNum={ARTICLE}',
    f'https://api1.emex.ru/v2/details/{ARTICLE}?brand={BRAND}',
    f'https://api1.emex.ru/details/full?num={ARTICLE}&brand={BRAND}',
]

for url in candidates:
    try:
        r = s.get(url, timeout=8, verify=False)
        sz = len(r.text)
        ct = r.headers.get('Content-Type', '')[:40]
        print(f'{r.status_code:3d}  {sz:7,}  {ct:35}  {url}')
        if r.status_code == 200 and sz > 50:
            try:
                d = r.json()
                keys = list(d.keys())[:6] if isinstance(d, dict) else f'list[{len(d)}]'
                print(f'   JSON keys: {keys}')
            except Exception:
                print(f'   text: {r.text[:100]}')
    except Exception as e:
        print(f'ERR  {url[:80]}: {str(e)[:50]}')
