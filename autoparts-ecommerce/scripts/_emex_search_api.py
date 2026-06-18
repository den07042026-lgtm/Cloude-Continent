"""
Тестирует найденный эндпоинт emex.ru/api/search/search
и другие варианты для получения originals/analogs.
"""
import sys, io, requests, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import warnings; warnings.filterwarnings('ignore')

ARTICLE = 'AMDFA475'
BRAND   = 'AMD'

s = requests.Session()
s.headers.update({
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36',
    'Accept': 'application/json, text/plain, */*',
    'Accept-Language': 'ru-RU,ru;q=0.9',
    'Referer': f'https://emex.ru/products/{ARTICLE}/{BRAND}',
    'Origin': 'https://emex.ru',
})

tests = [
    # Найденный эндпоинт (без координат)
    f'https://emex.ru/api/search/search?detailNum={ARTICLE}&make={BRAND}',
    # С координатами Москвы
    f'https://emex.ru/api/search/search?detailNum={ARTICLE}&make={BRAND}&leftBottomLatitude=55.5&leftBottomLongitude=37.0&rightTopLatitude=55.9&rightTopLongitude=37.9',
    # Варианты для originals/analogs
    f'https://emex.ru/api/search/originals?detailNum={ARTICLE}&make={BRAND}',
    f'https://emex.ru/api/search/analogs?detailNum={ARTICLE}&make={BRAND}',
    f'https://emex.ru/api/details/originals?detailNum={ARTICLE}&make={BRAND}',
    f'https://emex.ru/api/details/analogs?detailNum={ARTICLE}&make={BRAND}',
    f'https://emex.ru/api/details?detailNum={ARTICLE}&make={BRAND}',
    # f format
    f'https://emex.ru/f?detailNum={ARTICLE}&packet=-1',
    f'https://emex.ru/api/search/search?detailNum={ARTICLE}&packet=-1',
    # Вариант без make
    f'https://emex.ru/api/search/search?detailNum={ARTICLE}',
]

for url in tests:
    try:
        r = s.get(url, timeout=10, verify=False)
        sz  = len(r.text)
        ct  = r.headers.get('Content-Type', '')[:40]
        print(f'{r.status_code:3d}  {sz:8,}  {ct:35}  {url[:90]}')
        if r.status_code == 200 and sz > 100:
            try:
                d = r.json()
                if isinstance(d, dict):
                    keys = list(d.keys())
                    print(f'   keys: {keys}')
                    # Ищем originals/analogs
                    for k in ['originals', 'analogs', 'oe', 'oem', 'cross', 'searchResult']:
                        if k in d:
                            val = d[k]
                            cnt = len(val) if isinstance(val, list) else '?'
                            print(f'   {k}: {cnt} items  {str(val)[:200]}')
                elif isinstance(d, list):
                    print(f'   list[{len(d)}]  first: {str(d[0])[:200] if d else "empty"}')
            except Exception:
                print(f'   HTML/text: {r.text[:150]}')
    except Exception as e:
        print(f'ERR  {url[:80]}: {str(e)[:60]}')
    print()
