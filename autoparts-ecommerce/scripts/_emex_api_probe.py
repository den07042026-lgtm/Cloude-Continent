"""Пробует найти рабочий API эндпоинт emex.ru для данных о запчасти."""
import requests, sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    'Accept': 'application/json, text/plain, */*',
    'Accept-Language': 'ru-RU,ru;q=0.9',
    'Referer': 'https://emex.ru/products/CAF100493C/CHAMPION',
    'x-requested-with': 'XMLHttpRequest',
}

session = requests.Session()
session.headers.update(HEADERS)

num = 'CAF100493C'
make = 'CHAMPION'
make_id = '626'  # из state.details.makeId если есть

candidates = [
    f'/api/products/{num}/{make}',
    f'/api/products/{num}/{make}/originals',
    f'/api/products/{num}/{make}/analogs',
    f'/api/products/{num}/{make}/replacements',
    f'/api/products/{num}/{make}/details',
    f'/api/details/{num}/{make}',
    f'/api/details?num={num}&make={make}',
    f'/api/detail?detailNum={num}&make={make}',
    f'/api/cross-refs?num={num}&make={make}',
    f'/api/v1/products/{num}/{make}',
    f'/api/v2/products/{num}/{make}',
    f'/api/getDetailInfo?detailNum={num}&make={make}',
    f'/api/getDetailInfo?detailNum={num}&make={make}&locationId=0',
    f'/api/getOriginals?detailNum={num}&make={make}',
    f'/api/getAnalogs?detailNum={num}&make={make}',
    f'/api/applicability/{num}/{make}',
    f'/api/products/applicability?num={num}&make={make}',
    # Попробуем с makeId
    f'/api/products/{num}/{make_id}',
    f'/api/details/{num}/{make_id}',
]

base = 'https://emex.ru'
for path in candidates:
    url = base + path
    try:
        r = session.get(url, timeout=10)
        status = r.status_code
        size = len(r.text)
        ct = r.headers.get('Content-Type', '')[:40]
        if status == 200 and size > 50:
            print(f'OK {status} {size:7d}  {ct:40}  {path}')
            # Попробуем распарсить
            try:
                d = r.json()
                print(f'     JSON keys: {list(d.keys())[:8]}')
                # Сохраним
                fname = path.replace('/', '_').replace('?', '_').replace('&', '_')[:50]
                with open(f'data/analytics/emex_debug/probe{fname}.json', 'w', encoding='utf-8') as f:
                    json.dump(d, f, ensure_ascii=False, indent=2)
            except:
                print(f'     Not JSON: {r.text[:100]}')
        else:
            print(f'   {status} {size:5d}  {path}')
    except Exception as e:
        print(f'   ERR  {path}  ->  {str(e)[:60]}')
