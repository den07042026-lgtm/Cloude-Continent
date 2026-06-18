"""Исследует api1.emex.ru — реальный API бэкенд emex.ru."""
import requests, sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    'Accept': 'application/json, text/plain, */*',
    'Accept-Language': 'ru-RU,ru;q=0.9',
    'Origin': 'https://emex.ru',
    'Referer': 'https://emex.ru/products/CAF100493C/CHAMPION',
}
session = requests.Session()
session.headers.update(HEADERS)

article = 'CAF100493C'
brand   = 'CHAMPION'
make_id = '626'
city_id = '0'  # общий

# Базовые URL на api1.emex.ru
base = 'https://api1.emex.ru'

endpoints = [
    # Suggestions (уже знаем что работает)
    f'/suggestions/search-suggestions?searchString={article}',
    # Detail data
    f'/detail/info?detailNum={article}&make={brand}',
    f'/detail/info?detailNum={article}&make={brand}&cityId={city_id}',
    f'/detail/originals?detailNum={article}&make={brand}',
    f'/detail/analogs?detailNum={article}&make={brand}',
    f'/detail/applicability?detailNum={article}&make={brand}',
    f'/detail/makes?detailNum={article}',
    f'/detail?detailNum={article}&make={brand}',
    f'/product/{article}/{brand}',
    f'/product/{article}/{make_id}',
    f'/products/{article}/{brand}',
    f'/products?detailNum={article}&make={brand}',
    f'/cross-references?detailNum={article}&make={brand}',
    f'/oem?detailNum={article}&make={brand}',
    f'/details/{article}/{brand}',
    f'/details?num={article}&make={brand}',
    # Специфичные
    f'/detail/info?detailNum={article}&make={brand}&locationId=0',
    f'/detail/productInfo?detailNum={article}&make={brand}',
    f'/catalog/detail?detailNum={article}&make={brand}',
    # v1, v2, v3
    f'/v1/detail?detailNum={article}&make={brand}',
    f'/v2/detail?detailNum={article}&make={brand}',
]

print(f'Тестирование {len(endpoints)} эндпоинтов на api1.emex.ru\n')
found_any = False
for path in endpoints:
    url = base + path
    try:
        r = session.get(url, timeout=10)
        ct = r.headers.get('Content-Type', '')[:40]
        size = len(r.text)
        if r.status_code == 200 and size > 20:
            print(f'OK! {size:8,}  {ct:40}  {path}')
            try:
                d = r.json()
                if isinstance(d, list):
                    print(f'   list[{len(d)}]  first: {json.dumps(d[0] if d else {}, ensure_ascii=False)[:200]}')
                elif isinstance(d, dict):
                    print(f'   dict keys: {list(d.keys())[:10]}')
                slug = path.replace('/', '_').replace('?', '_').replace('&', '_')[:50]
                with open(f'data/analytics/emex_debug/api1{slug}.json', 'w', encoding='utf-8') as f:
                    json.dump(d, f, ensure_ascii=False, indent=2)
                print(f'   SAVED: api1{slug}.json')
                found_any = True
            except:
                print(f'   Not JSON: {r.text[:200]}')
        else:
            print(f'    {r.status_code:3}  {size:5}  {path}')
    except Exception as e:
        print(f'    ERR  {str(e)[:60]}  {path}')

if not found_any:
    print('\nНичего не найдено. Пробуем через Playwright перехватить больше запросов...')
