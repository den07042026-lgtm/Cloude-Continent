"""
Проверяет полную структуру searchResult из search/search API.
Смотрим есть ли originals/analogs/replacements при разных параметрах.
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

def check_url(label, url):
    r = s.get(url, timeout=15)
    d = r.json()
    sr = d.get('searchResult', {})
    print(f'\n=== {label} ({len(r.text):,} байт) ===')
    print(f'searchResult keys: {list(sr.keys()) if isinstance(sr, dict) else type(sr)}')
    for k in ['originals', 'analogs', 'replacements', 'oem', 'cross']:
        if k in sr:
            v = sr[k]
            cnt = len(v) if isinstance(v, list) else str(v)[:50]
            print(f'  {k}: {cnt}')
            if isinstance(v, list) and v:
                print(f'    first: {json.dumps(v[0], ensure_ascii=False)[:200]}')

# AMD фильтр — aftermarket
ARTICLE = 'AMDFA475'
BRAND   = 'AMD'

check_url('Без координат',
    f'https://emex.ru/api/search/search?detailNum={ARTICLE}&make={BRAND}')

check_url('С координатами Москвы',
    f'https://emex.ru/api/search/search?detailNum={ARTICLE}&make={BRAND}&leftBottomLatitude=55.5&leftBottomLongitude=37.0&rightTopLatitude=55.9&rightTopLongitude=37.9')

# Пробуем Honda OEM номер — там должны быть cross-references
check_url('Honda OEM номер',
    'https://emex.ru/api/search/search?detailNum=15400-RBA-F01&make=Honda')

# MANN filter
check_url('MANN W811/80',
    'https://emex.ru/api/search/search?detailNum=W81180&make=Mann')

# Пробуем другой параметр
check_url('packet=-1 format',
    f'https://emex.ru/api/search/search?detailNum={ARTICLE}&packet=-1')
