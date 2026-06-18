"""Дампает структуру searchResult из search/search."""
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
})

r = s.get(f'https://emex.ru/api/search/search?detailNum={ARTICLE}&make={BRAND}', timeout=15)
d = r.json()

sr = d.get('searchResult', {})
print('=== searchResult keys ===')
print(list(sr.keys()) if isinstance(sr, dict) else type(sr))
print()

if isinstance(sr, dict):
    for k, v in sr.items():
        if isinstance(v, list):
            print(f'{k}: list[{len(v)}]')
            if v:
                print(f'  first: {json.dumps(v[0], ensure_ascii=False)[:300]}')
        elif isinstance(v, dict):
            print(f'{k}: dict  keys={list(v.keys())[:8]}')
            # Рекурсивно для вложенных
            for k2, v2 in v.items():
                if isinstance(v2, list) and v2:
                    print(f'  {k2}: list[{len(v2)}]  first={json.dumps(v2[0], ensure_ascii=False)[:200]}')
                elif isinstance(v2, dict):
                    print(f'  {k2}: dict keys={list(v2.keys())}')
        else:
            print(f'{k}: {str(v)[:100]}')
