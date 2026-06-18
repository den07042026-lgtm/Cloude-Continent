"""Ищет API эндпоинты в основных чанках emex.ru."""
import re, sys, io, requests, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

HEADERS = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
session = requests.Session()
session.headers.update(HEADERS)

base = 'https://emex.ru/_next/static/chunks/'
big_chunks = [
    '01b8a8f6-fe9f9133fdc41620.js',
    '3304-5b257c45167ceb4e.js',
    '1220-102c66ac221b11f2.js',
    '3683-a61f2c977e4e1f46.js',
    '1197-ab2d97fc9d96e5b1.js',
    '1572-ab8fb60173df62ad.js',
    '3453-f2e6121c8709af12.js',
    '2951-0165c2a1be6c6b87.js',
    '5357-d217a2677f573419.js',
    '244-5cb7a0d0a5bdc430.js',
    'main-17ed643e4c05fa1b.js',
]

def analyze(chunk_name, js):
    size = len(js)
    print(f'\n=== {chunk_name} ({size:,} chars) ===')

    found_anything = False

    # fetchFlow
    if 'fetchFlow' in js or 'FetchFlow' in js:
        matches = re.findall(r'.{0,100}[Ff]etch[Ff]low.{0,200}', js)
        print(f'  fetchFlow x{len(matches)}')
        for m in matches[:3]:
            print(f'    {m[:250]}')
        found_anything = True

    # HTTP POST/GET calls
    http_calls = re.findall(r'(?:axios|fetch|request|http)\.[a-z]+\([^)]{0,200}\)', js, re.IGNORECASE)
    if http_calls:
        print(f'  HTTP calls x{len(http_calls)}:')
        for c in http_calls[:5]:
            print(f'    {c[:200]}')
        found_anything = True

    # Path patterns
    paths = re.findall(r'"(/(?:api|v\d|detail|product|catalog|original|analog|cross)[^"]{3,80})"', js)
    paths2 = re.findall(r"'(/(?:api|v\d|detail|product|catalog|original|analog|cross)[^']{3,80})'", js)
    all_paths = list(set(paths + paths2))
    if all_paths:
        print(f'  Paths x{len(all_paths)}:')
        for p in all_paths[:15]:
            print(f'    {p}')
        found_anything = True

    # Keyword search
    for kw in ['getOriginals', 'getAnalogs', 'getApplicability', 'getCross', 'getOem',
               'loadOriginals', 'loadAnalogs', 'originals', 'crossReferences']:
        if kw.lower() in js.lower():
            matches = re.findall(f'.{{0,80}}{kw}.{{0,200}}', js, re.IGNORECASE)
            if matches:
                print(f'  [{kw}]:')
                print(f'    {matches[0][:300]}')
                found_anything = True

    if not found_anything:
        print(f'  (ничего интересного)')

for chunk in big_chunks:
    url = base + chunk
    cache = f'data/analytics/emex_debug/chunk_{chunk}'
    if os.path.exists(cache):
        with open(cache, encoding='utf-8') as f:
            js = f.read()
        analyze(chunk, js)
    else:
        try:
            r = session.get(url, timeout=20)
            if r.status_code == 200:
                with open(cache, 'w', encoding='utf-8') as f:
                    f.write(r.text)
                analyze(chunk, r.text)
            else:
                print(f'\n=== SKIP {chunk}: {r.status_code} ===')
        except Exception as e:
            print(f'\n=== ERR {chunk}: {e} ===')
