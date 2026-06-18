"""Скачивает JS бандлы и ищет originals/analogs в обфусцированном JS."""
import sys, io, re, requests
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import warnings; warnings.filterwarnings('ignore')

s = requests.Session()
s.headers['User-Agent'] = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36'

# Список конкретных чанков из предыдущего Playwright логирования
KNOWN_CHUNKS = [
    '01b8a8f6-fe9f9133fdc41620.js',
    '4998-862f9a3262b93671.js',
    '4002-b5173d6b4d1b175d.js',
    '5127-e890ff479cba7085.js',
    '3304-5b257c45167ceb4e.js',
    '1292-e25edef9d5bb5987.js',
    '1070-4f548bebf04fe0ae.js',
    '1220-102c66ac221b11f2.js',
    '5856-64c6f6dca55157b8.js',
    '3683-a61f2c977e4e1f46.js',
    '1197-ab2d97fc9d96e5b1.js',
    '3840-d537a295893fe4ec.js',
    '1572-ab8fb60173df62ad.js',
    '3453-f2e6121c8709af12.js',
    '2951-0165c2a1be6c6b87.js',
    '5357-d217a2677f573419.js',
    '244-5cb7a0d0a5bdc430.js',
    'pages/products/%5BdetailNum%5D/%5Bmake%5D-7ff66ceb0b03312d.js',
    '6206.403b4f1ed98855e1.js',
    'framework-bcd2a2c8050aa341.js',
    'main-17ed643e4c05fa1b.js',
    'pages/_app-7d8986f63512d348.js',
]

KEYWORDS = ['originals', '"analogs"', 'fetchDetails', 'isAllowFetch',
            '"originals"', 'getOriginals', 'getAnalogs', 'loadOriginals', 'loadAnalogs']

total_size = 0
for chunk in KNOWN_CHUNKS:
    url = f'https://emex.ru/_next/static/chunks/{chunk}'
    try:
        r = s.get(url, timeout=15)
        if r.status_code != 200:
            continue
        content = r.text
        total_size += len(content)
        found = [kw for kw in KEYWORDS if kw in content]
        if found:
            print(f'\n=== {chunk} ({len(content):,} байт)  HIT: {found} ===')
            for kw in found:
                idx = content.find(kw)
                ctx = content[max(0, idx-80):idx+150].replace('\n', ' ')
                print(f'  [{kw}] {ctx[:200]}')
        else:
            print(f'  {chunk[:50]:50s}  {len(content):8,} байт  — нет совпадений')
    except Exception as e:
        print(f'  ERR {chunk}: {e}')

print(f'\nВсего загружено: {total_size:,} байт')
