"""Ищет API эндпоинты в основных JS чанках emex.ru."""
import re, sys, io, requests, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
}
session = requests.Session()
session.headers.update(HEADERS)

# Основные чанки из HTML
chunks = [
    '4998-862f9a3262b93671.js',
    '4002-b5173d6b4d1b175d.js',
    '5127-e890ff479cba7085.js',
    '1292-e25edef9d5bb5987.js',
    '1070-4f548bebf04fe0ae.js',
    '3840-d537a295893fe4ec.js',
]

base = 'https://emex.ru/_next/static/chunks/'

for chunk in chunks:
    url = base + chunk
    cache_path = f'data/analytics/emex_debug/chunk_{chunk}'
    if os.path.exists(cache_path):
        with open(cache_path, encoding='utf-8') as f:
            js = f.read()
    else:
        try:
            r = session.get(url, timeout=20)
            if r.status_code == 200:
                js = r.text
                with open(cache_path, 'w', encoding='utf-8') as f:
                    f.write(js)
            else:
                print(f'  SKIP {chunk}: {r.status_code}')
                continue
        except Exception as e:
            print(f'  ERR {chunk}: {e}')
            continue

    print(f'\n=== {chunk} ({len(js):,} chars) ===')

    # Ищем fetchFlow
    if 'fetchFlow' in js:
        matches = re.findall(r'.{0,100}fetchFlow.{0,200}', js)
        print(f'  fetchFlow x{len(matches)}:')
        for m in matches[:5]:
            print(f'    {m[:250]}')

    # Ищем HTTP запросы
    for pattern in [r'"(https?://[^"]{10,100})"', r"'(https?://[^']{10,100})'",
                    r'`(https?://[^`]{10,100})`']:
        urls = re.findall(pattern, js)
        if urls:
            print(f'  HTTP URLs x{len(urls)}:')
            for u in sorted(set(urls))[:10]:
                print(f'    {u}')

    # Ищем /v1/ /v2/ /api/ endpoints
    api_paths = re.findall(r'"(/(?:api|v1|v2|products|details|catalog)[^"]{3,80})"', js)
    if api_paths:
        print(f'  API paths x{len(api_paths)}:')
        for p in sorted(set(api_paths))[:15]:
            print(f'    {p}')

    # Ищем ключевые слова связанные с OEM
    for kw in ['original', 'analog', 'replacement', 'applicab', 'getOem', 'crossRef']:
        matches = re.findall(f'.{{0,40}}{kw}.{{0,100}}', js, re.IGNORECASE)
        if matches:
            print(f'  [{kw}] x{len(matches)} first match: {matches[0][:200]}')
