"""Fetches Next.js data endpoint for product page and explores playwright longer."""
import sys, io, json, re, requests
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    'Accept': 'application/json, text/plain, */*',
    'Accept-Language': 'ru-RU,ru;q=0.9',
    'Referer': 'https://emex.ru/',
}
session = requests.Session()
session.headers.update(HEADERS)

# _next/data endpoint
build_id = 'iL4MsGAEM_6ZRH-ENNYPM'
article  = 'CAF100493C'
brand    = 'CHAMPION'

urls = [
    f'https://emex.ru/_next/data/{build_id}/products/{article}/{brand}.json',
    f'https://emex.ru/_next/data/{build_id}/products/{article}/{brand.lower()}.json',
    f'https://emex.ru/_next/data/{build_id}/products/{article}.json',
    # index and help for comparison
    f'https://emex.ru/_next/data/{build_id}/help.json',
    f'https://emex.ru/_next/data/{build_id}/index.json',
]

for url in urls:
    print(f'\n--- {url}')
    try:
        r = session.get(url, timeout=15)
        ct = r.headers.get('Content-Type', '')[:50]
        print(f'    {r.status_code}  {len(r.text):,}  {ct}')
        if r.status_code == 200 and len(r.text) > 100:
            try:
                d = r.json()
                keys = list(d.keys())
                print(f'    keys: {keys[:10]}')
                pp = d.get('pageProps', {})
                if pp:
                    print(f'    pageProps keys: {list(pp.keys())[:15]}')
                fname = url.split('/')[-1].replace('.json', '')
                with open(f'data/analytics/emex_debug/nextdata_{fname}.json', 'w', encoding='utf-8') as f:
                    json.dump(d, f, ensure_ascii=False, indent=2)
                print(f'    Saved!')
            except Exception as e:
                print(f'    JSON error: {e}')
                print(f'    {r.text[:300]}')
    except Exception as e:
        print(f'    ERROR: {e}')
