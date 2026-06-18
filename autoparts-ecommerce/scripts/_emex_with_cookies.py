"""Использует куки Chrome для запросов к emex.ru."""
import sys, io, json, requests
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# Попытка 1: browser_cookie3
cookies_dict = {}
try:
    import browser_cookie3
    cookiejar = browser_cookie3.chrome(domain_name='.emex.ru')
    for c in cookiejar:
        cookies_dict[c.name] = c.value
    print(f'browser_cookie3: нашёл {len(cookies_dict)} кукис для emex.ru')
    if cookies_dict:
        print(f'  names: {list(cookies_dict.keys())[:10]}')
except ImportError:
    print('browser_cookie3 не установлен, устанавливаю...')
    import subprocess
    subprocess.run([r'.venv\Scripts\python.exe', '-m', 'pip', 'install', 'browser_cookie3'],
                   capture_output=True)
    import browser_cookie3
    cookiejar = browser_cookie3.chrome(domain_name='.emex.ru')
    for c in cookiejar:
        cookies_dict[c.name] = c.value
    print(f'browser_cookie3: нашёл {len(cookies_dict)} кукис для emex.ru')
except Exception as e:
    print(f'browser_cookie3 error: {e}')

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    'Accept': 'application/json, text/plain, */*',
    'Accept-Language': 'ru-RU,ru;q=0.9',
    'Origin': 'https://emex.ru',
    'Referer': 'https://emex.ru/products/CAF100493C/CHAMPION',
}

session = requests.Session()
session.headers.update(HEADERS)
if cookies_dict:
    session.cookies.update(cookies_dict)

# Тестируем с куками
article = 'CAF100493C'
brand   = 'CHAMPION'

# Получим страницу с куками и посмотрим на state
r = session.get(f'https://emex.ru/products/{article}/{brand}', timeout=15)
print(f'\nPage status: {r.status_code}  size: {len(r.text):,}')

# Парсим state
import re
idx = r.text.find(r'\"details\":{\"makes\"')
if idx >= 0:
    decoded = r.text[max(0,idx-5000):idx+50000].replace('\\"', '"').replace('\\/', '/').replace('\\n', ' ')
    idx2 = decoded.find('"details":{"makes"')
    if idx2 >= 0:
        chunk = decoded[idx2:]
        # Пробуем извлечь детали
        orig_idx = chunk.find('"originals"')
        if orig_idx >= 0:
            orig_ctx = chunk[orig_idx:orig_idx+500]
            print(f'\noriginals section: {orig_ctx}')
        else:
            print('\noriginals not found in state')

        # Проверим user type
        user_idx = chunk.find('"userType"')
        if user_idx >= 0:
            print(f'\nuserType context: {chunk[user_idx:user_idx+100]}')

# Попробуем api1.emex.ru с куками
print('\nПробуем api1.emex.ru с куками:')
endpoints = [
    f'/suggestions/search-suggestions?searchString={article}',
    f'/detail?detailNum={article}&make={brand}',
    f'/detail/originals?detailNum={article}&make={brand}',
    f'/api/detail?detailNum={article}&make={brand}',
    f'/offers?detailNum={article}&make={brand}&cityId=0',
]
for ep in endpoints:
    url = f'https://api1.emex.ru{ep}'
    try:
        r2 = session.get(url, timeout=10)
        print(f'  {r2.status_code:3}  {len(r2.text):7,}  {ep}')
        if r2.status_code == 200 and len(r2.text) > 50:
            try:
                d = r2.json()
                print(f'     keys: {list(d.keys())[:10] if isinstance(d, dict) else f"list[{len(d)}]"}')
            except:
                print(f'     {r2.text[:100]}')
    except Exception as e:
        print(f'  ERR  {ep}: {e}')
