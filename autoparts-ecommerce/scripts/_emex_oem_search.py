"""Ищет OEM-номера и применяемость в HTML emex.ru разными методами."""
import re, sys, io, requests, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    'Accept-Language': 'ru-RU,ru;q=0.9',
}
session = requests.Session()
session.headers.update(HEADERS)

with open('data/analytics/emex_debug/products_CAF100493C_CHAMPION.html', encoding='utf-8') as f:
    html = f.read()

# Разэкранируем
decoded = html.replace('\\"', '"').replace('\\/', '/').replace('\\n', ' ')

# 1. Ищем OEM паттерны (Ford: AV6N, 1372773; Mazda: LF10, PE01; VAG: etc.)
oem_patterns = [
    r'[A-Z0-9]{2,3}[-\s]?[A-Z0-9]{4,6}[-\s][A-Z0-9]{2,4}',  # Ford style AV6N-9601-AA
    r'\d{7,10}',  # numeric OEM like 1372773
    r'[A-Z]{2}\d{2}-\d{2,4}-[A-Z]\d{2}',  # Mazda style LF10-13-Z40
]
print('=== OEM patterns search ===')
all_oem = set()
for pat in oem_patterns:
    matches = re.findall(pat, decoded)
    for m in matches[:5]:
        all_oem.add(m)
print(f'Possible OEM numbers (sample): {list(all_oem)[:20]}')

# 2. Ищем конкретные Ford/Mazda OEM что знаем
known_oem = ['1372773', 'AV6N', 'LF10-13', 'CM5Z', '30864621', '4H23']
for oem in known_oem:
    if oem in decoded:
        idx = decoded.index(oem)
        print(f'FOUND {oem}: {decoded[max(0,idx-50):idx+100]}')

# 3. Попробуем загрузить страницу с accept заголовком для JSON
print('\n=== Попытка получить данные через Accept: application/json ===')
for url in [
    'https://emex.ru/products/CAF100493C/CHAMPION',
    'https://emex.ru/products/CAF100493C/champion',
]:
    h = dict(HEADERS)
    h['Accept'] = 'application/json'
    try:
        r = session.get(url, headers=h, timeout=15)
        ct = r.headers.get('Content-Type', '')[:60]
        print(f'{r.status_code}  {ct}  {len(r.text):,}  {url}')
    except Exception as e:
        print(f'ERROR: {e}')

# 4. Попробуем загрузить страницу catalog/original2 для конкретного автомобиля
print('\n=== exist.ru для OEM данных ===')
# Попробуем exist.ru который показывает OEM cross-reference
urls_exist = [
    'https://www.exist.ru/Part/SearchByOem/?oem=CAF100493C&BrandId=0',
    'https://exist.ru/api/part/search?q=CAF100493C',
]
for url in urls_exist:
    try:
        r = session.get(url, timeout=10)
        print(f'{r.status_code}  {len(r.text):,}  {url}')
        if r.status_code == 200 and len(r.text) > 100:
            print(f'  {r.text[:300]}')
    except Exception as e:
        print(f'ERROR: {e}')
