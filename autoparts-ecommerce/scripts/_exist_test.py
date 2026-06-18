"""Тестирует exist.ru как источник OEM-номеров."""
import requests, sys, io, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
from bs4 import BeautifulSoup

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,*/*;q=0.9',
    'Accept-Language': 'ru-RU,ru;q=0.9',
}
session = requests.Session()
session.headers.update(HEADERS)

article = 'CAF100493C'
brand   = 'CHAMPION'

# Тест различных URL exist.ru
urls = [
    f'https://www.exist.ru/Price/?pr={article}',
    f'https://www.exist.ru/Price/?pr={article}&brand={brand}',
    f'https://www.exist.ru/search/?query={article}',
    f'https://exist.ru/Price/SearchByOem/?oem={article}',
]

for url in urls:
    print(f'\n--- {url}')
    try:
        r = session.get(url, timeout=15, allow_redirects=True)
        print(f'    {r.status_code}  {len(r.text):,}  final_url={r.url}')
        if r.status_code == 200 and len(r.text) > 5000:
            soup = BeautifulSoup(r.text, 'lxml')
            # Ищем OEM числа
            text = soup.get_text()
            # Показываем title
            title = soup.find('title')
            print(f'    title: {title.text if title else "N/A"}')
            # Ищем OEM паттерны
            oem_matches = re.findall(r'\b[A-Z0-9]{5,20}\b', text)
            oem_candidates = [m for m in set(oem_matches) if len(m) >= 8]
            print(f'    OEM candidates ({len(oem_candidates)}): {oem_candidates[:10]}')
            # Сохраним
            with open(f'data/analytics/emex_debug/exist_{article}_{r.status_code}.html', 'w', encoding='utf-8') as f:
                f.write(r.text)
    except Exception as e:
        print(f'    ERROR: {e}')

# Тест AMD фильтра (первый в списке)
print('\n=== Тест AMD AMDFA576 ===')
article2 = 'AMDFA576'
brand2   = 'AMD'
url2 = f'https://www.exist.ru/Price/?pr={article2}'
try:
    r = session.get(url2, timeout=15)
    print(f'{r.status_code}  {len(r.text):,}  {r.url}')
    if r.status_code == 200:
        soup = BeautifulSoup(r.text, 'lxml')
        title = soup.find('title')
        print(f'title: {title.text if title else "N/A"}')
        # Ищем таблицы с данными
        tables = soup.find_all('table')
        print(f'tables: {len(tables)}')
        # Первые 2000 символов текста
        text = ' '.join(soup.get_text().split()[:200])
        print(f'text: {text}')
except Exception as e:
    print(f'ERROR: {e}')
