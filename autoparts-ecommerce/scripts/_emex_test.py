"""Тестирует доступ к emex.ru и сохраняет HTML для анализа структуры."""
import requests, sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'ru-RU,ru;q=0.9,en;q=0.8',
    'Accept-Encoding': 'gzip, deflate, br',
    'Connection': 'keep-alive',
}

# Тестируем несколько URL-форматов с реальным артикулом из нашего топ-500
# CAF100493C = фильтр воздушный Champion (есть в нашем списке)
tests = [
    'https://emex.ru/products/CAF100493C',
    'https://emex.ru/products/CAF100493C/CHAMPION',
    'https://emex.ru/search?query=CAF100493C',
    'https://emex.ru/api/search?number=CAF100493C',
]

session = requests.Session()
session.headers.update(HEADERS)

OUT_DIR = r'C:\Users\Admin\Documents\Autoparts_Ecommerce\data\analytics\emex_debug'
os.makedirs(OUT_DIR, exist_ok=True)

for url in tests:
    print(f'\n--- GET {url}')
    try:
        r = session.get(url, timeout=15, allow_redirects=True)
        print(f'    Status: {r.status_code}  |  URL после редиректов: {r.url}')
        print(f'    Content-Type: {r.headers.get("Content-Type", "?")[:60]}')
        print(f'    Размер: {len(r.text)} символов')
        # Сохраняем первые 5000 символов и полный HTML
        slug = url.replace('https://emex.ru/', '').replace('/', '_').replace('?', '_').replace('&', '_')[:50]
        fpath = os.path.join(OUT_DIR, f'{slug}.html')
        with open(fpath, 'w', encoding='utf-8') as f:
            f.write(r.text)
        print(f'    Сохранено: {fpath}')
        # Показываем начало
        snippet = r.text[:1500].replace('\n', ' ').replace('  ', ' ')
        print(f'    Начало: {snippet[:500]}')
    except Exception as e:
        print(f'    ОШИБКА: {e}')
