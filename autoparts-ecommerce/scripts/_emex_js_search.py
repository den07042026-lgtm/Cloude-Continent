"""
Скачивает JS бандлы emex.ru и ищет код загрузки originals/analogs.
"""
import sys, io, re, requests
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import warnings; warnings.filterwarnings('ignore')

s = requests.Session()
s.headers['User-Agent'] = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36'

# Получаем список JS бандлов со страницы
r = s.get('https://emex.ru/products/AMDFA475/AMD', timeout=20)
html = r.text

# Находим все _next/static JS чанки
chunks = re.findall(r'/_next/static/chunks/([\w.-]+\.js)', html)
build_id = re.search(r'/_next/static/([A-Za-z0-9_-]+)/_buildManifest', html)
build_id = build_id.group(1) if build_id else ''
print(f'Build ID: {build_id}')
print(f'Чанков: {len(chunks)}')

# Специфичные страничные чанки (не фреймворк)
page_chunks = list(set(chunks))  # все уникальные чанки
print(f'Страничных чанков: {len(page_chunks)}')

KEYWORDS = ['originals', 'analogs', 'fetchDetails', 'loadDetails', 'isAllowFetch']

for chunk in page_chunks:
    url = f'https://emex.ru/_next/static/chunks/{chunk}'
    try:
        r2 = s.get(url, timeout=10)
        content = r2.text
        found = {kw: kw in content for kw in KEYWORDS}
        if any(found.values()):
            print(f'\n=== {chunk} ({len(content):,} байт) ===')
            for kw, hit in found.items():
                if hit:
                    # Показываем контекст вокруг ключевого слова
                    idx = content.find(kw)
                    while idx >= 0:
                        snippet = content[max(0, idx-60):idx+120]
                        # Только уникальные короткие сниппеты
                        print(f'  [{kw}] ...{snippet.strip()[:150]}...')
                        idx = content.find(kw, idx + len(kw) + 1)
                        if idx > 0 and content.find(kw, idx) - idx > 500:
                            break  # Только первые несколько вхождений
    except Exception as e:
        pass
