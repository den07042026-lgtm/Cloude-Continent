"""Ищет login endpoint emex.ru в JS бандле."""
import sys, io, re, requests
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import warnings; warnings.filterwarnings('ignore')

s = requests.Session()
s.headers['User-Agent'] = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36'

# Качаем _app.js где нашли originals/analogs
url = 'https://emex.ru/_next/static/chunks/pages/_app-7d8986f63512d348.js'
r = s.get(url, timeout=20)
content = r.text
print(f'Размер _app.js: {len(content):,} байт')

# Ищем login/auth endpoints
login_patterns = [
    r'/api/(?:auth|login|signin|user)[^"\']{0,50}',
    r'\/auth\/[^\'"]{3,50}',
    r'"\/user\/[^\'"]{3,50}"',
    r'api[/.](?:auth|login|signin)',
    r'emexSessionId',
    r'token["\']?\s*[:=]',
    r'authToken',
    r'Bearer',
    r'Authorization',
]

for pat in login_patterns:
    matches = re.findall(pat, content)
    unique = list(set(matches))[:5]
    if unique:
        print(f'\n[{pat[:40]}]')
        for m in unique:
            print(f'  {m[:100]}')

# Ищем куки названия
cookie_patterns = re.findall(r'cookies?\.["\']([^"\']{3,40})["\']', content)
print(f'\nИмена куки в коде: {list(set(cookie_patterns))[:20]}')

# Ищем POST endpoints
post_endpoints = re.findall(r'\.post\(["\']([^"\']{5,80})["\']', content)
print(f'\nPOST endpoints: {list(set(post_endpoints))[:20]}')

# Ищем auth-related функции
auth_funcs = re.findall(r'function\s+\w*(?:[Ll]og[Ii]n|[Aa]uth|[Ss]ign)[^{]{0,100}', content)
print(f'\nAuth functions: {auth_funcs[:5]}')
