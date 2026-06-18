"""Тест: захватываем API-запросы emex.ru через Playwright."""
import sys, io, json, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

from playwright.sync_api import sync_playwright

article = 'CAF100493C'
brand   = 'CHAMPION'
url     = f'https://emex.ru/products/{article}/{brand}'

captured_requests = []
captured_responses = []

def on_request(request):
    u = request.url
    if any(x in u for x in ['/api/', 'emex', 'details', 'original', 'analog', 'product']):
        captured_requests.append({'method': request.method, 'url': u})

def on_response(response):
    u = response.url
    if any(x in u for x in ['/api/', '/v1/', '/v2/', 'detail', 'original', 'analog', 'applicab']):
        captured_responses.append({'status': response.status, 'url': u})

print(f'Открываю: {url}')
with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    context = browser.new_context(
        user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
        locale='ru-RU',
    )
    page = context.new_page()

    # Перехватываем все запросы
    all_reqs = []
    def capture_all(request):
        all_reqs.append(request.url)
    page.on('request', capture_all)

    page.goto(url, wait_until='networkidle', timeout=30000)

    # Ждём ещё немного для lazy-загрузки
    page.wait_for_timeout(3000)

    print(f'\nВсе запросы ({len(all_reqs)}):')
    for r in all_reqs:
        if not any(x in r for x in ['fonts', 'static', 'css', 'png', 'jpg', 'svg', 'ico', 'woff']):
            print(f'  {r}')

    # Получим state из страницы
    state = page.evaluate('''() => {
        if (window.__REDUX_STORE__) return window.__REDUX_STORE__.getState();
        if (window.__store__) return window.__store__.getState();
        // ищем в React devtools или глобальных переменных
        for (let key of Object.keys(window)) {
            if (key.includes('store') || key.includes('Store')) return key;
        }
        return null;
    }''')
    print(f'\nRedux store found: {state}')

    # Попробуем через fiber (React internal)
    details = page.evaluate('''() => {
        const app = document.getElementById('__next');
        if (!app) return null;
        // пробуем найти React fiber
        const key = Object.keys(app).find(k => k.startsWith('__reactFiber') || k.startsWith('__reactContainer'));
        return key ? 'fiber found' : 'no fiber key';
    }''')
    print(f'React fiber: {details}')

    browser.close()

print(f'\nВсего уникальных API-запросов: {len(set(r for r in all_reqs if "_next" not in r and "fonts" not in r))}')
