"""Playwright: блокируем внешние скрипты, перехватываем emex.ru API запросы."""
import sys, io, json, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

from playwright.sync_api import sync_playwright, Route, Request

article = 'CAF100493C'
brand   = 'CHAMPION'
url     = f'https://emex.ru/products/{article}/{brand}'

# Домены которые блокируем
BLOCKED_DOMAINS = [
    'yandex.ru', 'yastatic.net', 'yandex.net',
    'google', 'doubleclick', 'googlesyndication',
    'facebook', 'mc.yandex', 'counter.yadro',
    'sentry', 'sentry.io',
    'posthog', 'segment',
    'retailMedia', 'retail',
]

emex_api_calls = []

print(f'Загрузка: {url}')

with sync_playwright() as p:
    browser = p.chromium.launch(
        headless=True,
        args=['--disable-dev-shm-usage', '--no-sandbox', '--disable-gpu',
              '--disable-web-security', '--blink-settings=imagesEnabled=false']
    )
    context = browser.new_context(
        user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
        locale='ru-RU',
        ignore_https_errors=True,
    )
    page = context.new_page()

    # Блокируем внешние домены
    def route_handler(route: Route, request: Request):
        url_req = request.url
        should_block = any(d in url_req for d in BLOCKED_DOMAINS)
        if should_block:
            route.abort()
        else:
            # Логируем запросы к emex.ru
            if 'emex.ru' in url_req:
                rel = url_req.replace('https://emex.ru', '')
                if not any(x in rel for x in ['_next/static', '.woff', '.css', 'favicon']):
                    emex_api_calls.append({'method': request.method, 'url': url_req})
            route.continue_()

    page.route('**/*', route_handler)

    try:
        page.goto(url, wait_until='domcontentloaded', timeout=30000)
        print('  domcontentloaded OK')
    except Exception as e:
        print(f'  goto: {e}')

    # Ждём загрузки данных
    page.wait_for_timeout(10000)
    print('  10s wait done')

    print(f'\nEmex.ru запросы ({len(emex_api_calls)}):')
    for c in emex_api_calls:
        u = c['url']
        rel = u.replace('https://emex.ru', '')
        if rel and not rel.startswith('/_next/static'):
            print(f"  {c['method']:6}  {rel[:120]}")

    # Перехватим response bodies для интересных URL
    print('\nПытаемся получить state из страницы...')
    try:
        state_text = page.evaluate('''() => {
            const scripts = document.querySelectorAll("script");
            for (let s of scripts) {
                const t = s.textContent || "";
                if (t.includes("applePay") && t.length > 10000) {
                    return t.substring(0, 500);
                }
            }
            return "not found";
        }''')
        print(f'  state found: {state_text[:200]}')
    except Exception as e:
        print(f'  state eval error: {e}')

    # Попробуем ждать появления конкретного элемента (список аналогов)
    try:
        page.wait_for_selector('[data-testid="makes-list"], .makes-list, [class*="makes"]',
                                timeout=5000)
        print('  Makes list element found!')
    except:
        print('  Makes list element not found (timeout)')

    # Сохраним финальный HTML
    content = page.content()
    with open('data/analytics/emex_debug/playwright_final.html', 'w', encoding='utf-8') as f:
        f.write(content)
    print(f'  Final HTML: {len(content):,} chars saved')

    browser.close()

print('\nГотово!')
