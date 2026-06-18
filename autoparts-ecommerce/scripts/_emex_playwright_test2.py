"""Playwright: захватываем API-запросы emex.ru, используем load вместо networkidle."""
import sys, io, json, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

from playwright.sync_api import sync_playwright

article = 'CAF100493C'
brand   = 'CHAMPION'
url     = f'https://emex.ru/products/{article}/{brand}'

print(f'Открываю: {url}')
with sync_playwright() as p:
    browser = p.chromium.launch(
        headless=True,
        args=['--disable-dev-shm-usage', '--no-sandbox']
    )
    context = browser.new_context(
        user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
        locale='ru-RU',
    )
    page = context.new_page()

    # Захватываем только не-статичные запросы
    api_calls = []
    def on_req(request):
        u = request.url
        if not any(x in u for x in ['.woff', '.css', '.png', '.svg', '.ico', '_next/static',
                                     'fonts', 'yandex.ru/retail', 'googletagmanager']):
            api_calls.append({'method': request.method, 'url': u})

    page.on('request', on_req)

    try:
        page.goto(url, wait_until='load', timeout=45000)
        print('Page loaded (load event)')
        # Ждём ещё 5 секунд для AJAX
        page.wait_for_timeout(5000)
        print('Extra wait done')
    except Exception as e:
        print(f'goto error: {e}')

    print(f'\nAPI calls ({len(api_calls)}):')
    seen = set()
    for r in api_calls:
        u = r['url']
        # Показываем уникальные, без _next/image и tracking
        base_u = re.sub(r'\?.*', '', u)
        if base_u not in seen and not any(x in u for x in ['_next/image', 'sentry', 'gtm', 'retail', 'metrika']):
            seen.add(base_u)
            print(f"  {r['method']}  {u[:120]}")

    browser.close()
