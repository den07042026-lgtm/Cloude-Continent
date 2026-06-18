"""Playwright: ждём дольше, скроллим страницу для триггера lazy loading."""
import sys, io, json, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

from playwright.sync_api import sync_playwright

article = 'CAF100493C'
brand   = 'CHAMPION'
url     = f'https://emex.ru/products/{article}/{brand}'

BLOCKED = ['yandex.ru/retail', 'yastatic.net', 'sentry.io', 'sentrymxm',
           'google-analytics', 'mc.yandex', 'counter.yadro',
           'googletagmanager', 'googlesyndication', 'facebook']

emex_calls = []

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True, args=['--no-sandbox', '--disable-gpu'])
    context = browser.new_context(
        user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        locale='ru-RU',
        ignore_https_errors=True,
    )
    page = context.new_page()

    def route_handler(route, request):
        u = request.url
        if any(b in u for b in BLOCKED):
            route.abort()
        else:
            if 'emex.ru' in u:
                rel = u.replace('https://emex.ru', '').replace('https://api1.emex.ru', '[api1]').replace('https://apm.emex.ru', '[apm]')
                emex_calls.append({'method': request.method, 'url': u, 'rel': rel})
            route.continue_()

    page.route('**/*', route_handler)

    try:
        page.goto(url, wait_until='domcontentloaded', timeout=30000)
        print('domcontentloaded OK')
    except Exception as e:
        print(f'goto: {e}')

    # Скроллим страницу для инициации lazy loading
    for scroll_pos in [300, 600, 1000, 1500, 2000]:
        page.evaluate(f'window.scrollTo(0, {scroll_pos})')
        page.wait_for_timeout(500)

    # Ждём ещё
    page.wait_for_timeout(15000)

    print(f'\nВсе emex.ru запросы ({len(emex_calls)}):')
    seen = set()
    for c in emex_calls:
        base_url = re.sub(r'\?.*', '', c['url'])
        if base_url not in seen:
            seen.add(base_url)
            print(f"  {c['method']:4}  {c['rel'][:120]}")

    # Перехватываем response для api1
    print('\nПробуем вручную вызвать api1.emex.ru endpoints...')
    try:
        # Делаем fetch запрос изнутри страницы (с cookies браузера)
        result = page.evaluate('''async () => {
            const endpoints = [
                '/detail/info?detailNum=CAF100493C&make=CHAMPION',
                '/detail/originals?detailNum=CAF100493C&make=CHAMPION',
                '/detail/analogs?detailNum=CAF100493C&make=CHAMPION',
                '/search?detailNum=CAF100493C&make=CHAMPION',
                '/catalog?detailNum=CAF100493C&make=CHAMPION',
            ];
            const results = {};
            for (const ep of endpoints) {
                try {
                    const r = await fetch('https://api1.emex.ru' + ep, {
                        headers: {
                            'Accept': 'application/json',
                            'Referer': 'https://emex.ru/products/CAF100493C/CHAMPION',
                        }
                    });
                    results[ep] = {status: r.status, size: (await r.text()).length};
                } catch(e) {
                    results[ep] = {error: e.message};
                }
            }
            return results;
        }''')
        print('Manual fetch results:')
        for ep, res in result.items():
            print(f'  {ep}: {res}')
    except Exception as e:
        print(f'fetch eval error: {e}')

    browser.close()
