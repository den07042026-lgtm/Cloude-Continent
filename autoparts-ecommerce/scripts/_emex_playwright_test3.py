"""Playwright: domcontentloaded + перехват response body для API запросов."""
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
        args=['--disable-dev-shm-usage', '--no-sandbox', '--disable-gpu']
    )
    context = browser.new_context(
        user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
        locale='ru-RU',
        ignore_https_errors=True,
    )
    page = context.new_page()

    # Перехватываем ответы
    interesting_responses = []
    def on_response(response):
        u = response.url
        if any(x in u for x in ['/api/', '/v1/', '/v2/', 'detail', 'original', 'analog', 'applicab']):
            if not any(x in u for x in ['_next', 'static', 'fonts', '.css', '.js']):
                interesting_responses.append({'status': response.status, 'url': u})
                try:
                    body = response.text()
                    interesting_responses[-1]['body'] = body[:2000]
                except:
                    pass

    page.on('response', on_response)

    try:
        page.goto(url, wait_until='domcontentloaded', timeout=20000)
        print('domcontentloaded fired')
        page.wait_for_timeout(8000)
        print('8s wait done')
    except Exception as e:
        print(f'goto: {e}')

    # Пробуем извлечь state из JS
    try:
        state_js = page.evaluate('''() => {
            const scripts = document.querySelectorAll('script:not([src])');
            for (let s of scripts) {
                const t = s.textContent;
                if (t && t.includes('applePay') && t.includes('details')) {
                    return t.substring(0, 100);
                }
            }
            return 'NOT FOUND';
        }''')
        print(f'State script found: {state_js}')
    except Exception as e:
        print(f'evaluate error: {e}')

    # Посмотрим на все сетевые запросы через performance API
    try:
        perf = page.evaluate('''() => {
            return performance.getEntriesByType("resource")
                .filter(r => !r.name.includes("static") && !r.name.includes("font"))
                .map(r => r.name)
                .slice(0, 50);
        }''')
        print(f'\nPerformance resources ({len(perf)}):')
        for u in perf[:30]:
            print(f'  {u}')
    except Exception as e:
        print(f'perf error: {e}')

    print(f'\nInteresting API responses ({len(interesting_responses)}):')
    for r in interesting_responses:
        print(f'  {r.get("status")} {r.get("url")}')
        if r.get('body'):
            print(f'  body: {r["body"][:200]}')

    browser.close()
