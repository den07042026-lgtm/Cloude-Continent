"""
Перехватывает все сетевые запросы при открытии страницы emex.ru
чтобы найти AJAX-эндпоинты для originals и analogs.
"""
import sys, io, json, asyncio
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

from playwright.async_api import async_playwright

ARTICLE = 'AMDFA475'
BRAND   = 'AMD'
URL     = f'https://emex.ru/products/{ARTICLE}/{BRAND}'

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        ctx = await browser.new_context(
            locale='ru-RU',
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36',
        )
        page = await ctx.new_page()

        captured = []

        async def on_request(req):
            url = req.url
            if 'emex' in url or 'api1' in url:
                captured.append(('REQ', req.method, url[:120]))

        async def on_response(resp):
            url = resp.url
            if ('api1.emex' in url or 'emex.ru/api' in url) and resp.status == 200:
                try:
                    body = await resp.json()
                    keys = list(body.keys())[:8] if isinstance(body, dict) else f'list[{len(body)}]'
                    captured.append(('RESP_JSON', resp.status, url[:100], str(keys)))
                except Exception:
                    ct = resp.headers.get('content-type', '')
                    if 'json' in ct:
                        captured.append(('RESP_JSON?', resp.status, url[:100]))

        page.on('request', on_request)
        page.on('response', on_response)

        # Блокируем тяжёлые ресурсы для ускорения
        await page.route('**/*.{png,jpg,jpeg,gif,svg,ico,woff,woff2,ttf}', lambda r: r.abort())
        await page.route('**/yandex*', lambda r: r.abort())
        await page.route('**/google*analytics*', lambda r: r.abort())
        await page.route('**/sentry*', lambda r: r.abort())

        print(f'Открываю {URL}')
        await page.goto(URL, wait_until='commit', timeout=60000)

        # Ждём появления originals/analogs в Redux store
        print('Жду загрузки originals/analogs...')
        for i in range(20):
            await page.wait_for_timeout(1000)
            state = await page.evaluate("""() => {
                try {
                    const s = window.__NEXT_REDUX_STORE__.getState();
                    const d = s.details;
                    return {
                        originals_len: (d.originals || []).length,
                        analogs_len:   (d.analogs   || []).length,
                        isLoaded:       d.isLoaded,
                    };
                } catch(e) { return {error: e.toString()}; }
            }""")
            print(f'  [{i+1}s] {state}')
            orig = state.get('originals_len', -1)
            anal = state.get('analogs_len', -1)
            if orig > 0 or anal > 0:
                print('  -> Данные появились!')
                break
            if i >= 12:
                print('  -> Таймаут, данные не загрузились')
                break

        # Финальный дамп Redux state
        final = await page.evaluate("""() => {
            try {
                const s = window.__NEXT_REDUX_STORE__.getState();
                const d = s.details;
                return {
                    originals: d.originals || [],
                    analogs:   d.analogs   || [],
                    makes:     (d.makes && d.makes.list) ? d.makes.list : [],
                    name:      d.name,
                    num:       d.num,
                    make:      d.make,
                };
            } catch(e) { return {error: e.toString()}; }
        }""")
        print()
        print('=== FINAL REDUX STATE ===')
        print(json.dumps(final, ensure_ascii=False, indent=2)[:3000])

        print()
        print('=== ПЕРЕХВАЧЕННЫЕ ЗАПРОСЫ (api1/emex api) ===')
        for item in captured:
            print(item)

        await browser.close()

asyncio.run(main())
