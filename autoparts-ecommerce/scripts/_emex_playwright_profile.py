"""
Playwright с реальным Chrome-профилем для доступа к сессии emex.ru.
"""
import sys, io, json, asyncio, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

from playwright.async_api import async_playwright

CHROME_PROFILE = os.path.join(os.environ['USERPROFILE'],
    'AppData', 'Local', 'Google', 'Chrome', 'User Data')
ARTICLE = 'AMDFA475'
BRAND   = 'AMD'
URL     = f'https://emex.ru/products/{ARTICLE}/{BRAND}'

async def main():
    async with async_playwright() as p:
        # Запускаем с реальным профилем Chrome
        ctx = await p.chromium.launch_persistent_context(
            user_data_dir=CHROME_PROFILE,
            headless=True,
            channel='chrome',   # реальный Chrome, а не Chromium
            args=['--disable-blink-features=AutomationControlled'],
        )
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()

        api_calls = []

        async def on_response(resp):
            if resp.status == 200 and 'emex.ru' in resp.url:
                if '_next' not in resp.url and 'static' not in resp.url and 'font' not in resp.url:
                    try:
                        body = await resp.json()
                        api_calls.append((resp.url[:100], body))
                        print(f'  API {resp.url[:80]}')
                    except Exception:
                        pass

        page.on('response', on_response)

        await page.route('**/*.{png,jpg,gif,ico,woff,woff2,ttf,svg}', lambda r: r.abort())
        await page.route('**/mc.yandex*', lambda r: r.abort())

        print(f'Открываю {URL}')
        try:
            await page.goto(URL, wait_until='commit', timeout=60000)
        except Exception as e:
            print(f'goto: {e}')

        print('Ждём 8 сек...')
        await page.wait_for_timeout(8000)

        state = await page.evaluate("""() => {
            try {
                const s = window.__NEXT_REDUX_STORE__.getState();
                const d = s.details;
                return {
                    originals: d.originals || [],
                    analogs:   (d.analogs || []).slice(0, 5),
                    analogs_total: (d.analogs || []).length,
                    makes: (d.makes && d.makes.list) ? d.makes.list.map(m => m.make) : [],
                    name:  d.name,
                    isAuthenticated: s.user && s.user.isAuthenticated,
                    userType: s.user && s.user.type,
                };
            } catch(e) { return {error: e.toString()}; }
        }""")

        print()
        print('=== Redux State ===')
        print(json.dumps(state, ensure_ascii=False, indent=2))

        await ctx.close()

asyncio.run(main())
