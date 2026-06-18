"""
Перехватывает Redux actions и ищет эндпоинты для originals/analogs.
"""
import sys, io, json, asyncio
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

from playwright.async_api import async_playwright

ARTICLE = 'AMDFA475'
BRAND   = 'AMD'
URL     = f'https://emex.ru/f?detailNum={ARTICLE}&packet=-1'

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        ctx = await browser.new_context(locale='ru-RU',
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36')
        page = await ctx.new_page()

        all_api = []

        async def on_response(resp):
            url = resp.url
            if resp.status == 200 and 'emex.ru' in url:
                if not any(x in url for x in ['_next', 'static', 'font', 'sentry', 'posthog', 'apm']):
                    try:
                        body = await resp.json()
                        all_api.append((url[:120], body))
                        print(f'  API: {url[:100]}')
                    except Exception:
                        pass

        page.on('response', on_response)
        await page.route('**/*.{png,jpg,gif,ico,woff,woff2,ttf,svg}', lambda r: r.abort())
        await page.route('**/mc.yandex*', lambda r: r.abort())

        print(f'Открываю {URL}')
        try:
            await page.goto(URL, wait_until='commit', timeout=60000)
        except Exception as e:
            print(f'goto warn: {e}')

        # Ждём пока Redux store появится
        print('Ждём Redux store...')
        for i in range(15):
            await page.wait_for_timeout(1000)
            ready = await page.evaluate("""() => typeof window.__NEXT_REDUX_STORE__ !== 'undefined'""")
            if ready:
                print(f'  Redux store доступен через {i+1}с')
                break
        else:
            print('  Redux store не появился!')

        # Устанавливаем перехватчик dispatch
        await page.evaluate("""() => {
            const store = window.__NEXT_REDUX_STORE__;
            if (!store) return;
            const orig = store.dispatch.bind(store);
            store.dispatch = function(action) {
                if (typeof action === 'object' && action && action.type) {
                    if (!window._dispatchLog) window._dispatchLog = [];
                    window._dispatchLog.push({type: action.type, keys: Object.keys(action)});
                }
                return orig(action);
            };
            window._dispatchLog = [];
            console.log('Dispatch interceptor installed');
        }""")

        print('Ждём ещё 10 секунд для загрузки данных...')
        await page.wait_for_timeout(10000)

        # Получаем залогированные actions
        actions = await page.evaluate("() => window._dispatchLog || []")
        print(f'\n=== Redux actions ({len(actions)}) ===')
        for a in actions[:50]:
            print(f'  {a}')

        # Состояние
        state = await page.evaluate("""() => {
            try {
                const s = window.__NEXT_REDUX_STORE__.getState();
                const d = s.details;
                const u = s.user || {};
                return {
                    originals_len:      (d.originals || []).length,
                    analogs_len:        (d.analogs || []).length,
                    isAllowFetchDetails: d.isAllowFetchDetails,
                    userType:           u.type,
                    isAuthenticated:    u.isAuthenticated,
                    stateKeys:          Object.keys(s),
                };
            } catch(e) { return {error: e.toString()}; }
        }""")
        print(f'\n=== State ===')
        print(json.dumps(state, ensure_ascii=False, indent=2))

        # Все API calls
        print(f'\n=== API calls ===')
        for url, body in all_api:
            keys = list(body.keys())[:6] if isinstance(body, dict) else type(body).__name__
            print(f'  {url}  keys={keys}')

        await browser.close()

asyncio.run(main())
