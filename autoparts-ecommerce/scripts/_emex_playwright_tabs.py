"""
Playwright: открывает /f URL и пробует кликать по вкладкам аналогов/оригиналов,
перехватывает все JSON-ответы.
"""
import sys, io, json, asyncio
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

from playwright.async_api import async_playwright

ARTICLE = 'AMDFA475'
URL     = f'https://emex.ru/f?detailNum={ARTICLE}&packet=-1'

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        ctx = await browser.new_context(
            locale='ru-RU',
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36',
        )
        page = await ctx.new_page()

        api_responses = []

        async def on_response(resp):
            url = resp.url
            if resp.status == 200 and 'emex.ru' in url and '_next' not in url and 'static' not in url:
                try:
                    body = await resp.json()
                    api_responses.append((url, body))
                    print(f'  API: {url[:100]}')
                    if isinstance(body, dict):
                        print(f'       keys: {list(body.keys())[:8]}')
                except Exception:
                    pass

        page.on('response', on_response)

        # Блокируем лишнее
        await page.route('**/*.{png,jpg,gif,ico,woff,woff2,ttf}', lambda r: r.abort())
        await page.route('**/yandex*', lambda r: r.abort())
        await page.route('**/sentry*', lambda r: r.abort())
        await page.route('**/posthog*', lambda r: r.abort())

        print(f'Открываю {URL}')
        try:
            await page.goto(URL, wait_until='commit', timeout=60000)
        except Exception as e:
            print(f'goto error: {e}')

        print('Ждём загрузки страницы...')
        await page.wait_for_timeout(5000)

        # Проверяем Redux state
        state = await page.evaluate("""() => {
            try {
                const s = window.__NEXT_REDUX_STORE__.getState();
                const d = s.details;
                return {
                    originals: d.originals || [],
                    analogs:   d.analogs   || [],
                    makes_count: (d.makes && d.makes.list) ? d.makes.list.length : 0,
                    name: d.name,
                    isLoaded: d.isLoaded,
                };
            } catch(e) {
                return {error: e.toString()};
            }
        }""")
        print(f'Redux state: {json.dumps(state, ensure_ascii=False)[:300]}')

        # Ищем кнопки/вкладки с аналогами
        print()
        print('Ищем элементы страницы...')
        tabs = await page.query_selector_all('button, a[role="tab"], .tab, [class*="analog"], [class*="original"]')
        for t in tabs[:20]:
            txt = (await t.inner_text()).strip()[:50]
            cls = await t.get_attribute('class') or ''
            if txt:
                print(f'  element: {txt!r}  class={cls[:40]!r}')

        # Ищем специфичные кнопки аналогов/оригиналов
        selectors = [
            'text=Аналоги', 'text=Оригиналы', 'text=Оригинальные',
            '[class*="analog"]', '[class*="original"]', '[class*="cross"]',
            'button:has-text("Аналог")', 'a:has-text("Аналог")',
        ]
        for sel in selectors:
            try:
                el = await page.query_selector(sel)
                if el:
                    txt = (await el.inner_text()).strip()[:60]
                    print(f'  FOUND {sel!r}: {txt!r}')
                    await el.click()
                    await page.wait_for_timeout(3000)
                    state2 = await page.evaluate("""() => {
                        try {
                            const s = window.__NEXT_REDUX_STORE__.getState();
                            const d = s.details;
                            return {orig: (d.originals||[]).length, anal: (d.analogs||[]).length};
                        } catch(e) { return {error: e.toString()}; }
                    }""")
                    print(f'    После клика: {state2}')
            except Exception as e:
                pass

        # Финальный Redux state
        print()
        final = await page.evaluate("""() => {
            try {
                const s = window.__NEXT_REDUX_STORE__.getState();
                const d = s.details;
                return {
                    originals: d.originals || [],
                    analogs:   (d.analogs || []).slice(0, 5),
                    analogs_total: (d.analogs || []).length,
                };
            } catch(e) { return {error: e.toString()}; }
        }""")
        print('=== FINAL ===')
        print(json.dumps(final, ensure_ascii=False, indent=2)[:2000])

        # Все перехваченные API-ответы
        print()
        print(f'=== Всего API-ответов: {len(api_responses)} ===')
        for url, body in api_responses:
            print(f'  {url[:100]}')

        await browser.close()

asyncio.run(main())
