"""Playwright: получаем Redux state после полной загрузки страницы."""
import sys, io, json, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

from playwright.sync_api import sync_playwright

article = 'CAF100493C'
brand   = 'CHAMPION'
url     = f'https://emex.ru/products/{article}/{brand}'

BLOCKED = ['yandex.ru/retail', 'yastatic.net', 'sentry.io', 'sentrymxm',
           'mc.yandex', 'counter.yadro', 'googletagmanager']

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True, args=['--no-sandbox', '--disable-gpu'])
    context = browser.new_context(
        user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        locale='ru-RU',
        ignore_https_errors=True,
    )
    page = context.new_page()

    def route_handler(route, request):
        if any(b in request.url for b in BLOCKED):
            route.abort()
        else:
            route.continue_()

    page.route('**/*', route_handler)

    try:
        page.goto(url, wait_until='domcontentloaded', timeout=30000)
    except Exception as e:
        print(f'goto: {e}')

    # Ждём React hydration и data loading
    page.wait_for_timeout(15000)

    # Получаем Redux state
    try:
        state = page.evaluate('''() => {
            // Ищем Redux store через React Fiber
            const el = document.querySelector("[data-reactroot], #__next > *");
            if (!el) return {error: "no react root"};

            // Ищем через window keys
            const storeKey = Object.keys(window).find(k =>
                window[k] && typeof window[k].getState === "function" &&
                typeof window[k].dispatch === "function"
            );
            if (storeKey) {
                const state = window[storeKey].getState();
                return {
                    found: storeKey,
                    details_originals: state?.details?.originals,
                    details_analogs: state?.details?.analogs,
                    details_makes_count: state?.details?.makes?.list?.length,
                    details_isLoaded: state?.details?.isLoaded,
                    details_name: state?.details?.name,
                    details_num: state?.details?.num,
                    user_type: state?.user?.userType,
                    parts_vehicles: state?.parts?.vehicles?.slice?.(0, 3),
                };
            }
            return {error: "no store found", windowKeys: Object.keys(window).filter(k => k.length > 3 && k.length < 30)};
        }''')
        print(json.dumps(state, ensure_ascii=False, indent=2))
    except Exception as e:
        print(f'state eval error: {e}')

    # Также попробуем React DevTools подход
    try:
        fiber_state = page.evaluate('''() => {
            // Traverse React fiber tree
            function findStore(fiber) {
                if (!fiber) return null;
                if (fiber.memoizedState) {
                    // Check if it's a Redux store subscriber
                    let ms = fiber.memoizedState;
                    while (ms) {
                        if (ms.queue && ms.queue.dispatch) {
                            const store = ms.memoizedState;
                            if (store && store.details) return {
                                originals: store.details.originals,
                                analogs: store.details.analogs,
                                makes: store.details.makes?.header,
                                num: store.details.num,
                            };
                        }
                        ms = ms.next;
                    }
                }
                const result = findStore(fiber.child) || findStore(fiber.sibling);
                return result;
            }
            const root = document.getElementById("__next");
            const fiberKey = Object.keys(root || {}).find(k => k.startsWith("__reactFiber") || k.startsWith("__reactContainer"));
            if (!fiberKey || !root) return {error: "no fiber"};
            return findStore(root[fiberKey]);
        }''')
        print('\nFiber state:', json.dumps(fiber_state, ensure_ascii=False, indent=2)[:1000])
    except Exception as e:
        print(f'fiber error: {e}')

    browser.close()
