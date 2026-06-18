"""
wb_get_cookie.py
════════════════════════════════════════════════════════════════════════════
Открывает реальный Chrome через прокси, решает WBAAS antibot challenge
на wildberries.ru, сохраняет cookies в data/analytics/wb_cookies.json.

Cookies используются wb_product_indexer.py для обхода 429 на search.wb.ru.

Запуск:
  cd C:\\Users\\Admin\\Documents\\Autoparts_Ecommerce
  uv run --with playwright scripts/wb_get_cookie.py
  uv run --with playwright scripts/wb_get_cookie.py --proxy http://user:pass@host:port
  uv run --with playwright scripts/wb_get_cookie.py --headful   # показать браузер
"""

import sys, json, time, logging, argparse
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR    = Path(__file__).parent.parent
COOKIE_FILE = BASE_DIR / "data" / "analytics" / "wb_cookies.json"
LOG_FILE    = BASE_DIR / "logs" / "wb_get_cookie.log"
LOG_FILE.parent.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)

# Прокси по умолчанию (мобильный Билайн)
DEFAULT_PROXY = "http://os.environ.get("PROXY_URL", "")"


def get_cookies(proxy: str, headful: bool = False) -> dict:
    from playwright.sync_api import sync_playwright

    log.info(f"Запускаем Chrome (headless={not headful})...")

    # Парсим proxy URL: http://user:pass@host:port
    proxy_cfg = None
    if proxy:
        if "@" in proxy:
            creds, hostport = proxy.rsplit("@", 1)
            scheme_creds = creds.split("//", 1)
            creds_clean = scheme_creds[-1]
            user, password = creds_clean.split(":", 1) if ":" in creds_clean else (creds_clean, "")
            proxy_cfg = {
                "server": f"http://{hostport}",
                "username": user,
                "password": password,
            }
        else:
            proxy_cfg = {"server": proxy}
        log.info(f"Прокси: {proxy_cfg['server']}")

    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            headless=not headful,
            args=[
                "--no-sandbox",
                "--disable-blink-features=AutomationControlled",
            ],
        )
        ctx = browser.new_context(
            proxy=proxy_cfg,
            locale="ru-RU",
            timezone_id="Europe/Moscow",
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/148.0.0.0 Safari/537.36"
            ),
            viewport={"width": 1440, "height": 900},
            # Скрываем headless — убираем "HeadlessChrome" из sec-ch-ua
            extra_http_headers={
                "Accept-Language": "ru-RU,ru;q=0.9",
                "sec-ch-ua": '"Google Chrome";v="148", "Chromium";v="148", "Not/A)Brand";v="99"',
                "sec-ch-ua-mobile": "?0",
                "sec-ch-ua-platform": '"Windows"',
            },
        )
        page = ctx.new_page()

        log.info("Открываем https://www.wildberries.ru/ ...")
        page.goto("https://www.wildberries.ru/", timeout=60_000)

        # Ждём пока 498-страница сменится на реальную (challenge solved)
        max_wait = 90  # секунд
        for i in range(max_wait):
            title = page.title()
            if "Почти готово" in title or not title:
                if i % 5 == 0:
                    log.info(f"  [{i}с] Решаем antibot challenge: '{title}'...")
                time.sleep(1)
            else:
                log.info(f"  [{i}с] Challenge пройден! Страница: '{title}'")
                break
        else:
            log.warning("Challenge не решён за 90с — возможно нужен headful режим")

        # Дополнительное ожидание для полной загрузки cookies
        time.sleep(3)

        # Посещаем страницу поиска чтобы получить session cookies для search.wb.ru
        log.info("Открываем страницу поиска WB...")
        try:
            page.goto(
                "https://www.wildberries.ru/catalog/0/search.aspx?search=SP1243",
                timeout=30_000,
                wait_until="domcontentloaded",
            )
            time.sleep(3)
            log.info(f"  Страница поиска: '{page.title()}'")
        except Exception as e:
            log.warning(f"  Страница поиска: {e}")

        # Извлекаем все cookies
        cookies = ctx.cookies("https://www.wildberries.ru")
        cookie_dict = {c["name"]: c["value"] for c in cookies}
        log.info(f"Получено cookies: {list(cookie_dict.keys())}")

        # Перехватываем запросы к search.wb.ru чтобы увидеть headers
        search_headers: dict = {}
        def capture_request(request):
            if "search.wb.ru" in request.url:
                search_headers.update(dict(request.headers))

        page.on("request", capture_request)

        # Выполняем поиск в браузере — он сам отправит нужные headers
        log.info("Тест search.wb.ru через JS в браузере...")
        result = page.evaluate("""async () => {
            try {
                const r = await fetch(
                    'https://search.wb.ru/exactmatch/ru/common/v5/search?query=SP1243&resultset=catalog&limit=3&sort=popular&page=1&dest=-1257786',
                    {headers: {'Accept': 'application/json'}}
                );
                return {status: r.status, ok: r.ok};
            } catch(e) { return {error: String(e)}; }
        }""")
        log.info(f"  search.wb.ru через JS: {result}")

        if search_headers:
            log.info("  Заголовки, отправленные в search.wb.ru:")
            for k, v in search_headers.items():
                if k.lower() not in ("user-agent", "accept-encoding", "accept-language", "accept"):
                    log.info(f"    {k}: {v[:80]}")

        # Дополнительно: все cookies для обоих доменов
        wb_cookies = ctx.cookies("https://www.wb.ru")
        log.info(f"Cookies для wb.ru: {[c['name'] for c in wb_cookies]}")

        # Сохраняем все cookies для обоих доменов
        all_cookies = ctx.cookies(["https://www.wildberries.ru", "https://www.wb.ru", "https://search.wb.ru"])
        cookie_dict = {c["name"]: c["value"] for c in all_cookies}

        browser.close()
        return cookie_dict


def main():
    parser = argparse.ArgumentParser(description="WB Cookie Getter — решает WBAAS antibot")
    parser.add_argument("--proxy",  default=DEFAULT_PROXY,
                        help="Прокси http://user:pass@host:port")
    parser.add_argument("--headful", action="store_true",
                        help="Показать браузер (для отладки)")
    args = parser.parse_args()

    log.info("=" * 60)
    log.info("WB Cookie Getter  |  WBAAS antibot solver")
    log.info("=" * 60)

    cookies = get_cookies(args.proxy, args.headful)

    if cookies:
        COOKIE_FILE.parent.mkdir(parents=True, exist_ok=True)
        with open(COOKIE_FILE, "w", encoding="utf-8") as f:
            json.dump(cookies, f, ensure_ascii=False, indent=2)
        log.info(f"Cookies сохранены: {COOKIE_FILE}")
        log.info(f"  Всего: {len(cookies)} cookies")
    else:
        log.error("Cookies не получены!")
        sys.exit(1)


if __name__ == "__main__":
    main()
