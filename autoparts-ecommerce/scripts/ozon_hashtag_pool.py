"""
ozon_hashtag_pool.py
════════════════════════════════════════════════════════════════════════════

Собирает пул реально существующих хэштегов с публичных страниц Озон.

Что делает:
  1. По каждому поисковому запросу (типы автозапчастей) открывает Ozon
  2. Собирает ссылки на карточки товаров из выдачи
  3. Заходит в каждую карточку и вытаскивает все хэштеги
  4. Строит рейтинг по частоте использования
  5. Сохраняет результат в data/ozon_hashtag_pool.json

Результат используется в ozon_hashtags.py как разрешённый пул.

━━━ Запуск ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Обычный запуск (видимый браузер, ~15–25 мин):
  cd C:\\Users\\Admin\\Documents\\Autoparts_Ecommerce
  uv run --with playwright scripts/ozon_hashtag_pool.py

Быстрый тест (только 2 запроса, 5 карточек каждый):
  uv run --with playwright scripts/ozon_hashtag_pool.py --fast

Скрытый браузер (если сайт не блокирует):
  uv run --with playwright scripts/ozon_hashtag_pool.py --headless

Показать топ-100 хэштегов из уже собранного пула:
  uv run --with playwright scripts/ozon_hashtag_pool.py --show
"""

import sys
import re
import json
import time
import random
import logging
import argparse
from pathlib import Path
from collections import Counter
from urllib.parse import quote as urlquote, unquote

sys.stdout.reconfigure(encoding="utf-8")

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
except ImportError:
    sys.exit("Установи playwright: uv run --with playwright python -m playwright install chromium")

# ── Логирование ───────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)-7s %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger(__name__)

# ── Конфигурация ──────────────────────────────────────────────────────────────
DATA_DIR   = Path(__file__).parent.parent / "data"
POOL_FILE  = DATA_DIR / "ozon_hashtag_pool.json"
URLS_CACHE = DATA_DIR / "ozon_product_urls.json"  # чтобы не перебирать выдачу повторно

OZON_BASE  = "https://www.ozon.ru"

# Поисковые запросы — все типы товаров из нашего магазина автозапчастей
SEARCH_QUERIES = [
    "колодки тормозные дисковые",
    "фильтр масляный",
    "амортизатор передний",
    "амортизатор задний",
    "фильтр воздушный",
    "фильтр топливный",
    "фильтр салона",
    "диск тормозной",
    "свечи зажигания",
    "подшипник ступичный",
    "ремень ГРМ",
    "ремень приводной",
    "помпа водяная",
    "термостат",
    "датчик кислородный",
    "датчик температуры",
    "сайлентблок",
    "шаровая опора",
    "рулевая тяга",
    "стойка стабилизатора",
    "ШРУС",
    "пыльник ШРУСа",
    "катализатор",
    "прокладка головки",
    "цепь ГРМ",
    "натяжитель цепи",
]

PRODUCTS_PER_QUERY = 12   # карточек с каждого поискового запроса
MIN_DELAY = 3.5           # минимальная пауза между страницами (сек)
MAX_DELAY = 7.0           # максимальная пауза
MIN_TAG_FREQ = 2          # минимальная частота для включения в пул


# ═══════════════════════════ СБОР URL ТОВАРОВ ════════════════════════════════

def get_product_urls(page, query: str, count: int) -> list[str]:
    """
    Открывает страницу поиска Озон и возвращает список URL карточек.
    """
    search_url = f"{OZON_BASE}/search/?text={urlquote(query)}&from_global=true"
    log.info("  Поиск: '%s'", query)
    try:
        page.goto(search_url, timeout=35_000, wait_until="domcontentloaded")
        time.sleep(random.uniform(2.5, 4.0))
    except Exception as e:
        log.warning("  Ошибка поиска '%s': %s", query, e)
        return []

    # Проверка на капчу
    if _is_captcha(page):
        log.warning("  Капча! Ждём 30с...")
        time.sleep(30)
        return []

    try:
        urls: list[str] = page.evaluate(f"""() => {{
            const seen = new Set();
            const result = [];
            // Карточки товаров содержат /product/ в URL
            const anchors = Array.from(document.querySelectorAll('a[href*="/product/"]'));
            for (const a of anchors) {{
                const href = a.href.split('?')[0];  // без параметров
                if (href.includes('/product/') && !seen.has(href)) {{
                    seen.add(href);
                    result.push(href);
                    if (result.length >= {count * 2}) break;
                }}
            }}
            return result;
        }}""")
        # Фильтруем — только страницы товаров (не категорий)
        product_urls = [u for u in urls if re.search(r'/product/[^/]+-\d+', u)]
        log.info("  Найдено карточек: %d", len(product_urls))
        return product_urls[:count]
    except Exception as e:
        log.warning("  Ошибка парсинга URL: %s", e)
        return []


# ═══════════════════════════ СБОР ХЭШТЕГОВ ═══════════════════════════════════

def extract_hashtags(page, url: str) -> list[str]:
    """
    Открывает карточку товара и вытаскивает все хэштеги.
    Ищет: ссылки на поиск по хэштегу, текстовые узлы с #, атрибуты data-*.
    """
    try:
        page.goto(url, timeout=30_000, wait_until="domcontentloaded")
        time.sleep(random.uniform(2.0, 3.5))
    except Exception as e:
        log.debug("  Ошибка загрузки %s: %s", url, e)
        return []

    if _is_captcha(page):
        log.warning("  Капча на товаре, пропуск...")
        time.sleep(20)
        return []

    try:
        tags: list[str] = page.evaluate(r"""() => {
            const found = new Set();

            // Метод 1: ссылки вида /search/?text=%23хэштег
            const hashLinks = Array.from(document.querySelectorAll(
                'a[href*="text=%23"], a[href*="text=#"]'
            ));
            for (const a of hashLinks) {
                // Извлекаем из href
                const m = a.href.match(/text=%23([^&]+)/i);
                if (m) {
                    try {
                        found.add('#' + decodeURIComponent(m[1]).replace(/\+/g, '_'));
                    } catch(e) {}
                }
                // Или берём innerText если начинается с #
                const txt = a.innerText.trim();
                if (txt.startsWith('#') && txt.length <= 32 && !/\s/.test(txt)) {
                    found.add(txt);
                }
            }

            // Метод 2: текстовые узлы с # (хэштеги могут быть не в ссылках)
            const walker = document.createTreeWalker(
                document.body, NodeFilter.SHOW_TEXT, null, false
            );
            let node;
            while ((node = walker.nextNode())) {
                const chunks = node.textContent.split(/\\s+/);
                for (const ch of chunks) {
                    if (ch.startsWith('#') && ch.length >= 2 && ch.length <= 32
                        && /^#[\\wА-ЯЁа-яё]+$/u.test(ch)) {
                        found.add(ch);
                    }
                }
            }

            // Метод 3: элементы с классами содержащими "tag" или "hash"
            const tagEls = Array.from(document.querySelectorAll(
                '[class*="tag"], [class*="Tag"], [class*="hash"], [class*="Hash"]'
            ));
            for (const el of tagEls) {
                const txt = el.innerText.trim();
                if (txt.startsWith('#') && txt.length <= 32 && !/\s/.test(txt)) {
                    found.add(txt);
                }
            }

            return Array.from(found);
        }""")

        # Финальная валидация
        valid = []
        for tag in tags:
            tag = tag.strip()
            if not tag.startswith("#"):
                continue
            # Только буквы, цифры, подчёркивание
            body = re.sub(r'[^А-ЯЁа-яёA-Za-z0-9_]', '', tag[1:])
            body = body.strip("_")
            if len(body) < 2:
                continue
            clean = "#" + body
            if len(clean) <= 32:
                valid.append(clean)

        return valid

    except Exception as e:
        log.debug("  JS-ошибка на %s: %s", url, e)
        return []


def _is_captcha(page) -> bool:
    """Проверяет наличие капчи или блокировки."""
    try:
        text = page.title().lower()
        return "captcha" in text or "robot" in text or "доступ" in text
    except Exception:
        return False


# ═══════════════════════════════ ВЫВОД ═══════════════════════════════════════

def show_pool(path: Path):
    """Выводит топ хэштегов из сохранённого пула."""
    if not path.exists():
        log.error("Файл пула не найден: %s", path)
        return
    data = json.loads(path.read_text(encoding="utf-8"))
    tags = data.get("hashtags", [])
    log.info("\n━━ Пул хэштегов Озон (топ-%d из %d) ━━", min(100, len(tags)), len(tags))
    for i, item in enumerate(tags[:100], 1):
        log.info("  %3d. %-35s  встречается: %d раз", i, item["tag"], item["count"])


# ═══════════════════════════════ MAIN ════════════════════════════════════════

def main():
    ap = argparse.ArgumentParser(description="Собирает хэштеги с Озон")
    ap.add_argument("--fast",     action="store_true",
                    help="Быстрый тест: первые 2 запроса, 5 карточек каждый")
    ap.add_argument("--headless", action="store_true",
                    help="Скрытый браузер")
    ap.add_argument("--show",     action="store_true",
                    help="Показать уже собранный пул и выйти")
    args = ap.parse_args()

    if args.show:
        show_pool(POOL_FILE)
        return

    queries = SEARCH_QUERIES[:2] if args.fast else SEARCH_QUERIES
    per_query = 5 if args.fast else PRODUCTS_PER_QUERY

    log.info("━" * 58)
    log.info("Сборщик хэштегов Озон")
    log.info("  Поисковых запросов: %d", len(queries))
    log.info("  Карточек за запрос: %d", per_query)
    log.info("  Ожидаемое время:    ~%d мин",
             len(queries) * per_query * 5 // 60 + 1)
    log.info("━" * 58)

    # Загружаем ранее собранные URL (чтобы не дублировать обход)
    collected_urls: dict = {}
    if URLS_CACHE.exists():
        try:
            collected_urls = json.loads(URLS_CACHE.read_text(encoding="utf-8"))
        except Exception:
            pass

    counter: Counter = Counter()
    visited_urls: set[str] = set()

    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            str(Path.home() / ".ozon_hashtag_scraper"),
            headless=args.headless,
            slow_mo=80,
            viewport={"width": 1440, "height": 900},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/125.0.0.0 Safari/537.36"
            ),
            args=["--disable-blink-features=AutomationControlled"],
            locale="ru-RU",
        )
        page = ctx.new_page()
        page.set_extra_http_headers({
            "Accept-Language": "ru-RU,ru;q=0.9",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        })

        # ── Фаза 1: собираем URL товаров ─────────────────────────────────────
        log.info("ФАЗА 1 / Собираем ссылки на карточки...")
        for q in queries:
            if q in collected_urls:
                log.info("  [кэш] '%s': %d URL", q, len(collected_urls[q]))
                continue
            urls = get_product_urls(page, q, per_query)
            collected_urls[q] = urls
            # Сохраняем промежуточно
            DATA_DIR.mkdir(parents=True, exist_ok=True)
            URLS_CACHE.write_text(
                json.dumps(collected_urls, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            time.sleep(random.uniform(MIN_DELAY, MAX_DELAY))

        # ── Фаза 2: обходим карточки, собираем хэштеги ───────────────────────
        log.info("ФАЗА 2 / Обходим карточки и собираем хэштеги...")

        all_urls: list[str] = []
        for q, urls in collected_urls.items():
            if q in queries:  # только текущие запросы
                all_urls.extend(urls)

        # Убираем дубли
        seen_set: set[str] = set()
        unique_urls = []
        for u in all_urls:
            if u not in seen_set:
                seen_set.add(u)
                unique_urls.append(u)

        log.info("  Всего уникальных карточек: %d", len(unique_urls))

        for idx, url in enumerate(unique_urls, 1):
            if url in visited_urls:
                continue
            visited_urls.add(url)

            tags = extract_hashtags(page, url)
            log.info(
                "  [%d/%d] Найдено хэштегов: %d  %s",
                idx, len(unique_urls),
                len(tags),
                tags[:5] if tags else "(нет)",
            )
            counter.update(tags)

            # Сохраняем промежуточно каждые 10 карточек
            if idx % 10 == 0:
                _save_pool(counter, idx, len(unique_urls))

            time.sleep(random.uniform(MIN_DELAY, MAX_DELAY))

        ctx.close()

    # ── Финальное сохранение ─────────────────────────────────────────────────
    total = _save_pool(counter, len(unique_urls), len(unique_urls))

    log.info("━" * 58)
    log.info("Готово! Уникальных хэштегов: %d", total)
    log.info("Файл пула: %s", POOL_FILE)
    log.info("\nТоп-20 самых частых:")
    for tag, cnt in counter.most_common(20):
        log.info("  %-35s %d", tag, cnt)


def _save_pool(counter: Counter, visited: int, total: int) -> int:
    """Сохраняет пул в JSON. Возвращает количество уникальных тегов."""
    filtered = {tag: cnt for tag, cnt in counter.items() if cnt >= MIN_TAG_FREQ}
    sorted_tags = sorted(filtered.items(), key=lambda x: -x[1])

    data = {
        "source": "ozon.ru — автозапчасти",
        "visited_pages": visited,
        "total_pages": total,
        "unique_hashtags": len(sorted_tags),
        "min_frequency": MIN_TAG_FREQ,
        "hashtags": [{"tag": tag, "count": cnt} for tag, cnt in sorted_tags],
        # Плоский список для удобства использования в промпте
        "tags_flat": [tag for tag, _ in sorted_tags],
    }

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    POOL_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    log.info("  [saved] пул: %d хэштегов (из %d карточек)", len(sorted_tags), visited)
    return len(sorted_tags)


if __name__ == "__main__":
    main()
