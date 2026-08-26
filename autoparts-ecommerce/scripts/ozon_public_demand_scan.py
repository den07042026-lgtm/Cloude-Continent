"""
ozon_public_demand_scan.py
════════════════════════════════════════════════════════════════════════════
Открытый бесплатный источник данных о спросе на Ozon: публичная выдача поиска
(без логина, без платных API). Для каждой товарной категории (из списка
категорий ozon_top500_selector.py) делает поисковый запрос на ozon.ru,
собирает у первых ~15 карточек цену, рейтинг и число отзывов.

Число отзывов на карточке — прокси накопленного объёма продаж по этой
позиции (Ozon начисляет отзывы только подтверждённым покупателям), поэтому
средний/медианный «отзывный» показатель по категории — рабочий сигнал
относительного спроса между категориями, не привязанный к WB/MPStats.

Требует видимый (headless=False) браузер — Ozon блокирует headless Chromium
своей антибот-защитой (страница "Похоже, нет соединения" вместо капчи).

Результат: data/analytics/top500_ozon/ozon_public_demand.json
  {category: {avg_reviews, median_price, sample_size, queried_at}}

Запуск:
  cd C:\\Users\\Admin\\Documents\\Autoparts_Ecommerce
  uv run --with playwright scripts/ozon_public_demand_scan.py
"""

import sys
import re
import json
import time
import random
import logging
from pathlib import Path
from datetime import datetime
from statistics import median

sys.stdout.reconfigure(encoding="utf-8")

from playwright.sync_api import sync_playwright

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)-7s %(message)s")
log = logging.getLogger(__name__)

ROOT = Path(__file__).parent.parent
OUT_FILE = ROOT / "data" / "analytics" / "top500_ozon" / "ozon_public_demand.json"
OUT_FILE.parent.mkdir(parents=True, exist_ok=True)

# category -> представительный поисковый запрос (должны совпадать по смыслу
# с CATEGORY_KW из ozon_top500_selector.py)
CATEGORY_QUERIES = {
    "Фильтры масляные": "фильтр масляный",
    "Фильтры воздушные": "фильтр воздушный",
    "Фильтры салонные": "фильтр салонный",
    "Фильтры топливные": "фильтр топливный",
    "Свечи зажигания": "свечи зажигания",
    "Колодки тормозные": "колодки тормозные",
    "Тормозные диски": "диск тормозной",
    "Подшипники ступицы": "подшипник ступицы",
    "Амортизаторы/стойки": "амортизатор автомобильный",
    "Опоры амортизатора": "опора амортизатора",
    "Стойки стабилизатора": "стойка стабилизатора",
    "Сайлентблоки": "сайлентблок",
    "Шаровые опоры": "шаровая опора",
    "ШРУСы": "шрус",
    "Ремни ГРМ/привода": "ремень грм",
    "Катушки зажигания": "катушка зажигания",
    "Генераторы": "генератор автомобильный",
    "Стартеры": "стартер автомобильный",
    "Насосы/помпы": "помпа водяная",
    "Форсунки": "форсунка топливная",
    "Радиаторы": "радиатор охлаждения",
    "Термостаты": "термостат автомобильный",
    "Сальники": "сальник коленвала",
    "Тросы": "трос стояночного тормоза",
    "Бачки расширительные": "бачок расширительный",
    "Болты ГБЦ": "болт гбц",
    "Прокладки ГБЦ": "прокладка гбц",
    "Кольца поршневые": "кольца поршневые",
    "Зеркала": "зеркало заднего вида",
    "Бамперы": "бампер передний",
    "Дроссельные заслонки": "дроссельная заслонка",
    "Датчики": "датчик кислородный",
    "Фонари/лампы": "фонарь задний",
    "Аккумуляторы": "аккумулятор автомобильный",
    "Рулевые рейки/наконечники": "рулевой наконечник",
    "Рычаги подвески": "рычаг подвески",
    "Ролики натяжные/обводные": "ролик натяжной ремня грм",
    "Диски/корзины сцепления": "диск сцепления",
    "Пружины подвески": "пружина подвески",
    "Ступицы в сборе": "ступица в сборе",
    "Опоры/подушки двигателя": "опора двигателя",
    "Тормозные цилиндры": "рабочий тормозной цилиндр",
    "Поршни": "поршень с пальцем",
    "Дворники/щётки стеклоочистителя": "щетки стеклоочистителя",
    "Клапаны": "клапан двигателя впускной",
    "Выхлопная система": "глушитель автомобильный",
    "Рычаги/тяги подвески": "рычаг подвески передний",
    "Вкладыши двигателя": "вкладыши коленвала",
    "Крестовины карданные": "крестовина карданного вала",
    "Тормозные барабаны": "барабан тормозной",
    "Патрубки/шланги/гофры": "патрубок радиатора",
    "Наконечники рулевые": "наконечник рулевой тяги",
    "Пыльники ШРУС/аморт.": "пыльник шруса",
    "Натяжители цепи/ремня": "натяжитель ремня грм",
    "Втулки стабилизатора": "втулка стабилизатора",
    "Сцепление (комплект)": "комплект сцепления",
    "Вентиляторы охлаждения": "вентилятор охлаждения радиатора",
    "Подшипники (общие)": "подшипник ступицы",
    "Ремкомплекты тормозного суппорта": "ремкомплект тормозного суппорта",
}

CARDS_PER_QUERY = 16
MIN_DELAY = 3.0
MAX_DELAY = 6.0


def _is_blocked(page) -> bool:
    try:
        t = page.title().lower()
        return "нет соединения" in t or "captcha" in t or "доступ ограничен" in t
    except Exception:
        return False


def scan_category(page, category: str, query: str) -> dict | None:
    from urllib.parse import quote as urlquote
    url = f"https://www.ozon.ru/search/?text={urlquote(query)}&from_global=true"
    for attempt in range(3):
        try:
            page.goto(url, timeout=40_000, wait_until="domcontentloaded")
        except Exception as e:
            log.warning("  '%s': ошибка перехода (%s), попытка %d", category, e, attempt + 1)
            time.sleep(5)
            continue
        time.sleep(random.uniform(3.5, 5.5))
        if _is_blocked(page):
            log.warning("  '%s': блок/капча, ждём и повторяем...", category)
            time.sleep(15)
            continue
        break
    else:
        log.error("  '%s': не удалось получить выдачу после 3 попыток", category)
        return None

    try:
        cards = page.evaluate(f"""() => {{
            const anchors = Array.from(document.querySelectorAll('a[href*="/product/"]'));
            const seen = new Set();
            const out = [];
            for (const a of anchors) {{
                const href = a.href.split('?')[0];
                if (seen.has(href)) continue;
                seen.add(href);
                let node = a, text = '';
                for (let i = 0; i < 6 && node; i++) {{
                    node = node.parentElement;
                    if (node) text = node.innerText || '';
                    if (text.length > 40) break;
                }}
                out.push(text.slice(0, 400));
                if (out.length >= {CARDS_PER_QUERY}) break;
            }}
            return out;
        }}""")
    except Exception as e:
        log.warning("  '%s': ошибка парсинга DOM (%s)", category, e)
        return None

    reviews, prices, ratings = [], [], []
    for text in cards:
        m_rev = re.search(r"([\d\s]{1,7})\s*отзыв", text)
        if m_rev:
            n = re.sub(r"\s", "", m_rev.group(1))
            if n.isdigit():
                reviews.append(int(n))
        m_price = re.search(r"([\d\s]{2,8})\s*₽", text)
        if m_price:
            n = re.sub(r"\s", "", m_price.group(1))
            if n.isdigit():
                prices.append(int(n))
        m_rating = re.search(r"\b([0-5]\.\d)\b", text)
        if m_rating:
            ratings.append(float(m_rating.group(1)))

    if not reviews:
        log.warning("  '%s': не найдено ни одной карточки с отзывами (карточек всего %d)", category, len(cards))
        return None

    result = {
        "avg_reviews": sum(reviews) / len(reviews),
        "median_reviews": median(reviews),
        "median_price": median(prices) if prices else 0,
        "avg_rating": sum(ratings) / len(ratings) if ratings else 0,
        "sample_size": len(reviews),
        "cards_found": len(cards),
        "queried_at": datetime.now().isoformat(timespec="seconds"),
    }
    log.info("  '%s' (%s): %d карточек, ср. отзывов=%.0f, медиана цены=%.0f₽",
              category, query, result["sample_size"], result["avg_reviews"], result["median_price"])
    return result


def main():
    results = {}
    if OUT_FILE.exists():
        try:
            results = json.loads(OUT_FILE.read_text(encoding="utf-8"))
        except Exception:
            results = {}

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, args=["--disable-blink-features=AutomationControlled"])
        ctx = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                       "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
            viewport={"width": 1366, "height": 900},
            locale="ru-RU",
        )
        page = ctx.new_page()

        for i, (category, query) in enumerate(CATEGORY_QUERIES.items(), 1):
            if category in results:
                log.info("[%d/%d] %s - уже в кэше, пропуск", i, len(CATEGORY_QUERIES), category)
                continue
            log.info("[%d/%d] %s", i, len(CATEGORY_QUERIES), category)
            res = scan_category(page, category, query)
            if res:
                results[category] = res
                OUT_FILE.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
            time.sleep(random.uniform(MIN_DELAY, MAX_DELAY))

        browser.close()

    log.info("Готово. Категорий с данными: %d/%d -> %s", len(results), len(CATEGORY_QUERIES), OUT_FILE)


if __name__ == "__main__":
    main()
