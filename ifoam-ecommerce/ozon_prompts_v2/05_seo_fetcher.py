# 05_seo_fetcher.py - Сохранение страниц поисковых запросов Ozon Seller как PDF
# Запуск: сначала открой Chrome командой:
#   chrome.exe --remote-debugging-port=9222 --user-data-dir="C:/Users/1/AppData/Local/Google/Chrome/User Data/Default"
# Затем: C:/Python314/python.exe c:/Users/1/Downloads/ozon_prompts_v2/05_seo_fetcher.py

from playwright.async_api import async_playwright
import asyncio

# ══════════════════════════════════════════════════════
# НАСТРОЙКИ — заполняются из предыдущих промптов
# ══════════════════════════════════════════════════════

# Из Промпта 1 (карточка товара) + Промпта 3 (названия конкурентов)
# Claude подставляет автоматически на основе данных товара
SEARCH_QUERIES = [
    "кондиционер для белья",                    # ВЧ — основной запрос категории
    "ополаскиватель для белья",                  # ВЧ — синоним, высокая частота
    "парфюмированный кондиционер для белья",     # СЧ — прямое попадание
    "кондиционер для белья концентрат",          # СЧ — тип продукта
    "кондиционер для белья ультраконцентрат",    # СЧ — наше УТП
    "кондиционер для белья с ароматом",          # СЧ — сценарий (аромат на белье)
    "кондиционер для белья 1 литр",              # СЧ — наш объём
    "ополаскиватель для белья парфюмированный",  # НЧ — комбо
    "кондиционер для белья длительный аромат",   # НЧ — главная боль рынка
    "кондиционер для белья экономичный",         # НЧ — УТП (100 стирок)
    "aroma emotions кондиционер",                # из конкурента Synergetic
    "кондиционер для белья мягкость",            # сценарий применения
]

# Группа запросов — выбрать из списка ниже исходя из категории товара
# Автомобили / Автотовары / Аксессуары / Антиквариат и коллекционирование /
# Аптека / Бытовая техника / Бытовая химия / Детские товары / Дом и сад /
# Канцелярские товары / Книги / Красота и здоровье / Мебель / Обувь / Одежда /
# Продукты питания / Спортивные товары / Строительство и ремонт /
# Товары для животных / Туризм, рыбалка, охота / Хобби и творчество /
# Цифровые товары / Электроника / Ювелирные украшения
GROUP_NAME = "Бытовая химия"

# Данные товара — для именования папки (из Промпта 1)
BRAND   = "ifoamHOME"
PRODUCT = "КондиционерBalmy_РайскиеЦветы"
ARTICLE = "770758"

# Папка продукта (та же что у всех остальных промптов)
import sys
sys.path.insert(0, "C:/Users/1/Downloads/ozon_prompts_v2")
from save_helpers import get_product_folder
OUTPUT_DIR = get_product_folder(BRAND, PRODUCT, ARTICLE) / "Поисковые запросы"
URL = "https://seller.ozon.ru/app/analytics/what-to-sell/all-queries"

# ══════════════════════════════════════════════════════

async def select_group(page, group_name: str):
    """Выбирает группу запросов в фильтре."""
    print(f"  Выбираю группу '{group_name}'...", end=" ")
    try:
        # Открыть дропдаун группы
        await page.get_by_text("Группа запросов").first.click()
        await asyncio.sleep(1.5)

        # Поле поиска в дропдауне имеет placeholder "Поиск"
        search = page.get_by_placeholder("Поиск")
        await search.wait_for(state="visible", timeout=5000)
        await search.fill(group_name[:4])
        await asyncio.sleep(1)

        # Кликнуть по строке с точным именем группы в списке
        item = page.get_by_role("option", name=group_name).or_(
            page.locator("li, [role='listitem']").filter(has_text=group_name).first
        )
        await item.click()
        await asyncio.sleep(0.5)

        # Нажать "Применить" внутри tippy-попапа
        apply_btn = page.locator("[data-tippy-root] button").filter(has_text="Применить")
        await apply_btn.click()
        await asyncio.sleep(2)
        print("OK")
    except Exception as e:
        print(f"⚠️ не удалось выбрать группу: {e}")


async def run():
    OUTPUT_DIR.mkdir(exist_ok=True)

    async with async_playwright() as p:
        browser = await p.chromium.connect_over_cdp("http://localhost:9222")
        ctx = browser.contexts[0] if browser.contexts else await browser.new_context()
        page = await ctx.new_page()

        print("Открываю страницу...")
        await page.goto(URL, wait_until="domcontentloaded", timeout=30000)
        await asyncio.sleep(4)

        # Выбрать группу запросов один раз
        if GROUP_NAME:
            await select_group(page, GROUP_NAME)

        for query in SEARCH_QUERIES:
            print(f"⏳ '{query}' ...", end=" ", flush=True)

            # Ввести запрос в поле поиска
            search_input = page.get_by_placeholder("Поисковый запрос").or_(
                page.locator("input[type='text']").first
            )
            await search_input.fill(query)
            await asyncio.sleep(0.3)
            await page.keyboard.press("Enter")

            # Ждём загрузки результатов
            await asyncio.sleep(4)

            # Сохранить как PDF
            safe_name = query.replace(" ", "_").replace("/", "-")
            pdf_path = OUTPUT_DIR / f"all_queries_{safe_name}.pdf"
            await page.pdf(
                path=str(pdf_path),
                format="A4",
                print_background=True,
            )
            print(f"-> {pdf_path.name}")

    print(f"\nВсе PDF в: {OUTPUT_DIR}")

if __name__ == "__main__":
    asyncio.run(run())
