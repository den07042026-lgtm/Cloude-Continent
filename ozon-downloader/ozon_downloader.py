"""
Скрипт для скачивания Excel-шаблонов категорий с seller.ozon.ru
Запуск: python ozon_downloader.py
"""

import asyncio
import os
import json
import socket
import subprocess
from datetime import datetime
from playwright.async_api import async_playwright, TimeoutError as PWTimeout

# ───────── Настройки ─────────
CATEGORIES_FILE = r"C:\Users\Admin\Desktop\На загрузку\Категории.txt"
DOWNLOAD_DIR    = r"C:\Users\Admin\Desktop\На загрузку"
IMPORT_URL      = "https://seller.ozon.ru/app/products/import/file"
LOG_FILE        = r"C:\Users\Admin\Desktop\На загрузку\download_log.txt"
PROGRESS_FILE   = r"C:\Users\Admin\Desktop\На загрузку\download_progress.json"
SCREENSHOTS_DIR = r"C:\Users\Admin\Desktop\На загрузку\screenshots"
CHROME_EXE      = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
CHROME_PROFILE  = r"C:\Users\Admin\Desktop\На загрузку\chrome_profile"
CDP_PORT        = 9222
# ─────────────────────────────


def read_categories() -> list[str]:
    categories = []
    seen = set()
    with open(CATEGORIES_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            parts = line.split("\t", 1)
            name = parts[1].strip() if len(parts) == 2 else parts[0].strip()
            if name and name not in seen:
                seen.add(name)
                categories.append(name)
    return categories


def load_progress() -> dict:
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_progress(progress: dict):
    with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
        json.dump(progress, f, ensure_ascii=False, indent=2)


def log(msg: str):
    ts = datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def is_port_open(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(1)
        return s.connect_ex(("127.0.0.1", port)) == 0




async def wait_for_any(page, selectors: list[str], timeout=5000):
    for sel in selectors:
        try:
            el = await page.wait_for_selector(sel, timeout=timeout)
            if el:
                return el, sel
        except Exception:
            pass
    return None, None


async def screenshot(page, name: str):
    os.makedirs(SCREENSHOTS_DIR, exist_ok=True)
    # Убираем символы, недопустимые в именах файлов
    safe = "".join(c if c.isalnum() or c in " _-" else "_" for c in name)
    path = os.path.join(SCREENSHOTS_DIR, f"{safe[:60]}.png")
    try:
        await page.screenshot(path=path, full_page=True)
    except Exception:
        pass


async def dump_page_info(page):
    """Сохраняет скриншот и список всех input/button элементов для диагностики."""
    await screenshot(page, "PAGE_STRUCTURE")
    os.makedirs(SCREENSHOTS_DIR, exist_ok=True)
    try:
        inputs = await page.eval_on_selector_all(
            "input, button, [role='button'], [role='combobox'], [role='searchbox']",
            "els => els.map(e => ({ tag: e.tagName, type: e.type||'', placeholder: e.placeholder||'', class: e.className.substring(0,80), text: (e.innerText||'').substring(0,40) }))"
        )
        dump_path = os.path.join(SCREENSHOTS_DIR, "PAGE_ELEMENTS.txt")
        with open(dump_path, "w", encoding="utf-8") as f:
            for el in inputs:
                f.write(str(el) + "\n")
        log(f"  Диагностика сохранена: {dump_path}")
    except Exception as e:
        log(f"  Ошибка диагностики: {e}")


async def process_category(page, category: str) -> str:
    """Возвращает: 'ok' | 'skip' | 'error'"""
    try:
        await page.goto(IMPORT_URL, wait_until="domcontentloaded", timeout=40000)
        await asyncio.sleep(2)

        # ── Шаг 1: кликаем на поле "Категория и тип" ────────────────────
        field_selectors = [
            'input[placeholder*="атегория" i]',
            'input[placeholder*="Категория" i]',
            '[placeholder*="Категория"]',
            'div[class*="Select"]:has-text("Категория")',
            'div[class*="select"]:has-text("Категория")',
            'div[class*="Input"]:has-text("Категория")',
        ]
        field, _ = await wait_for_any(page, field_selectors, timeout=3000)
        if field:
            await field.click()
        else:
            try:
                await page.get_by_text("Категория и тип", exact=False).first.click()
            except Exception:
                await dump_page_info(page)
                log(f"  ОШИБКА: поле 'Категория и тип' не найдено — {category}")
                return "error"

        await asyncio.sleep(1)

        # ── Поле поиска в модале ─────────────────────────────────────────
        search_input, _ = await wait_for_any(page, [
            'input[placeholder*="азвание категории" i]',
            'input[placeholder*="азвание" i]',
            '[role="dialog"] input',
            'input[type="search"]',
        ], timeout=5000)

        if not search_input:
            await dump_page_info(page)
            log(f"  ОШИБКА: поле поиска в модале не найдено — {category}")
            return "error"

        await search_input.fill(category)
        await asyncio.sleep(2)

        # ── Выбор категории из результатов ──────────────────────────────
        # Ищем строку с точным или частичным совпадением имени
        picked = False
        candidates = await page.query_selector_all("li, [role='option'], [role='listitem']")
        for item in candidates:
            try:
                text = (await item.inner_text()).strip()
                if text.lower() == category.lower():
                    await item.click()
                    picked = True
                    break
            except Exception:
                pass

        if not picked:
            for item in candidates:
                try:
                    text = (await item.inner_text()).strip()
                    if category.lower() in text.lower():
                        await item.click()
                        picked = True
                        break
                except Exception:
                    pass

        if not picked:
            await screenshot(page, f"not_found_{category}")
            log(f"  ПРОПУСК: не найдена в Ozon — «{category}»")
            # Закрываем модал
            try:
                await page.keyboard.press("Escape")
            except Exception:
                pass
            return "skip"

        await asyncio.sleep(0.5)

        # ── Кнопка "Подтвердить" в модале ───────────────────────────────
        confirm_btn, _ = await wait_for_any(page, [
            'button:has-text("Подтвердить")',
            'button:has-text("Выбрать")',
            'button:has-text("Применить")',
        ], timeout=5000)
        if confirm_btn:
            await confirm_btn.click()
            await asyncio.sleep(1)

        # ── Шаг 2: Excel ─────────────────────────────────────────────────
        excel_el, _ = await wait_for_any(page, [
            'label:has-text("Excel")',
            '[class*="radio"]:has-text("Excel")',
            'span:has-text("Excel")',
        ], timeout=5000)
        if excel_el:
            await excel_el.click()
            await asyncio.sleep(0.5)

        # ── Скачать шаблон ───────────────────────────────────────────────
        dl_btn, _ = await wait_for_any(page, [
            'button:has-text("Скачать шаблон")',
            'a:has-text("Скачать шаблон")',
            'button:has-text("Скачать")',
        ], timeout=8000)

        if not dl_btn:
            await screenshot(page, f"no_dl_btn_{category}")
            log(f"  ОШИБКА: кнопка «Скачать шаблон» не найдена — {category}")
            return "error"

        async with page.expect_download(timeout=60000) as dl_info:
            await dl_btn.click()

        download = await dl_info.value
        save_path = os.path.join(DOWNLOAD_DIR, download.suggested_filename)
        await download.save_as(save_path)
        log(f"  OK: {download.suggested_filename}")
        return "ok"

    except PWTimeout as e:
        await screenshot(page, f"timeout_{category}")
        log(f"  ТАЙМАУТ — {category}: {e}")
        return "error"
    except Exception as e:
        await screenshot(page, f"error_{category}")
        log(f"  ОШИБКА — {category}: {e}")
        return "error"


async def main():
    categories = read_categories()
    progress   = load_progress()

    log(f"Категорий всего: {len(categories)}, уже обработано: {len(progress)}")
    remaining = [c for c in categories if c not in progress]
    log(f"Осталось: {len(remaining)}")

    if not remaining:
        log("Все готово. Удалите download_progress.json чтобы начать заново.")
        return

    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    # ── Подключение к Chrome ─────────────────────────────────────────────
    if not is_port_open(CDP_PORT):
        print(f"ОШИБКА: Chrome не запущен с портом отладки {CDP_PORT}.")
        print("Запустите ЗАПУСТИТЬ.bat — он откроет Chrome автоматически.")
        return

    async with async_playwright() as p:
        # Подключаемся к уже запущенному Chrome — без флагов автоматизации
        browser = await p.chromium.connect_over_cdp(f"http://localhost:{CDP_PORT}")
        context = browser.contexts[0] if browser.contexts else await browser.new_context()

        # Настраиваем перехват скачиваний
        context.set_default_timeout(30000)

        pages = context.pages
        page = pages[0] if pages else await context.new_page()

        print("\n" + "=" * 60)
        print("Chrome открыт. Войдите на seller.ozon.ru если нужно.")
        print("Затем нажмите ENTER для начала скачивания.")
        print("=" * 60)
        input()

        ok_count    = sum(1 for v in progress.values() if v == "ok")
        skip_count  = sum(1 for v in progress.values() if v == "skip")
        error_count = sum(1 for v in progress.values() if v == "error")

        for i, category in enumerate(remaining, 1):
            log(f"\n[{i}/{len(remaining)}] {category}")
            result = await process_category(page, category)
            progress[category] = result
            save_progress(progress)

            if result == "ok":
                ok_count += 1
            elif result == "skip":
                skip_count += 1
            else:
                error_count += 1

            await asyncio.sleep(1)

        summary = (
            f"\n{'=' * 60}\n"
            f"ИТОГО: скачано={ok_count}, пропущено={skip_count}, ошибок={error_count}\n"
            f"Лог: {LOG_FILE}"
        )
        log(summary)


if __name__ == "__main__":
    with open(LOG_FILE, "w", encoding="utf-8") as f:
        f.write(f"Ozon Template Downloader — {datetime.now()}\n{'=' * 60}\n")
    asyncio.run(main())
