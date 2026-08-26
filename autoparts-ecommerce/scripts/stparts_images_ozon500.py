# /// script
# requires-python = ">=3.10"
# dependencies = [
#   "selenium",
#   "webdriver-manager",
#   "openpyxl",
#   "requests",
# ]
# ///
"""
stparts_images_ozon500.py
════════════════════════════════════════════════════════════════════════════
Сбор фото с stparts.ru для ВСЕХ 500 позиций топ-500 Ozon (пользователь сам
выбирает лучшие фото из нескольких источников). Адаптировано из
"Топ ВБ 1306\\stparts_images_1406.py" — та же логика поиска/скачивания,
другой источник артикулов (наш Excel) и папка вывода.

Сохраняет в отдельную подпапку Изображения\\stparts\\, чтобы не перезаписывать
уже найденные фото из Mikado/Yandex.Disk — пользователь потом сравнивает.

Прогресс пишется построчно в stparts_progress.json (code -> статус) — при
повторном запуске уже обработанные коды пропускаются, можно продолжать
батчами (обработка всех 500 через Selenium может занять 2-3 часа).

Все пути захардкожены (см. GUIDE_OZON_500.md граблю #8 — кириллица в argv
ломается и в Bash, и в PowerShell).

Запуск:
  cd C:\\Users\\Admin\\Documents\\Autoparts_Ecommerce
  uv run --with selenium,webdriver-manager,openpyxl,requests scripts/stparts_images_ozon500.py
  uv run ... scripts/stparts_images_ozon500.py --limit 50   # обработать только 50 за раз
"""

import re
import sys
import json
import time
import argparse
from pathlib import Path

import requests
import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

sys.stdout.reconfigure(encoding="utf-8")

# ── Захардкоженные пути (НЕ передавать кириллицу через argv!) ───────────────
EXCEL_PATH    = Path(r"C:\Users\Admin\Desktop\Озон-500\Топ-500_Ozon_2026-07-06.xlsx")
OUT_DIR       = Path(r"C:\Users\Admin\Desktop\Озон-500\Изображения\stparts")
PROGRESS_FILE = Path(r"C:\Users\Admin\Desktop\Озон-500\stparts_progress.json")

SITE_URL    = "https://stparts.ru"
LOGIN_EMAIL = "control.vlz2@gmail.com"
LOGIN_PASS  = "140886continent"
PAGE_WAIT   = 3.0

NOT_FOUND = "not_found"
NO_IMAGE  = "no_image"
OK        = "ok"


def brands_match(a: str, b: str) -> bool:
    a, b = a.strip().upper(), b.strip().upper()
    return bool(a) and bool(b) and (a in b or b in a)


def safe_filename(s: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", s)


def load_progress() -> dict:
    if PROGRESS_FILE.exists():
        try:
            return json.loads(PROGRESS_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_progress(progress: dict):
    PROGRESS_FILE.write_text(json.dumps(progress, ensure_ascii=False, indent=2), encoding="utf-8")


def img_size(driver, el) -> tuple[int, int]:
    try:
        w = driver.execute_script("return arguments[0].naturalWidth;", el) or 0
        h = driver.execute_script("return arguments[0].naturalHeight;", el) or 0
        return w, h
    except Exception:
        return 0, 0


def make_driver() -> webdriver.Chrome:
    opts = Options()
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=opts)
    driver.set_window_size(1400, 900)
    return driver


def _is_logged_in(driver) -> bool:
    try:
        src = driver.page_source
        markers = ["Выход", "Выйти", "выход", "выйти",
                   "подключены как", "Личный кабинет", "личный кабинет",
                   LOGIN_EMAIL, "logout"]
        return any(m in src for m in markers)
    except Exception:
        return False


def do_login(driver):
    print("Авторизация на stparts.ru...")
    driver.get(SITE_URL)
    time.sleep(2)
    if _is_logged_in(driver):
        print("  Уже авторизован.")
        return
    for xpath in ["//*[normalize-space(text())='Войти']", "//*[normalize-space(text())='Вход']",
                  "//a[contains(@href,'login')]", "//a[contains(@href,'auth')]"]:
        try:
            el = driver.find_element(By.XPATH, xpath)
            if el.is_displayed():
                driver.execute_script("arguments[0].click();", el)
                time.sleep(1.5)
                break
        except Exception:
            pass
    for xpath in ["//input[@type='email']", "//input[@name='email']",
                  "//input[@name='login']", "//input[@name='username']",
                  "//input[@id='email']", "//input[@id='login']",
                  "//input[contains(@placeholder,'mail') or contains(@placeholder,'огин')]"]:
        try:
            f = WebDriverWait(driver, 8).until(EC.visibility_of_element_located((By.XPATH, xpath)))
            f.clear(); f.send_keys(LOGIN_EMAIL); break
        except Exception:
            pass
    try:
        pw = WebDriverWait(driver, 8).until(EC.visibility_of_element_located((By.XPATH, "//input[@type='password']")))
        pw.clear(); pw.send_keys(LOGIN_PASS)
    except Exception:
        raise RuntimeError("Поле пароля не найдено на stparts.ru")
    for xpath in ["//button[contains(normalize-space(.),'Вход')]", "//button[contains(normalize-space(.),'Войти')]",
                  "//button[@type='submit']", "//input[@type='submit']"]:
        try:
            btn = driver.find_element(By.XPATH, xpath)
            if btn.is_displayed():
                driver.execute_script("arguments[0].click();", btn); break
        except Exception:
            pass
    time.sleep(3)
    print("  Авторизован." if _is_logged_in(driver) else "  [!] Авторизация не удалась")


def ensure_logged_in(driver):
    if not _is_logged_in(driver):
        print("  Сессия истекла — повторная авторизация...")
        do_login(driver)


def do_search(driver, article: str):
    search_el = None
    for xpath in ["//input[@name='search']", "//input[@type='search']",
                  "//input[contains(@class,'search')]",
                  "//input[contains(@placeholder,'рти') or contains(@placeholder,'оиск')]",
                  "//form//input[@type='text']"]:
        try:
            for el in driver.find_elements(By.XPATH, xpath):
                if el.is_displayed():
                    search_el = el; break
            if search_el:
                break
        except Exception:
            pass
    if search_el is None:
        raise RuntimeError("Поле поиска не найдено")
    search_el.clear(); time.sleep(0.3)
    search_el.send_keys(article); time.sleep(0.4)
    submitted = False
    for btn_xpath in ["./following-sibling::button[1]", "./following-sibling::*[@type='submit'][1]",
                      "./parent::*/button", "./parent::*//*[@type='submit']"]:
        try:
            btn = search_el.find_element(By.XPATH, btn_xpath)
            if btn.is_displayed():
                driver.execute_script("arguments[0].click();", btn); submitted = True; break
        except Exception:
            pass
    if not submitted:
        search_el.send_keys(Keys.ENTER)
    time.sleep(PAGE_WAIT)


def _has_product_card(driver) -> bool:
    skip_words = ("logo", "icon", "banner", "btn", "arrow", "cart", "sprite")
    try:
        for img in driver.find_elements(By.TAG_NAME, "img"):
            src = img.get_attribute("src") or ""
            if not src.startswith("http"):
                continue
            if any(x in src.lower() for x in skip_words):
                continue
            if not img.is_displayed():
                continue
            w, h = img_size(driver, img)
            if w > 150 and h > 150:
                return True
    except Exception:
        pass
    return False


def _get_product_card_brand(driver) -> str | None:
    for xpath in ["//h1", "//h2", "//h3", "//*[contains(@class,'title')]", "//*[contains(@class,'brand')]"]:
        try:
            for el in driver.find_elements(By.XPATH, xpath):
                if not el.is_displayed():
                    continue
                text = el.text.strip()
                if text and 3 <= len(text) <= 100:
                    return text.split()[0]
        except Exception:
            pass
    return None


def find_brand_in_table(driver, excel_brand: str):
    try:
        rows = driver.find_elements(By.XPATH, "//table//tr[.//td]")
    except Exception:
        return None
    for row in rows:
        cells = row.find_elements(By.TAG_NAME, "td")
        if len(cells) < 2:
            continue
        for cell in cells[:3]:
            links = cell.find_elements(By.TAG_NAME, "a")
            txt = (links[0] if links else cell).text.strip().split("\n")[0]
            if txt and 2 <= len(txt) <= 50 and brands_match(excel_brand, txt):
                return cell
    return None


def _get_gallery_urls(driver) -> list[str]:
    try:
        urls = driver.execute_script("""
            var found = []; var seen = {};
            var IMG_EXT = /\\.(jpg|jpeg|png|webp|bmp)(\\?|$)/i;
            var pageW = document.documentElement.scrollWidth || 1400;
            var containers = [
                document.querySelector('.product-images, .product-photo, .item-photo, .product-card, .catalog-item__image, .good-card__image'),
                document.body
            ];
            var selectors = [
                'a[data-fancybox]','a[data-gallery]','a[data-lightbox]',
                'a.fancybox','a[rel="gallery"]','a[rel="fancybox"]',
                'a[href$=".jpg"]','a[href$=".jpeg"]','a[href$=".png"]','a[href$=".webp"]'
            ];
            for (var c = 0; c < containers.length; c++) {
                var root = containers[c]; if (!root) continue;
                for (var s = 0; s < selectors.length; s++) {
                    var links = root.querySelectorAll(selectors[s]); if (!links.length) continue;
                    for (var i = 0; i < links.length; i++) {
                        var a = links[i];
                        var rect = a.getBoundingClientRect();
                        if (rect.left > pageW * 0.6) continue;
                        var img = a.querySelector('img');
                        if (!img) continue;
                        var src = img.src || '';
                        if (!src || src.indexOf('logo') >= 0 || src.indexOf('icon') >= 0) continue;
                        var u = a.href || a.getAttribute('href') || src;
                        if (u && !seen[u] && u.indexOf('http') === 0 && (IMG_EXT.test(u) || a.getAttribute('data-fancybox') || a.getAttribute('data-gallery'))) {
                            seen[u] = true; found.push(u);
                        }
                    }
                    if (found.length) return found;
                }
            }
            return found;""")
        return [u for u in (urls or []) if u.startswith("http")]
    except Exception:
        return []


def _get_main_product_image(driver):
    skip = ("logo", "icon", "banner", "btn", "arrow", "cart", "sprite", "adv", "schatz")
    try:
        page_width = driver.execute_script("return document.documentElement.scrollWidth;") or 1400
    except Exception:
        page_width = 1400
    for xpath in ["//a[@data-fancybox]//img", "//a[@data-gallery]//img",
                  "//a[@data-lightbox]//img", "//a[contains(@class,'fancybox')]//img",
                  "//a[contains(@rel,'gallery') or contains(@rel,'fancybox')]//img",
                  "//a[contains(@href,'.jpg') or contains(@href,'.jpeg') or contains(@href,'.png')]//img"]:
        try:
            for img in driver.find_elements(By.XPATH, xpath):
                if not img.is_displayed():
                    continue
                src = img.get_attribute("src") or ""
                if not src.startswith("http"):
                    continue
                if any(x in src.lower() for x in skip):
                    continue
                x = img.location.get("x", page_width)
                if x > page_width * 0.6:
                    continue
                w, h = img_size(driver, img)
                if w > 50 and h > 50:
                    return img
        except Exception:
            pass
    best_score, best_el = 0, None
    try:
        for img in driver.find_elements(By.TAG_NAME, "img"):
            src = img.get_attribute("src") or ""
            if not src.startswith("http"):
                continue
            if any(x in src.lower() for x in skip):
                continue
            if not img.is_displayed():
                continue
            x = img.location.get("x", page_width)
            if x > page_width * 0.6:
                continue
            w, h = img_size(driver, img)
            if w > 80 and h > 80:
                score = w * h
                if score > best_score:
                    best_score = score; best_el = img
    except Exception:
        pass
    return best_el


def _get_lightbox_image_url(driver) -> str | None:
    for xpath in ["//*[contains(@class,'fancybox')]//img[@src]", "//*[contains(@class,'lightbox')]//img[@src]",
                  "//*[contains(@id,'lightbox')]//img[@src]",
                  "//*[contains(@class,'modal') and not(contains(@class,'fade'))]//img[@src]",
                  "//*[contains(@class,'popup')]//img[@src]"]:
        try:
            best_area, best_url = 0, ""
            for img in driver.find_elements(By.XPATH, xpath):
                src = img.get_attribute("src") or ""
                if not src.startswith("http"):
                    continue
                w, h = img_size(driver, img)
                if w > 100 and w * h > best_area:
                    best_area, best_url = w * h, src
            if best_url:
                return best_url
        except Exception:
            pass
    try:
        best_area, best_url = 0, ""
        for img in driver.find_elements(By.TAG_NAME, "img"):
            src = img.get_attribute("src") or ""
            if not src.startswith("http"):
                continue
            w, h = img_size(driver, img)
            if w > 300 and w * h > best_area:
                best_area, best_url = w * h, src
        if best_url:
            return best_url
    except Exception:
        pass
    return None


def _click_lightbox_next(driver) -> bool:
    for xpath in ["//*[contains(@class,'next') and not(contains(@class,'disabled'))]",
                  "//*[contains(@class,'slick-next') and not(contains(@class,'slick-disabled'))]",
                  "//*[contains(@class,'carousel-control-next')]", "//button[normalize-space(text())='>']"]:
        try:
            for el in driver.find_elements(By.XPATH, xpath):
                if el.is_displayed():
                    driver.execute_script("arguments[0].click();", el); return True
        except Exception:
            pass
    return False


def _close_lightbox(driver):
    try:
        driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
    except Exception:
        pass
    time.sleep(0.5)


def _download(url: str, cookies: dict, save_path: Path) -> bool:
    try:
        resp = requests.get(url, cookies=cookies, headers={"User-Agent": "Mozilla/5.0", "Referer": SITE_URL}, timeout=30)
        resp.raise_for_status()
        if resp.content:
            save_path.write_bytes(resp.content)
            return True
    except Exception:
        pass
    return False


def download_all_images(driver, article: str) -> list[str]:
    base = safe_filename(article)
    cookies = {c["name"]: c["value"] for c in driver.get_cookies()}
    saved: list[str] = []

    gallery_urls = _get_gallery_urls(driver)
    if gallery_urls:
        for idx, url in enumerate(gallery_urls, 1):
            filename = f"{base}_{idx}.jpg"
            if _download(url, cookies, OUT_DIR / filename):
                saved.append(filename)
        return saved

    main_img = _get_main_product_image(driver)
    if not main_img:
        return []
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", main_img)
        time.sleep(0.3)
        driver.execute_script("arguments[0].click();", main_img)
        time.sleep(2.5)
    except Exception:
        return []

    seen: set[str] = set()
    for idx in range(1, 21):
        time.sleep(1.0)
        url = _get_lightbox_image_url(driver)
        if not url or url in seen:
            break
        seen.add(url)
        filename = f"{base}_{idx}.jpg"
        if _download(url, cookies, OUT_DIR / filename):
            saved.append(filename)
        if not _click_lightbox_next(driver):
            break
        time.sleep(1.5)
        new_url = _get_lightbox_image_url(driver)
        if not new_url or new_url == url or new_url in seen:
            break
    _close_lightbox(driver)
    return saved


def process_article(driver, article: str, excel_brand: str):
    try:
        do_search(driver, article)
    except RuntimeError as e:
        print(f"  Ошибка поиска: {e}")
        return NOT_FOUND, []

    page = driver.page_source or ""
    if len(page) < 500:
        return NOT_FOUND, []
    nf_phrases = ["ничего не найдено", "не найдено", "нет результатов", "0 результатов", "not found", "no results"]
    if any(p in page.lower() for p in nf_phrases):
        return NOT_FOUND, []

    if _has_product_card(driver):
        card_brand = _get_product_card_brand(driver)
        if card_brand and brands_match(excel_brand, card_brand):
            files = download_all_images(driver, article)
            return (OK, files) if files else (NO_IMAGE, [])
        brand_el = find_brand_in_table(driver, excel_brand)
        if not brand_el:
            return NOT_FOUND, []
        driver.execute_script("arguments[0].click();", brand_el)
        time.sleep(PAGE_WAIT)
    else:
        brand_el = find_brand_in_table(driver, excel_brand)
        if not brand_el:
            return NOT_FOUND, []
        driver.execute_script("arguments[0].click();", brand_el)
        time.sleep(PAGE_WAIT)

    files = download_all_images(driver, article)
    return (OK, files) if files else (NO_IMAGE, [])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=0, help="Обработать не более N кодов за этот запуск (0 = все оставшиеся)")
    args = ap.parse_args()

    OUT_DIR.mkdir(parents=True, exist_ok=True)

    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    ws = wb["Топ-500"]
    items = [(str(r[1]).strip(), str(r[2] or "").strip()) for r in ws.iter_rows(min_row=2, values_only=True)]
    print(f"Всего позиций в топ-500: {len(items)}")

    progress = load_progress()
    todo = [(code, brand) for code, brand in items if code not in progress]
    print(f"Уже обработано ранее: {len(items) - len(todo)}")
    print(f"Осталось обработать: {len(todo)}")

    if args.limit:
        todo = todo[: args.limit]
        print(f"Ограничение --limit: обработаем {len(todo)} за этот запуск")

    if not todo:
        print("Всё уже обработано!")
        return

    driver = make_driver()
    stats = {"ok": 0, "not_found": 0, "no_image": 0}
    try:
        do_login(driver)
        for i, (code, brand) in enumerate(todo, 1):
            ensure_logged_in(driver)
            print(f"\n[{i}/{len(todo)}] {code} [{brand}]")
            try:
                status, files = process_article(driver, code, brand)
            except Exception as e:
                print(f"  Ошибка: {e}")
                status, files = NOT_FOUND, []
            progress[code] = {"status": status, "files": files}
            save_progress(progress)
            if status == OK:
                print(f"  ✓ {len(files)} фото")
                stats["ok"] += 1
            else:
                print(f"  ✗ {status}")
                stats[status] += 1
    except KeyboardInterrupt:
        print("\nОстановлено пользователем")
    finally:
        driver.quit()
        save_progress(progress)

    print(f"\n{'='*60}")
    print(f"Готово за этот запуск.")
    print(f"  ✓ Найдено фото : {stats['ok']}")
    print(f"  ✗ Не найдено   : {stats['not_found']}")
    print(f"  ✗ Без фото     : {stats['no_image']}")
    print(f"  Прогресс сохранён: {PROGRESS_FILE}")
    print(f"  Всего обработано с начала: {len(progress)} / {len(items)}")


if __name__ == "__main__":
    main()
