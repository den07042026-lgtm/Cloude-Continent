"""
emex_images_ozon500.py
════════════════════════════════════════════════════════════════════════════
Сбор фото с emex.ru для всех 500 позиций топ-500 Ozon (доп. источник для
сравнения, наравне с mikado_all и уже найденными фото).

Механика:
  1. emex.ru/api/search/search?detailNum={code}&make={brand} — рабочий JSON API
     (уже исследован в проекте, см. scripts/_emex_search_api.py). Возвращает
     makes.list — список брендов/URL, торгующих этим номером.
  2. Ищем в этом списке НАШ бренд (нечёткое сравнение) -> получаем URL карточки
     товара (например /K340119/asp).
  3. Открываем карточку через Playwright (headed - как и Ozon, emex, скорее
     всего, тоже не любит headless) и вытаскиваем фото из DOM.

Прогресс пишется в emex_progress.json (code -> статус), можно продолжать
батчами. Все пути захардкожены (кириллица в argv ломается — см. GUIDE_OZON_500.md).

Запуск:
  cd C:\\Users\\Admin\\Documents\\Autoparts_Ecommerce
  uv run --with playwright,requests,openpyxl scripts/emex_images_ozon500.py
  uv run ... scripts/emex_images_ozon500.py --limit 20
"""

import re
import sys
import json
import time
import random
import argparse
from pathlib import Path

import requests
import openpyxl
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding="utf-8")
import warnings
warnings.filterwarnings("ignore")

EXCEL_PATH    = Path(r"C:\Users\Admin\Desktop\Озон-500\Топ-500_Ozon_2026-07-06.xlsx")
OUT_DIR       = Path(r"C:\Users\Admin\Desktop\Озон-500\Изображения\emex")
PROGRESS_FILE = Path(r"C:\Users\Admin\Desktop\Озон-500\emex_progress.json")

API_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "ru-RU,ru;q=0.9",
    "Origin": "https://emex.ru",
}

OK, NOT_FOUND, NO_IMAGE = "ok", "not_found", "no_image"


def brands_match(a: str, b: str) -> bool:
    a = re.sub(r"[^A-Za-zА-Яа-я0-9]", "", a).upper()
    b = re.sub(r"[^A-Za-zА-Яа-я0-9]", "", b).upper()
    return bool(a) and bool(b) and (a in b or b in a)


def load_progress() -> dict:
    if PROGRESS_FILE.exists():
        try:
            return json.loads(PROGRESS_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_progress(p: dict):
    PROGRESS_FILE.write_text(json.dumps(p, ensure_ascii=False, indent=2), encoding="utf-8")


def find_product_url(session: requests.Session, code: str, brand: str) -> str | None:
    url = f"https://emex.ru/api/search/search?detailNum={code}&make={brand}"
    try:
        r = session.get(url, timeout=15, verify=False)
        if r.status_code != 200:
            return None
        d = r.json()
    except Exception:
        return None
    sr = d.get("searchResult") or {}
    makes = (sr.get("makes") or {}).get("list") or []
    for m in makes:
        if brands_match(brand, m.get("make", "")):
            return "https://emex.ru" + m["url"]
    return None


def extract_images(page) -> list[str]:
    try:
        urls = page.evaluate("""() => {
            const out = new Set();
            const imgs = Array.from(document.querySelectorAll('img'));
            for (const img of imgs) {
                let src = img.currentSrc || img.src || '';
                if (!src) continue;
                // Next.js image optimizer: /_next/image?url=ENCODED&w=..&q=..
                const m = src.match(/\\/_next\\/image\\?url=([^&]+)/);
                if (m) {
                    try { src = decodeURIComponent(m[1]); } catch(e) {}
                }
                if (!src.startsWith('http')) continue;
                if (/logo|icon|sprite|placeholder|avatar/i.test(src)) continue;
                const w = img.naturalWidth || 0, h = img.naturalHeight || 0;
                if (w < 100 || h < 100) continue;
                out.add(src);
            }
            return Array.from(out);
        }""")
        return urls or []
    except Exception:
        return []


def download(url: str, cookies: list[dict], save_path: Path) -> bool:
    try:
        jar = {c["name"]: c["value"] for c in cookies}
        r = requests.get(url, cookies=jar, headers={"User-Agent": API_HEADERS["User-Agent"], "Referer": "https://emex.ru"}, timeout=30, verify=False)
        r.raise_for_status()
        if r.content:
            save_path.write_bytes(r.content)
            return True
    except Exception:
        pass
    return False


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=0, help="Обработать не более N кодов за этот запуск (0 = все оставшиеся)")
    args = ap.parse_args()

    OUT_DIR.mkdir(parents=True, exist_ok=True)

    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    ws = wb["Топ-500"]
    items = [(str(r[1]).strip(), str(r[2] or "").strip()) for r in ws.iter_rows(min_row=2, values_only=True)]
    print(f"Всего позиций: {len(items)}")

    progress = load_progress()
    todo = [(c, b) for c, b in items if c not in progress]
    print(f"Уже обработано: {len(items) - len(todo)}, осталось: {len(todo)}")
    if args.limit:
        todo = todo[: args.limit]
        print(f"--limit: обработаем {len(todo)} за этот запуск")
    if not todo:
        print("Всё уже обработано!")
        return

    api_session = requests.Session()
    api_session.headers.update(API_HEADERS)

    stats = {"ok": 0, "not_found": 0, "no_image": 0}

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, args=["--disable-blink-features=AutomationControlled"])
        ctx = browser.new_context(
            user_agent=API_HEADERS["User-Agent"],
            viewport={"width": 1366, "height": 900},
            locale="ru-RU",
        )
        page = ctx.new_page()

        for i, (code, brand) in enumerate(todo, 1):
            print(f"\n[{i}/{len(todo)}] {code} [{brand}]")
            product_url = find_product_url(api_session, code, brand)
            if not product_url:
                print("  бренд не найден в выдаче emex")
                progress[code] = {"status": NOT_FOUND, "files": []}
                save_progress(progress)
                time.sleep(random.uniform(0.5, 1.0))
                continue

            try:
                page.goto(product_url, timeout=30_000, wait_until="domcontentloaded")
                time.sleep(random.uniform(2.0, 3.5))
                img_urls = extract_images(page)
                cookies = ctx.cookies()
                saved = []
                for idx, u in enumerate(img_urls[:6], 1):
                    fn = f"{code}_{idx}.jpg"
                    if download(u, cookies, OUT_DIR / fn):
                        saved.append(fn)
                if saved:
                    print(f"  ✓ {len(saved)} фото")
                    progress[code] = {"status": OK, "files": saved}
                    stats["ok"] += 1
                else:
                    print("  фото не найдено на странице")
                    progress[code] = {"status": NO_IMAGE, "files": []}
                    stats["no_image"] += 1
            except Exception as e:
                print(f"  ошибка: {e}")
                progress[code] = {"status": NOT_FOUND, "files": []}
                stats["not_found"] += 1
            save_progress(progress)
            time.sleep(random.uniform(1.5, 3.0))

        browser.close()

    print(f"\n{'='*60}")
    print(f"Готово за этот запуск: ok={stats['ok']} not_found={stats['not_found']} no_image={stats['no_image']}")
    print(f"Всего обработано с начала: {len(progress)} / {len(items)}")


if __name__ == "__main__":
    main()
