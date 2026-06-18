"""
Тест MPStats /wb/get/category/items — глубина пагинации для авто-категорий.
"""
import requests, os, sys, time, json
from pathlib import Path
from dotenv import load_dotenv

sys.stdout.reconfigure(encoding="utf-8")
load_dotenv(Path(__file__).parent.parent / ".env")
TOKEN = os.getenv("MPSTATS_TOKEN", "")
H = {"X-Mpstats-TOKEN": TOKEN}
BASE = "https://mpstats.io/api"
d1, d2 = "2026-05-01", "2026-05-25"

def fetch_items(path, start, end):
    r = requests.get(
        f"{BASE}/wb/get/category/items",
        headers=H,
        params={"path": path, "d1": d1, "d2": d2, "startRow": start, "endRow": end},
        timeout=60,
    )
    if r.status_code != 200:
        return None, r.status_code, r.text[:200]
    data = r.json()
    rows = data if isinstance(data, list) else (data or {}).get("data", [])
    return rows if isinstance(rows, list) else [], 200, ""

# ── 1. Ищем правильный path для авто-categories ──────────────────────────────
print("=" * 65)
print("1. Тест форматов path")
print("=" * 65)

test_paths = [
    "Автотовары",
    "Автозапчасти",
    "Автотовары/Автозапчасти",
    "Автозапчасти / Колодки автомобильные",
    "Автозапчасти/Колодки автомобильные",
    "Автотовары/Автозапчасти/Колодки автомобильные",
]

working_path = None
for path in test_paths:
    rows, code, err = fetch_items(path, 0, 10)
    if rows is not None and len(rows) > 0:
        print(f"  ✓ '{path}' → {len(rows)} rows, пример: id={rows[0].get('id')} brand={rows[0].get('brand','')[:20]}")
        if working_path is None:
            working_path = path
    else:
        print(f"  ✗ '{path}' → {code}  {err[:60]}")
    time.sleep(0.5)

# ── 2. Ищем path через /wb/get/category ──────────────────────────────────────
print()
print("=" * 65)
print("2. Список доступных категорий")
print("=" * 65)

r = requests.get(f"{BASE}/wb/get/categories", headers=H, timeout=30)
print(f"  /wb/get/categories → {r.status_code}")
if r.status_code == 200 and r.text.strip() and r.text.strip() != "null":
    cats = r.json()
    if isinstance(cats, list):
        print(f"  Всего: {len(cats)}")
        for c in cats[:10]:
            print(f"    {c}")
    elif isinstance(cats, dict):
        print(f"  Ключи: {list(cats.keys())}")
time.sleep(0.5)

r2 = requests.get(f"{BASE}/wb/get/category", headers=H,
                  params={"path": "Автотовары", "d1": d1, "d2": d2}, timeout=30)
print(f"  /wb/get/category path=Автотовары → {r2.status_code} body={r2.text[:200]!r}")
time.sleep(0.5)

# ── 3. Если нашли рабочий path — тест глубины пагинации ──────────────────────
if working_path:
    print()
    print("=" * 65)
    print(f"3. Пагинация для '{working_path}'")
    print("=" * 65)

    seen_ids = set()
    for start in range(0, 1001, 100):
        rows, code, err = fetch_items(working_path, start, start + 100)
        if rows is None or len(rows) == 0:
            print(f"  start={start:4d}: стоп ({code} {err[:40]})")
            break
        new_ids = {r.get("id") for r in rows}
        overlap  = seen_ids & new_ids
        seen_ids |= new_ids
        print(f"  start={start:4d}: {len(rows):3d} rows | новых={len(new_ids-overlap):3d} | накоплено={len(seen_ids):5d}")
        if len(rows) < 100:
            print(f"  → последняя страница")
            break
        time.sleep(0.4)

    print(f"\n  Итого уникальных nm_id: {len(seen_ids)}")
else:
    print("\n  Рабочий path не найден — пробуем /wb/get/category/brands")
    r3 = requests.get(f"{BASE}/wb/get/category/brands",
                      headers=H, params={"path": "Автотовары", "d1": d1, "d2": d2}, timeout=30)
    print(f"  /wb/get/category/brands → {r3.status_code}")
    if r3.status_code == 200:
        bd = r3.json()
        if isinstance(bd, list):
            print(f"  Брендов: {len(bd)}")
            print(f"  Пример полей: {list(bd[0].keys()) if bd else []}")

# ── 4. Поля одного товара ────────────────────────────────────────────────────
print()
print("=" * 65)
print("4. Поля товара (полный объект)")
print("=" * 65)

path_for_fields = working_path or "Автотовары"
rows, code, err = fetch_items(path_for_fields, 0, 1)
if rows:
    print(json.dumps(rows[0], ensure_ascii=False, indent=2))
else:
    print(f"  Не удалось получить данные: {code} {err}")
