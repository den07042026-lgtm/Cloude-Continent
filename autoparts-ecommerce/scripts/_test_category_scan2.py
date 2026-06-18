"""
Тест 2: /wb/get/category с пагинацией для авто-категорий.
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

# ── 1. Получаем все категории и фильтруем авто-запчасти ──────────────────────
print("=" * 65)
print("1. Авто-категории из /wb/get/categories")
print("=" * 65)

r = requests.get(f"{BASE}/wb/get/categories", headers=H, timeout=60)
all_cats = r.json() if r.status_code == 200 else []

# Фильтр: только пути связанные с запчастями/маслами
keywords = ["запчаст", "масл", "фильтр", "колодк", "аморт", "тормоз",
            "свеч", "ремен", "подшипн", "сайлент", "прокладк", "двигател"]
auto_cats = [
    c for c in all_cats
    if any(kw in c.get("path", "").lower() for kw in keywords)
]
print(f"Всего категорий: {len(all_cats)}")
print(f"Авто-запчасти путей: {len(auto_cats)}")
print("\nПервые 20:")
for c in auto_cats[:20]:
    print(f"  {c['path']}")

# ── 2. Тест /wb/get/category с пагинацией ────────────────────────────────────
print()
print("=" * 65)
print("2. Пагинация /wb/get/category")
print("=" * 65)

def fetch_cat(path, start, end):
    r = requests.get(
        f"{BASE}/wb/get/category",
        headers=H,
        params={"path": path, "d1": d1, "d2": d2, "startRow": start, "endRow": end},
        timeout=60,
    )
    if r.status_code != 200:
        return None, r.status_code, r.text[:150]
    data = r.json()
    if isinstance(data, list):
        return data, 200, ""
    if isinstance(data, dict):
        rows = data.get("data") or data.get("rows") or data.get("items") or []
        return rows if isinstance(rows, list) else [], 200, str(list(data.keys()))
    return [], 200, ""

# Берём первую найденную авто-категорию с достаточной глубиной
test_path = auto_cats[0]["path"] if auto_cats else "Автотовары"
print(f"Тест на: '{test_path}'")
print()

seen_ids = set()
for start in range(0, 1201, 100):
    rows, code, info = fetch_cat(test_path, start, start + 100)
    if rows is None or len(rows) == 0:
        print(f"  start={start:5d}: стоп — {code} {info[:60]}")
        break
    new_ids = {r.get("id") or r.get("nm_id") or r.get("nmId") for r in rows if isinstance(r, dict)}
    new_ids.discard(None)
    overlap  = seen_ids & new_ids
    seen_ids |= new_ids
    print(f"  start={start:5d}: {len(rows):3d} rows | новых={len(new_ids-overlap):3d} | накоплено={len(seen_ids):5d}")
    if len(rows) < 100:
        print(f"  → последняя страница")
        break
    time.sleep(0.3)

print(f"\n  Итого уникальных: {len(seen_ids)}")

# ── 3. Поля одного товара ────────────────────────────────────────────────────
print()
print("=" * 65)
print("3. Поля одного товара")
print("=" * 65)
rows0, _, _ = fetch_cat(test_path, 0, 1)
if rows0 and isinstance(rows0[0], dict):
    print(json.dumps(rows0[0], ensure_ascii=False, indent=2))
else:
    print(f"  Пусто или нет данных. rows0={str(rows0)[:200]}")

# ── 4. Дополнительно: пробуем эндпоинт для листовых subject-путей ────────────
print()
print("=" * 65)
print("4. Тест нескольких авто-путей — сколько товаров")
print("=" * 65)
sample_paths = auto_cats[:8]
for c in sample_paths:
    rows, code, info = fetch_cat(c["path"], 0, 10)
    if rows is not None:
        nm_ids = [r.get("id") or r.get("nm_id") for r in rows[:3] if isinstance(r, dict)]
        print(f"  {len(rows):3d} rows  '{c['path'][:55]}' | nm_ids={nm_ids[:2]}")
    else:
        print(f"  err {code}  '{c['path'][:55]}'")
    time.sleep(0.3)
