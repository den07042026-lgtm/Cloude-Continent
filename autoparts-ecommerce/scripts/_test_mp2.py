"""Исследование category/brands и правильный формат brand-запроса."""
import requests, os, sys, json, time
from pathlib import Path
from dotenv import load_dotenv

sys.stdout.reconfigure(encoding="utf-8")
load_dotenv(Path(__file__).parent.parent / ".env")
token = os.getenv("MPSTATS_TOKEN", "")
H = {"X-Mpstats-TOKEN": token}
base = "https://mpstats.io/api"
d1, d2 = "2026-04-22", "2026-05-22"

OUR = {"FENOX","TRIALLI","ATE","BOSCH","STARTVOLT","FEBI","SKF","TOYOTA","LUK",
       "VALEO","SACHS","BILSTEIN","KAYABA","WEEN","HOLA","MILES","LYNXAUTO",
       "ICER","BREMBO","TRW","DELPHI","NGKK","DENSO","MANN","MAHLE"}

# 1. Полный список брендов в Автотовары
r = requests.get(f"{base}/wb/get/category/brands",
                 headers=H, params={"path": "Автотовары", "d1": d1, "d2": d2}, timeout=60)
brands_data = r.json() if r.status_code == 200 else []
print(f"Всего брендов в Автотовары: {len(brands_data)}")
print(f"Пример полей: {list(brands_data[0].keys()) if brands_data else []}")
print()

# Фильтруем наши бренды
our_found = [b for b in brands_data if b.get("name","").upper() in OUR]
print(f"Наших брендов найдено: {len(our_found)}")
for b in sorted(our_found, key=lambda x: x.get("items_with_sells", 0), reverse=True):
    print(f"  {b['name']:20s}  товаров с продажами: {b.get('items_with_sells',0):>5}  всего SKU: {b.get('items',0):>6}")

# Топ-30 по items_with_sells
print("\nТоп-30 брендов по активным SKU:")
for b in sorted(brands_data, key=lambda x: x.get("items_with_sells", 0), reverse=True)[:30]:
    marker = " ◄ НАШИ" if b.get("name","").upper() in OUR else ""
    print(f"  {b['name']:30s}  SKU={b.get('items_with_sells',0):>5}{marker}")

time.sleep(0.5)

# 2. Если FENOX есть — пробуем получить его items через /wb/get/brand
# с использованием точного имени из списка
fenox_entry = next((b for b in brands_data if "FENOX" in b.get("name","").upper()), None)
if fenox_entry:
    brand_name = fenox_entry["name"]
    print(f"\nПробуем /wb/get/brand с brand='{brand_name}'")
    r2 = requests.get(f"{base}/wb/get/brand",
                      headers=H,
                      params={"brand": brand_name, "path": "Автотовары",
                              "d1": d1, "d2": d2, "startRow": 0, "endRow": 200},
                      timeout=60)
    print(f"  status={r2.status_code}")
    if r2.status_code == 200 and r2.text.strip():
        rows = r2.json()
        if isinstance(rows, list):
            fenox_rows = [row for row in rows if "FENOX" in row.get("brand","").upper()]
            print(f"  total rows={len(rows)}, FENOX rows={len(fenox_rows)}")
            for row in fenox_rows[:5]:
                print(f"    nmID={row.get('id')} brand={row.get('brand')} {row.get('name','')[:50]}")
            if not fenox_rows:
                print(f"  first item: brand={rows[0].get('brand')} id={rows[0].get('id')}")
else:
    print("\nFENOX не найден в списке брендов категории Автотовары")
    # Поиск частичного совпадения
    partial = [b for b in brands_data if "FENOX" in b.get("name","").upper() or
               any(x in b.get("name","").upper() for x in ["TRIALLI","BOSCH","STARTVOLT"])]
    print(f"Частичные совпадения: {[b['name'] for b in partial[:10]]}")

time.sleep(0.5)

# 3. POST-вариант /wb/get/brand — может требует POST?
print("\n\nПробуем POST /wb/get/brand:")
r3 = requests.post(f"{base}/wb/get/brand",
                   headers={**H, "Content-Type": "application/json"},
                   json={"brand": "FENOX", "path": "Автотовары",
                         "d1": d1, "d2": d2, "startRow": 0, "endRow": 100},
                   timeout=30)
print(f"  POST status={r3.status_code} body={r3.text[:200]!r}")

# 4. GET /wb/get/brands (с path)
print("\nПробуем GET /wb/get/brands?path=Автотовары:")
r4 = requests.get(f"{base}/wb/get/brands",
                  headers=H,
                  params={"path": "Автотовары", "d1": d1, "d2": d2,
                          "startRow": 0, "endRow": 50},
                  timeout=30)
print(f"  status={r4.status_code} body={r4.text[:300]!r}")
