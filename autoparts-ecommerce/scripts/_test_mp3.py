"""Тест получения items для TRIALLI через MPStats, + WB card API."""
import requests, os, sys, json, time
from pathlib import Path
from dotenv import load_dotenv

sys.stdout.reconfigure(encoding="utf-8")
load_dotenv(Path(__file__).parent.parent / ".env")
token = os.getenv("MPSTATS_TOKEN", "")
H = {"X-Mpstats-TOKEN": token}
base = "https://mpstats.io/api"
d1, d2 = "2026-04-22", "2026-05-22"

# ── 1. Попытки получить TRIALLI items с разными path-вариантами ────────
paths = [
    "Автотовары",
    "Автотовары/Автозапчасти и аксессуары",
    "Запчасти",
    "Автозапчасти",
    "Тормозные диски",
    "Тормозная система",
]
for path in paths:
    r = requests.get(f"{base}/wb/get/brand",
                     headers=H,
                     params={"brand": "TRIALLI", "path": path,
                             "d1": d1, "d2": d2, "startRow": 0, "endRow": 200},
                     timeout=30)
    if r.status_code != 200:
        print(f"brand=TRIALLI path={path!r}: {r.status_code}")
        time.sleep(0.2)
        continue
    rows = r.json() if isinstance(r.json(), list) else r.json().get("data", [])
    trialli = [x for x in rows if "TRIALLI" in x.get("brand","").upper()] if isinstance(rows, list) else []
    total = len(rows) if isinstance(rows, list) else "?"
    print(f"brand=TRIALLI path={path!r}: total={total}, TRIALLI={len(trialli)}")
    if trialli:
        for x in trialli[:3]:
            print(f"   nmID={x.get('id')} {x.get('name','')[:50]}")
        break
    if isinstance(rows, list) and rows:
        print(f"   first brand: {rows[0].get('brand')} id={rows[0].get('id')}")
    time.sleep(0.3)

# ── 2. category/brands с endRow побольше и другим path ─────────────────
print("\n\nPATH Автотовары/Автозапчасти и аксессуары — brands:")
r = requests.get(f"{base}/wb/get/category/brands",
                 headers=H,
                 params={"path": "Автотовары/Автозапчасти и аксессуары", "d1": d1, "d2": d2},
                 timeout=30)
if r.status_code == 200:
    bdata = r.json()
    print(f"Брендов: {len(bdata)}")
    OUR = {"FENOX","TRIALLI","ATE","BOSCH","STARTVOLT","FEBI","SKF","LUK","VALEO","SACHS","KAYABA","MILES","WEEN"}
    found = [b for b in bdata if b.get("name","").upper() in OUR]
    print(f"Наших: {[b['name'] for b in found]}")
    for b in sorted(bdata, key=lambda x: x.get("items_with_sells",0), reverse=True)[:15]:
        marker = " ◄" if b.get("name","").upper() in OUR else ""
        print(f"  {b['name']:25s} SKU={b.get('items_with_sells',0):>5}{marker}")
else:
    print(f"status={r.status_code} body={r.text[:200]}")
time.sleep(0.5)

# ── 3. WB card API (не search, не catalog) — должен быть открыт ────────
from curl_cffi import requests as cffi_requests
wb = cffi_requests.Session(impersonate="chrome124")

# Тест с известным nmID (тормозной диск TRIALLI)
# Попробуем несколько nmID из предыдущих тестов или угадаем диапазон
test_nm_ids = [213500000, 174316000, 221900000, 96700000, 143000000]
print("\n\nWB card API тест:")
for nm in test_nm_ids[:2]:
    try:
        r2 = wb.get(
            "https://card.wb.ru/cards/v1/detail",
            params={"appType": 1, "curr": "rub", "dest": -1257786, "nm": nm},
            timeout=10
        )
        print(f"  nmID={nm} status={r2.status_code}", end="")
        if r2.status_code == 200:
            data = r2.json()
            prods = (data.get("data") or {}).get("products") or []
            if prods:
                p = prods[0]
                print(f" brand={p.get('brand')} {p.get('name','')[:50]}")
            else:
                print(f" keys={list(data.keys())}")
        else:
            print(f" body={r2.text[:100]}")
    except Exception as e:
        print(f" ERROR: {e}")
    time.sleep(0.5)

# ── 4. WB brand-page API ───────────────────────────────────────────────
print("\n\nWB brand page API:")
for brand_slug, supplier_id in [("trialli", None), ("startvolt", None)]:
    # Вариант 1: через catalog.wb.ru с brandName
    try:
        r3 = wb.get(
            "https://catalog.wb.ru/brands/brands/catalog",
            params={"appType": 1, "curr": "rub", "dest": -1257786,
                    "brandName": brand_slug, "limit": 20, "sort": "popular", "page": 1},
            timeout=10
        )
        print(f"  brandName={brand_slug} catalog: {r3.status_code} body={r3.text[:150]}")
    except Exception as e:
        print(f"  brandName={brand_slug} ERROR: {e}")
    time.sleep(0.5)
