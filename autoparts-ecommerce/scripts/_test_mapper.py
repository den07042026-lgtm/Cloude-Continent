"""Тест разных подходов к маппингу артикулов → nmID."""
import requests, os, sys, json, time
from pathlib import Path
from dotenv import load_dotenv

sys.stdout.reconfigure(encoding="utf-8")
load_dotenv(Path(__file__).parent.parent / ".env")
token = os.getenv("MPSTATS_TOKEN", "")
mp_headers = {"X-Mpstats-TOKEN": token}
base = "https://mpstats.io/api"
d1, d2 = "2026-04-22", "2026-05-22"


def mp_get(url, params, label):
    r = requests.get(url, headers=mp_headers, params=params, timeout=60)
    print(f"\n[{label}] status={r.status_code} len={len(r.text)}")
    if r.status_code != 200 or not r.text.strip():
        return []
    data = r.json()
    rows = data.get("data", data) if isinstance(data, dict) else data
    if isinstance(rows, dict):
        # показываем ключи
        print(f"  keys: {list(rows.keys())[:10]}")
        return []
    print(f"  count={len(rows)}")
    if rows:
        print(f"  first item keys: {list(rows[0].keys())[:15]}")
        print(f"  first item: id={rows[0].get('id')} brand={rows[0].get('brand')} name={str(rows[0].get('name',''))[:50]}")
    return rows if isinstance(rows, list) else []


# ──────────────────────────────────────────────
# 1. MPStats /wb/get/brand с явным брендом и разными путями
# ──────────────────────────────────────────────
for brand, path in [
    ("FENOX", "Автотовары"),
    ("FENOX", ""),
    ("TRIALLI", "Автотовары"),
    ("BOSCH", "Автотовары"),
]:
    rows = mp_get(f"{base}/wb/get/brand",
                  {"brand": brand, "path": path, "d1": d1, "d2": d2, "startRow": 0, "endRow": 200},
                  f"brand={brand} path={path!r}")
    fenox_rows = [r for r in rows if r.get("brand","").upper() == brand.upper()]
    print(f"  → из них {brand}: {len(fenox_rows)}")
    if fenox_rows:
        for r in fenox_rows[:3]:
            print(f"     nmID={r.get('id')} {r.get('name','')[:50]}")
    time.sleep(0.5)

# ──────────────────────────────────────────────
# 2. MPStats /wb/get/subject — есть ли такой эндпоинт?
# ──────────────────────────────────────────────
mp_get(f"{base}/wb/get/subject/items",
       {"subject_id": 130, "d1": d1, "d2": d2, "startRow": 0, "endRow": 20},
       "subject/items id=130")

# ──────────────────────────────────────────────
# 3. WB catalog API (не search — другой хост)
# ──────────────────────────────────────────────
from curl_cffi import requests as cffi_requests
wb_s = cffi_requests.Session(impersonate="chrome124")

# 3a. WB brand catalog page
for brand_slug in ["fenox", "trialli"]:
    try:
        r = wb_s.get(
            "https://catalog.wb.ru/brands/search/catalog",
            params={"brand": brand_slug, "limit": 20, "sort": "popular",
                    "page": 1, "dest": -1257786},
            timeout=15
        )
        print(f"\n[WB catalog brand={brand_slug}] status={r.status_code}")
        if r.status_code == 200:
            data = r.json()
            prods = (data.get("data") or {}).get("products") or []
            print(f"  products: {len(prods)}")
            if prods:
                print(f"  first: id={prods[0].get('id')} {prods[0].get('name','')[:50]}")
        else:
            print(f"  body: {r.text[:200]}")
    except Exception as e:
        print(f"  error: {e}")
    time.sleep(1)

# 3b. WB search (проверяем статус)
try:
    r = wb_s.get(
        "https://search.wb.ru/exactmatch/ru/common/v5/search",
        params={"query": "FE43096", "resultset": "catalog", "limit": 5,
                "sort": "popular", "page": 1, "dest": -1257786,
                "appType": 1, "curr": "rub", "spp": 27},
        timeout=15
    )
    print(f"\n[WB search FE43096] status={r.status_code}")
    if r.status_code == 200:
        data = r.json()
        prods = data.get("products") or (data.get("data") or {}).get("products") or []
        print(f"  products: {len(prods)}")
        for p in prods[:3]:
            print(f"  id={p.get('id')} brand={p.get('brand')} {p.get('name','')[:50]}")
    else:
        print(f"  body: {r.text[:200]}")
except Exception as e:
    print(f"  error: {e}")

# 3c. WB seller search по имени бренда (filters/catalog)
try:
    r = wb_s.get(
        "https://search.wb.ru/exactmatch/ru/common/v5/search",
        params={"query": "FENOX", "resultset": "catalog", "limit": 20,
                "sort": "popular", "page": 1, "dest": -1257786,
                "appType": 1, "curr": "rub", "spp": 27},
        timeout=15
    )
    print(f"\n[WB search query=FENOX] status={r.status_code}")
    if r.status_code == 200:
        data = r.json()
        prods = data.get("products") or (data.get("data") or {}).get("products") or []
        print(f"  products: {len(prods)}")
        for p in prods[:3]:
            print(f"  id={p.get('id')} brand={p.get('brand')} {p.get('name','')[:50]}")
except Exception as e:
    print(f"  error: {e}")
