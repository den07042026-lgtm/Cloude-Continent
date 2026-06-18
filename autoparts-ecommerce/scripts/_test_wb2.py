"""Тест альтернативных WB эндпоинтов — card.wb.ru, suggestions, main site."""
import sys, time, json
sys.stdout.reconfigure(encoding="utf-8")
from curl_cffi import requests as cffi_requests

s = cffi_requests.Session(impersonate="chrome124")

# Прогрев основного сайта
try:
    r0 = s.get("https://www.wildberries.ru/", timeout=15)
    print(f"Main site warmup: {r0.status_code}")
except Exception as e:
    print(f"Main site: {e}")
time.sleep(2)

# 1. card.wb.ru с несколькими диапазонами nmID (поиск рабочего диапазона)
print("\n=== card.wb.ru ===")
# Реальные nmID TRIALLI примерно 96-100M (автозапчасти 2022-2023)
# Попробуем несколько диапазонов
test_ids = [
    96700015,   # примерный диапазон автозапчастей
    143212000,
    174316100,
    221900001,
    312000000,
    388835321,  # этот id мы видели в MPStats (бренд "Автотовары")
]
for nm in test_ids:
    try:
        r = s.get("https://card.wb.ru/cards/v1/detail",
                  params={"appType": 1, "curr": "rub", "dest": -1257786, "nm": nm},
                  timeout=10)
        if r.status_code == 200:
            data = r.json()
            prods = (data.get("data") or {}).get("products") or []
            if prods:
                p = prods[0]
                print(f"  nmID={nm} OK: brand={p.get('brand')} name={p.get('name','')[:60]}")
            else:
                print(f"  nmID={nm} 200 но пустой data")
        else:
            print(f"  nmID={nm} {r.status_code} body={r.text[:80]}")
    except Exception as e:
        print(f"  nmID={nm} ERROR: {e}")
    time.sleep(0.3)

# 2. Попробуем 388835321 — этот id нам вернул MPStats
print("\n=== card.wb.ru nmID=388835321 (from MPStats) ===")
r = s.get("https://card.wb.ru/cards/v1/detail",
          params={"appType": 1, "curr": "rub", "dest": -1257786, "nm": 388835321},
          timeout=10)
print(f"status={r.status_code}")
if r.status_code == 200:
    data = r.json()
    prods = (data.get("data") or {}).get("products") or []
    print(f"products={len(prods)}")
    if prods:
        p = prods[0]
        print(f"  brand={p.get('brand')} name={p.get('name','')[:80]}")
        print(f"  subjectId={p.get('subjectId')} subjectName={p.get('subjectName')}")
else:
    print(f"body={r.text[:200]}")
time.sleep(0.5)

# 3. WB suggestions/autocomplete
print("\n=== WB suggestions ===")
try:
    r = s.get(
        "https://search.wb.ru/suggests/api/v3/hint",
        params={"query": "TRIALLI", "dest": -1257786, "lang": "ru"},
        timeout=10
    )
    print(f"suggestions status={r.status_code} body={r.text[:300]}")
except Exception as e:
    print(f"suggestions error: {e}")

# 4. WB seller page (другой домein?)
print("\n=== WB seller API ===")
try:
    r = s.get(
        "https://www.wildberries.ru/webapi/seller/data/short/100111",
        timeout=10
    )
    print(f"seller/data status={r.status_code} body={r.text[:200]}")
except Exception as e:
    print(f"seller/data error: {e}")

# 5. WB популярные товары (main page)
print("\n=== WB popular ===")
try:
    r = s.get(
        "https://www.wildberries.ru/webapi/home/getPopular",
        params={"limit": 5},
        timeout=10
    )
    print(f"popular status={r.status_code} body={r.text[:200]}")
except Exception as e:
    print(f"popular error: {e}")

# 6. Попробуем search.wb.ru без Chrome fingerprint (обычный requests)
import requests as plain_requests
print("\n=== search.wb.ru plain requests (no TLS spoof) ===")
try:
    r = plain_requests.get(
        "https://search.wb.ru/exactmatch/ru/common/v5/search",
        params={"query": "FE43096", "resultset": "catalog", "limit": 5,
                "sort": "popular", "page": 1, "dest": -1257786,
                "appType": 1, "curr": "rub", "spp": 27},
        headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"},
        timeout=10
    )
    print(f"plain status={r.status_code} body={r.text[:200]}")
except Exception as e:
    print(f"plain error: {e}")
