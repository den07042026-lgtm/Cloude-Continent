"""Test WB catalog URL patterns with curl_cffi and mobile API."""
import json
import sys
import time

sys.stdout.reconfigure(encoding="utf-8")

raw = json.loads(open("data/analytics/cache/wb_auto_subjects_2026-04-22.json", encoding="utf-8").read())
auto = [s for s in raw if any(k in s.get("name", "").lower() for k in ("колодк", "фильтр", "тормоз", "амортизатор"))]
subj = auto[0]
print(f"Предмет: id={subj['id']}  name={subj['name']}")

try:
    from curl_cffi import requests as cffi
    sess = cffi.Session(impersonate="chrome124")
    CURL_OK = True
    print("curl_cffi: OK")
except ImportError:
    import requests as cffi
    sess = cffi.Session()
    CURL_OK = False
    print("curl_cffi: НЕТ — используем requests")

# Прогрев
try:
    sess.get("https://www.wildberries.ru/", timeout=12)
    time.sleep(1.5)
    print("WB прогрев: OK")
except Exception as e:
    print(f"WB прогрев: {e}")

dest = -1257786
sid  = subj["id"]

TESTS = [
    # (label, url, params_extra)
    ("catalog/auto-v2",     "https://catalog.wb.ru/catalog/avtomobili-i-mototekhnika/v2/catalog", {}),
    ("catalog/root-v2",     "https://catalog.wb.ru/catalog/0/v2/catalog", {}),
    ("search-xsubject",     "https://search.wb.ru/exactmatch/ru/common/v5/search",
                            {"query": "", "resultset": "catalog", "limit": 15, "spp": 27}),
    ("search-exact-query",  "https://search.wb.ru/exactmatch/ru/common/v5/search",
                            {"query": "колодки тормозные", "resultset": "catalog", "limit": 15, "spp": 27}),
    # WB mobile API
    ("mobile-v2",           "https://search.wb.ru/exactmatch/ru/common/v5/search",
                            {"query": "колодки", "resultset": "catalog", "limit": 10, "appType": 2, "spp": 27}),
    # Новый search эндпоинт
    ("search-v6",           "https://search.wb.ru/exactmatch/ru/common/v6/search",
                            {"query": "колодки тормозные", "resultset": "catalog", "limit": 15, "spp": 27}),
]

base = {"appType": 1, "curr": "rub", "dest": dest, "sort": "popular", "page": 1, "xsubject": sid}

for label, url, extra in TESTS:
    params = {**base, **extra}
    try:
        r = sess.get(url, params=params, timeout=15)
        if r.status_code == 200:
            data = r.json()
            prods = data.get("products") or (data.get("data") or {}).get("products") or []
            print(f"  OK  {len(prods):4} prods  [{label}]  {url}")
            if prods:
                p = prods[0]
                print(f"        nm_id={p.get('id')}  brand={p.get('brand')}  subjectId={p.get('subjectId')}  {str(p.get('name',''))[:40]}")
        else:
            print(f"  {r.status_code}  [{label}]  {url[:70]}")
    except Exception as e:
        print(f"  ERR  [{label}]: {e}")
    time.sleep(2)
