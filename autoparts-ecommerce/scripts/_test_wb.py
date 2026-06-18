import sys, time
sys.stdout.reconfigure(encoding="utf-8")
from curl_cffi import requests as cffi_requests

s = cffi_requests.Session(impersonate="chrome124")
s.get("https://www.wildberries.ru/", timeout=15)
time.sleep(2)

r = s.get("https://search.wb.ru/exactmatch/ru/common/v5/search",
    params={"query": "FE43096", "resultset": "catalog", "limit": 5,
            "sort": "popular", "page": 1, "dest": -1257786,
            "appType": 1, "curr": "rub", "spp": 27},
    timeout=20)

print("Status:", r.status_code)
if r.status_code == 200:
    data = r.json()
    prods = data.get("products") or (data.get("data") or {}).get("products") or []
    print("Найдено товаров:", len(prods))
    for p in prods[:5]:
        print(f"  nmID={p['id']}  brand={p.get('brand','?')}  {p.get('name','?')[:60]}")
elif r.status_code == 429:
    print("IP ещё заблокирован (429)")
else:
    print("Ответ:", r.text[:300])
