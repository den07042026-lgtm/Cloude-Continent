"""Тест MPStats /wb/get/brand — пагинация и сколько товаров у наших брендов."""
import requests, os, json, sys, time
from dotenv import load_dotenv
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
load_dotenv(Path(__file__).parent.parent / ".env")
TOKEN = os.getenv("MPSTATS_TOKEN", "")
H = {"X-Mpstats-TOKEN": TOKEN}
d1, d2 = "2026-04-01", "2026-04-30"
URL = "https://mpstats.io/api/wb/get/brand"


def fetch_brand(brand, start=0, end=100):
    r = requests.get(URL, headers=H,
        params={"path": brand, "d1": d1, "d2": d2, "startRow": start, "endRow": end},
        timeout=30)
    if r.status_code != 200:
        return None, r.status_code
    data = r.json()
    rows = data if isinstance(data, list) else (data or {}).get("data", [])
    return rows if isinstance(rows, list) else [], 200


# 1. Zekkert pagination
print("=== Zekkert пагинация ===")
all_ids = set()
for start in range(0, 601, 100):
    rows, code = fetch_brand("Zekkert", start, start + 100)
    if rows is None or len(rows) == 0:
        print(f"  start={start}: {code if rows is None else 'empty'} — стоп")
        break
    new_ids = {r["id"] for r in rows}
    overlap = all_ids & new_ids
    all_ids |= new_ids
    print(f"  start={start}: {len(rows)} rows, новых={len(new_ids - overlap)}, накоплено={len(all_ids)}")
    if len(rows) < 100:
        print(f"  → последняя страница")
        break
    time.sleep(0.5)

print(f"Zekkert всего nm_id: {len(all_ids)}")

# 2. Несколько ключевых брендов — сколько у каждого
print("\n=== Топ-брендов каталога: кол-во товаров на WB через MPStats ===")
brands_test = ["TRIALLI", "FENOX", "TRW", "KAYABA", "SKF", "BOSCH", "Zekkert"]
for brand in brands_test:
    rows, code = fetch_brand(brand, 0, 100)
    if rows is None:
        print(f"  {brand}: HTTP {code}")
    else:
        print(f"  {brand}: {len(rows)} (первые 100, может быть больше)")
    time.sleep(0.3)

# 3. Один пример — показать ключевые поля первой записи Zekkert
print("\n=== Zekkert[0] — все полезные поля ===")
rows, _ = fetch_brand("Zekkert", 0, 1)
if rows:
    r = rows[0]
    for k in ["id", "name", "brand", "subject_id", "subject", "sales", "revenue",
              "lost_profit_percent", "final_price", "commission_fbs", "balance", "is_fbs"]:
        print(f"  {k}: {r.get(k)}")
