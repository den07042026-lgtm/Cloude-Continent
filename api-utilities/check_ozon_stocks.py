import os
import requests

CLIENT_ID = os.environ.get("OZON_CLIENT_ID", "")
API_KEY = os.environ.get("OZON_API_KEY", "")

headers = {
    "Client-Id": CLIENT_ID,
    "Api-Key": API_KEY,
    "Content-Type": "application/json"
}

# Получаем первые 10 товаров
r = requests.post(
    "https://api-seller.ozon.ru/v2/product/list",
    headers=headers,
    json={"filter": {}, "limit": 10},
    timeout=30
)
print("product/list status:", r.status_code)
data = r.json()
items = data.get("result", {}).get("items", [])
print(f"Товаров в ответе: {len(items)}")

if not items:
    print("Нет товаров:", data)
    exit()

offer_ids = [x.get("offer_id") for x in items[:5] if x.get("offer_id")]
print("Примеры offer_id:", offer_ids)

# Проверяем остатки через v4
r2 = requests.post(
    "https://api-seller.ozon.ru/v4/product/info/stocks",
    headers=headers,
    json={"filter": {"offer_id": offer_ids}, "limit": 100, "last_id": ""},
    timeout=30
)
print("\nstocks v4 status:", r2.status_code)
if r2.status_code == 200:
    items2 = r2.json().get("result", {}).get("items", [])
    for it in items2:
        oid = it.get("offer_id")
        stocks = it.get("stocks", [])
        print(f"  {oid}: {stocks}")
else:
    print(r2.text[:500])

# Статусы товаров
print("\n--- Статусы / видимость ---")
r3 = requests.post(
    "https://api-seller.ozon.ru/v3/product/info/list",
    headers=headers,
    json={"offer_id": offer_ids[:3]},
    timeout=30
)
print("info/list status:", r3.status_code)
if r3.status_code == 200:
    for item in r3.json().get("result", {}).get("items", []):
        oid = item.get("offer_id")
        status = item.get("status", {})
        vis = item.get("visibility_details", {})
        print(f"  {oid}:")
        print(f"    state={status.get('state')} | state_name={status.get('state_name')}")
        print(f"    has_price={vis.get('has_price')} | has_stock={vis.get('has_stock')} | active_product={vis.get('active_product')}")
else:
    print(r3.text[:500])

