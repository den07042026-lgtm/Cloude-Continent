import requests, time
url = "https://basket-01.wbbasket.ru/vol23/part2316/2316566/info/ru/card.json"
print("Проверяем CDN...")
t = time.time()
try:
    r = requests.get(url, timeout=10, headers={"User-Agent": "Mozilla/5.0"})
    print(f"Статус: {r.status_code}  время: {time.time()-t:.1f}с")
    if r.status_code == 200:
        d = r.json()
        print(f"vendor_code: {d.get('vendor_code')}")
except Exception as e:
    print(f"Ошибка: {e}  время: {time.time()-t:.1f}с")
