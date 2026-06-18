import sys, json, sqlite3, requests
sys.stdout.reconfigure(encoding="utf-8")

_BASKET_LIMITS = [
    (143,"01"),(287,"02"),(431,"03"),(575,"04"),(719,"05"),(863,"06"),
    (1007,"07"),(1151,"08"),(1295,"09"),(1439,"10"),(1583,"11"),(1727,"12"),
    (1871,"13"),(2015,"14"),(2159,"15"),(2303,"16"),(2591,"17"),(2879,"18"),
    (3167,"19"),(3455,"20"),(3743,"21"),(4031,"22"),(4319,"23"),(6399,"24"),
    (8191,"25"),(10239,"26"),(12287,"27"),
]

def basket_url(nm_id):
    vol  = nm_id // 100_000
    part = nm_id // 1_000
    basket = "28"
    for limit, num in _BASKET_LIMITS:
        if vol <= limit:
            basket = num
            break
    return f"https://basket-{basket}.wbbasket.ru/vol{vol}/part{part}/{nm_id}/info/ru/card.json"

conn = sqlite3.connect("data/analytics/wb_index.db")

# Берём 5 nm_id у которых есть vendorCode из авто-категорий
rows = conn.execute("""
    SELECT v.nm_id, v.vc_raw, p.name, p.brand, p.subject
    FROM vendor_codes v
    JOIN wb_products p ON p.nm_id = v.nm_id
    WHERE v.vc_raw IS NOT NULL AND v.vc_raw != ''
      AND p.subject LIKE 'Автозапчасти%'
    LIMIT 5
""").fetchall()
conn.close()

s = requests.Session()
s.headers["User-Agent"] = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/124.0"

for nm_id, vc_raw, name, brand, subject in rows:
    url = basket_url(nm_id)
    r = s.get(url, timeout=10)
    print(f"\n{'='*65}")
    print(f"nm_id={nm_id}  brand={brand}  subject={subject}")
    print(f"name={name[:60]}")
    print(f"vc_raw (в БД): {vc_raw}")
    if r.status_code == 200:
        data = r.json()
        # Печатаем все поля верхнего уровня
        top = {k: v for k, v in data.items() if not isinstance(v, (list, dict))}
        print(f"Поля верхнего уровня: {json.dumps(top, ensure_ascii=False)}")
        # Ищем характеристики
        for key in ("params", "options", "characteristics", "grouped_params"):
            if key in data and data[key]:
                print(f"  [{key}]:")
                for item in data[key][:10]:
                    print(f"    {item}")
    else:
        print(f"HTTP {r.status_code}")
