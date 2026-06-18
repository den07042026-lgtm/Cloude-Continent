"""Проверка доступности всех источников данных."""
import requests, time

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36"}
TIMEOUT = 8

tests = [
    # WB APIs
    ("WB catalog auto",   "https://catalog.wb.ru/catalog/avtotovary/v2/catalog?sort=popular&limit=10&cat=9005"),
    ("WB catalog search", "https://search.wb.ru/exactmatch/ru/common/v5/search?query=%D1%84%D0%B8%D0%BB%D1%8C%D1%82%D1%80+%D0%BC%D0%B0%D1%81%D0%BB%D1%8F%D0%BD%D1%8B%D0%B9&sort=popular&resultset=catalog&limit=5"),
    ("WB card.wb.ru",     "https://card.wb.ru/cards/v2/detail?appType=1&curr=rub&dest=-1257786&nm=9783993"),
    ("WB static basket",  "https://static-basket-01.wbbasket.ru/vol0/data/main-menu-ru-ru-v3.json"),
    # Analytics
    ("Shopstat",          "https://shopstat.ru/"),
    ("MPStats",           "https://mpstats.io/"),
    # Car forums
    ("Drive2",            "https://www.drive2.ru/"),
    ("Drom",              "https://www.drom.ru/"),
    # Search
    ("Yandex Wordstat",   "https://wordstat.yandex.ru/"),
    # Price comparison
    ("Exist.ru",          "https://exist.ru/"),
    ("Autodoc",           "https://www.autodoc.ru/"),
]

for name, url in tests:
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True)
        print(f"  OK  {r.status_code}  {name:30s}  {len(r.content):>8} bytes")
    except Exception as e:
        err = str(e)[:60]
        print(f"  FAIL        {name:30s}  {err}")
    time.sleep(0.3)
