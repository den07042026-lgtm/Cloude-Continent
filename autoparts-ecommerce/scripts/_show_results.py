import sqlite3, json, re, xlrd
from pathlib import Path

conn = sqlite3.connect("data/analytics/wb_index.db")

cached  = conn.execute("SELECT COUNT(*) FROM wb_card_oem").fetchone()[0]
with_vc = conn.execute("SELECT COUNT(*) FROM wb_card_oem WHERE oem_list != '[]'").fetchone()[0]
print(f"wb_card_oem: всего={cached:,}, с vendor_code={with_vc:,}")

def norm(s):
    return re.sub(r"[^A-ZА-ЯЁ0-9]", "", s.upper())

# Автолига
files = sorted(Path("data/suppliers/autoliga").glob("*.xls*"),
               key=lambda p: p.stat().st_mtime, reverse=True)
wb2 = xlrd.open_workbook(str(files[0]), encoding_override="cp1251")
ws  = wb2.sheet_by_index(0)
al = {}
for r in range(9, ws.nrows):
    row = ws.row_values(r)
    if len(row) >= 3 and str(row[2]).strip():
        art   = str(row[2]).strip()
        brand = str(row[1]).strip()
        n = norm(art)
        if n:
            al[n] = (art, brand)
print(f"Автолига: {len(al):,} артикулов")

AUTO_KW = (
    "авт", "запч", "тормоз", "амортизатор", "подвеск", "рулев",
    "сцепл", "ремен", "фильтр", "подшипник", "сальник", "датчик",
    "насос", "привод", "генератор", "стартер", "трос", "рычаг",
    "масла", "аккумулятор", "колодк", "прокладк", "шрус", "пружин",
)
kw_conds = " OR ".join(f"LOWER(p.subject) LIKE '%{kw}%'" for kw in AUTO_KW)

# Матчи из уже скачанного (только авто-товары)
rows = conn.execute(f"""
    SELECT p.nm_id, p.brand, p.price_rub, p.sales_30d, c.oem_list
    FROM wb_products p JOIN wb_card_oem c ON p.nm_id=c.nm_id
    WHERE c.oem_list != '[]'
    AND (p.subject IS NULL OR p.subject = '' OR {kw_conds})
""").fetchall()

matches = []
for nm_id, wb_brand, wb_price, sales, oem_j in rows:
    try:
        vcs = json.loads(oem_j)
    except Exception:
        continue
    for vc in vcs:
        n = norm(vc)
        if n and n in al:
            art, al_brand = al[n]
            matches.append((art, al_brand, nm_id, wb_brand, wb_price or 0, sales or 0, vc))

print(f"Матчей в текущем кэше: {len(matches)}")
print()
if matches:
    print(f"{'AL-бренд':15s} {'Артикул':15s} {'vendor_code':18s} {'nm_id':10s} {'WB-бренд':15s} {'WB-цена':8s} {'Продажи/30д':12s}")
    print("-" * 100)
    for art, al_brand, nm_id, wb_brand, wb_price, sales, vc in matches[:30]:
        print(f"{al_brand:15s} {art:15s} {vc!r:18s} {nm_id:<10d} {wb_brand or '':15s} {wb_price:8.0f} {sales:12.0f}")
else:
    print("(пока нет — полный прогон ещё идёт)")

conn.close()
