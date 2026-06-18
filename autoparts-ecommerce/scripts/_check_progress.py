import sqlite3, json, re, xlrd
from pathlib import Path

conn = sqlite3.connect("data/analytics/wb_index.db")

def norm(s):
    return re.sub(r"[^A-ZА-ЯЁ0-9]", "", s.upper())

# Загружаем Автолигу
files = sorted(Path("data/suppliers/autoliga").glob("*.xls*"),
               key=lambda p: p.stat().st_mtime, reverse=True)
wb2 = xlrd.open_workbook(str(files[0]), encoding_override="cp1251")
ws = wb2.sheet_by_index(0)
al = {}
for r in range(9, ws.nrows):
    row = ws.row_values(r)
    if len(row) >= 3 and str(row[2]).strip():
        art   = str(row[2]).strip()
        brand = str(row[1]).strip()
        n = norm(art)
        if n:
            al[n] = (art, brand)

# Что за subjects у TRIALLI продуктов в кэше
print("=== TRIALLI продукты в wb_card_oem ===")
trialli = conn.execute("""
    SELECT p.nm_id, p.brand, p.subject, c.oem_list
    FROM wb_products p JOIN wb_card_oem c ON p.nm_id=c.nm_id
    WHERE UPPER(p.brand) LIKE '%TRIALLI%' AND c.oem_list != '[]'
    LIMIT 10
""").fetchall()
for nm_id, brand, subj, oem_j in trialli:
    vcs = json.loads(oem_j)
    vc = vcs[0] if vcs else ""
    n = norm(vc)
    match = "MATCH!" if n in al else ""
    print(f"  nm={nm_id} subj={subj!r:30s} vc={vc!r:15s} norm={n!r} {match}")

print()
# Всё без фильтра subject
rows_all = conn.execute("""
    SELECT p.nm_id, p.brand, p.price_rub, c.oem_list
    FROM wb_products p JOIN wb_card_oem c ON p.nm_id=c.nm_id
    WHERE c.oem_list != '[]'
""").fetchall()
matches_all = []
for nm_id, wb_brand, wb_price, oem_j in rows_all:
    vcs = json.loads(oem_j)
    for vc in vcs:
        n = norm(vc)
        if n and n in al:
            matches_all.append((al[n][0], al[n][1], nm_id, wb_brand, wb_price, vc))

print(f"Матчей без фильтра: {len(matches_all)}")
for m in matches_all[:15]:
    art, al_brand, nm_id, wb_brand, wb_price, vc = m
    print(f"  {al_brand:12s} | {art:15s} | vc={vc!r:18s} | nm={nm_id} | wb={wb_brand} | {wb_price}руб")
conn.close()
