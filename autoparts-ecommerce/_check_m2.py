import sys, sqlite3
sys.stdout.reconfigure(encoding="utf-8")
conn = sqlite3.connect("data/analytics/wb_index.db")

oem_done = conn.execute(
    "SELECT COUNT(DISTINCT our_article) FROM wb_matches WHERE our_source='autoliga' AND method='oem'"
).fetchone()[0]
oem_pairs = conn.execute(
    "SELECT COUNT(*) FROM wb_matches WHERE our_source='autoliga' AND method='oem'"
).fetchone()[0]

print(f"M2 (OEM search) — Автолига:")
print(f"  Артикулов обработано : {oem_done}")
print(f"  Пар артикул→nm_id    : {oem_pairs}")

# Примеры
print("\nПримеры матчей (топ-10 по продажам):")
rows = conn.execute("""
    SELECT m.our_article, m.our_brand, p.name, p.brand, p.price_rub, p.sales_30d, p.oos_pct
    FROM wb_matches m
    JOIN wb_products p ON m.nm_id = p.nm_id
    WHERE m.our_source='autoliga' AND m.method='oem'
      AND p.sales_30d > 0
    ORDER BY p.sales_30d DESC
    LIMIT 10
""").fetchall()
for r in rows:
    print(f"  [{r[0]}] {r[1]} → WB: {r[3]} '{r[2][:40]}' {r[4]:.0f}₽ прод={r[5]:.0f} OOS={r[6]:.0f}%")

conn.close()
