import sqlite3
conn = sqlite3.connect("data/analytics/wb_index.db")

print("=== wb_matches: структура и примеры ===")
total = conn.execute("SELECT COUNT(*) FROM wb_matches").fetchone()[0]
print(f"Всего пар: {total}")

rows = conn.execute("SELECT * FROM wb_matches WHERE our_source='autoliga' LIMIT 5").fetchall()
print(f"Колонки: {[d[0] for d in conn.execute('SELECT * FROM wb_matches LIMIT 1').description]}")
for r in rows:
    print(f"  {r}")

print("\n=== Score диапазон (autoliga) ===")
for r in conn.execute("""
    SELECT
        SUM(CASE WHEN CAST(score AS REAL) >= 0.9 THEN 1 ELSE 0 END) as s09,
        SUM(CASE WHEN CAST(score AS REAL) >= 0.7 AND CAST(score AS REAL) < 0.9 THEN 1 ELSE 0 END) as s07,
        SUM(CASE WHEN CAST(score AS REAL) >= 0.5 AND CAST(score AS REAL) < 0.7 THEN 1 ELSE 0 END) as s05,
        SUM(CASE WHEN CAST(score AS REAL) < 0.5 THEN 1 ELSE 0 END) as slow,
        COUNT(DISTINCT our_article) as uniq_articles
    FROM wb_matches WHERE our_source='autoliga'
""").fetchall():
    print(f"  score>=0.9: {r[0]}  0.7-0.9: {r[1]}  0.5-0.7: {r[2]}  <0.5: {r[3]}")
    print(f"  Уникальных артикулов: {r[4]}")

print("\n=== Пример хороших совпадений (score>=0.9) ===")
for r in conn.execute("""
    SELECT m.our_article, m.our_brand, CAST(m.score AS REAL) as sc,
           p.name, p.brand, p.price_rub, p.sales_30d, p.oos_pct
    FROM wb_matches m
    JOIN wb_products p ON m.nm_id = p.nm_id
    WHERE m.our_source='autoliga' AND CAST(m.score AS REAL) >= 0.9
    ORDER BY p.oos_pct DESC LIMIT 10
""").fetchall():
    print(f"  [{r[0]}] {r[1]} sc={r[2]:.2f} → '{r[4]} {r[3][:25]}' WB={r[5]}₽ OOS={r[7]}%")

print("\n=== Топ по дефициту (score>=0.7, oos>30%) ===")
for r in conn.execute("""
    SELECT m.our_article, m.our_brand, CAST(m.score AS REAL) as sc,
           p.name, p.brand, p.price_rub, p.sales_30d, p.oos_pct
    FROM wb_matches m
    JOIN wb_products p ON m.nm_id = p.nm_id
    WHERE m.our_source='autoliga' AND CAST(m.score AS REAL) >= 0.7 AND p.oos_pct >= 30
    ORDER BY p.oos_pct DESC, p.sales_30d DESC LIMIT 15
""").fetchall():
    print(f"  [{r[0]}] sc={r[2]:.2f} OOS={r[7]}% прод={r[6]} WB={r[5]}₽ | {r[4]} {r[3][:30]}")

conn.close()
