import sys, sqlite3
sys.stdout.reconfigure(encoding="utf-8")
conn = sqlite3.connect("data/analytics/wb_index.db")
rows = conn.execute("""
    SELECT p.subject, COUNT(*) as cnt
    FROM vendor_codes v
    JOIN wb_products p ON p.nm_id = v.nm_id
    WHERE v.vc_raw IS NOT NULL AND v.vc_raw != ''
    GROUP BY p.subject
    ORDER BY cnt DESC
    LIMIT 60
""").fetchall()
for r in rows:
    print(f"  {r[1]:5d}  {r[0]}")
conn.close()
