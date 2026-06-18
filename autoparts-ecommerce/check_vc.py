import sqlite3
conn = sqlite3.connect("data/analytics/wb_index.db")
total = conn.execute("SELECT COUNT(*) FROM wb_products").fetchone()[0]
fetched = conn.execute("SELECT COUNT(*) FROM vendor_codes WHERE fetched=1").fetchone()[0]
with_vc = conn.execute("SELECT COUNT(*) FROM vendor_codes WHERE fetched=1 AND vc_raw IS NOT NULL AND vc_raw != ''").fetchone()[0]
print(f"Всего nm_ids: {total}")
print(f"Запрошено: {fetched}")
print(f"С vendorCode: {with_vc}")
print(f"Осталось: {total - fetched}")
conn.close()
