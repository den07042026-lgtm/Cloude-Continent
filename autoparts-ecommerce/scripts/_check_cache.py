import json, sys
f = sys.argv[1] if len(sys.argv) > 1 else "data/analytics/cache/mpbrand_BOSCH_2026-04-20.json"
data = json.load(open(f, encoding="utf-8"))
print(f"Всего записей: {len(data)}")
for r in sorted(data, key=lambda x: x.get("sales", 0), reverse=True)[:20]:
    print(f"{r.get('sales',0):>4} прод | OOS {r.get('lost_profit_percent',0):>3}% | {r.get('avg_price',0):>6} | fbs {r.get('commision_fbs',0)}% | {r['name']}")
