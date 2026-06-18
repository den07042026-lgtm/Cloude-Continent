import sqlite3, sys
sys.stdout.reconfigure(encoding="utf-8")
conn = sqlite3.connect('data/analytics/wb_index.db')

print('=== ТОП subject в wb_products ===')
for s, c in conn.execute('''
    SELECT subject, COUNT(*) cnt FROM wb_products
    GROUP BY subject ORDER BY cnt DESC LIMIT 30
''').fetchall():
    print(f'  {c:6d}  {s}')

print()
empty = conn.execute("SELECT COUNT(*) FROM wb_products WHERE subject IS NULL OR subject=''").fetchone()[0]
total = conn.execute('SELECT COUNT(*) FROM wb_products').fetchone()[0]
print(f'Пустых subject: {empty} из {total}')

print('\n=== Примеры матчей НЕ авто (subject не содержит авто-кл.слов) ===')
AUTO_KW = ["автозапчаст","запчаст","колодк","амортизатор","тормоз",
           "подвеск","рулев","сцеплен","ремен","шаровая","стойк",
           "свеч","зажиган","фильтр","подшипник","сальник","прокладк",
           "датчик","насос","шрус","привод","масла","аккумулятор",
           "генератор","стартер","ступиц","зеркал","трос","наконечник",
           "рычаг","пружин","ремкомплект","уплотнитель","реле"]
rows = conn.execute('''
    SELECT p.subject, p.name, m.our_article, m.method
    FROM wb_matches m JOIN wb_products p ON m.nm_id=p.nm_id
    LIMIT 5000
''').fetchall()
shown = 0
for subj, name, art, method in rows:
    s = (subj or '').lower()
    if s and not any(kw in s for kw in AUTO_KW):
        print(f'  [{method}] art={art}  subj={subj}  name={name[:60]}')
        shown += 1
        if shown >= 20:
            break
if shown == 0:
    print('  (все субъекты содержат авто-ключевые слова)')
conn.close()
