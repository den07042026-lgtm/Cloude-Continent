"""Находит и декодирует JSON, встроенный в HTML emex.ru."""
import re, sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

with open(r'data/analytics/emex_debug/products_CAF100493C_CHAMPION.html', encoding='utf-8') as f:
    html = f.read()

# Ищем экранированный JSON (\"bestPrice\")
idx = html.find(r'\"bestPrice\"')
print(f'Позиция escaped bestPrice: {idx}')

if idx > 0:
    # Данные вложены в строку JavaScript — нужно найти границы строки
    # Идём назад до начала строки (кавычка без предшествующего \)
    chunk = html[max(0, idx-2000):idx+5000]
    print('Контекст [-2000..+5000]:')
    print(chunk[:1000])
    print('...')
    print(chunk[-1000:])

print()
# Попробуем выдернуть весь JS-строковый литерал вокруг этих данных
# Ищем паттерн: начало = "productGroups" или "groups"
for kw in [r'\"productGroups\"', r'\"groups\"', r'\"analogs\"', r'\"crossRefs\"']:
    cnt = html.count(kw)
    if cnt > 0:
        idx2 = html.index(kw)
        print(f'FOUND {kw} x{cnt} at {idx2}:')
        print(html[max(0,idx2-100):idx2+500])
        print()
