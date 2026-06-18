"""Декодирует весь встроенный JSON из HTML emex.ru и показывает структуру."""
import re, sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

with open(r'data/analytics/emex_debug/products_CAF100493C_CHAMPION.html', encoding='utf-8') as f:
    html = f.read()

# Данные хранятся в большом escaped JS-строковом литерале
# Найдём его — ищем большой JSON-объект, экранированный \"
# Обычно это window.__STATE__ = "..." или передаётся через React props

# Ищем блок: большой объект с \"bestPrice\" и \"make\"
# Попробуем извлечь все данные вокруг этого блока
idx = html.find(r'\"bestPrice\"')

# Идём назад до открывающей " большого строкового литерала
# Смотрим на контекст шире
big_chunk = html[max(0, idx-20000):idx+20000]

# Разэкранируем — \" → "
decoded = big_chunk.replace('\\"', '"').replace('\\/', '/').replace('\\n', '\n').replace('\\t', '\t')

# Теперь ищем ключевые слова в разэкранированном тексте
print('=== Разэкранированный фрагмент вокруг bestPrice ===')
idx2 = decoded.find('"bestPrice"')
if idx2 >= 0:
    ctx = decoded[max(0,idx2-3000):idx2+5000]
    print(ctx[:4000])

print()
print('=== Ищем OEM / применяемость в разэкранированном тексте ===')
for kw in ['"oem"', '"oemNumbers"', '"originalNumbers"', '"applies"',
           '"applicability"', '"vehicles"', '"cars"', '"models"',
           '"crossRefs"', '"productGroups"', '"groups"', 'Ford', 'Focus', 'FORD']:
    if kw.lower() in decoded.lower():
        idx3 = decoded.lower().index(kw.lower())
        print(f'\n[{kw}] at {idx3}:')
        print(decoded[max(0,idx3-100):idx3+500])
