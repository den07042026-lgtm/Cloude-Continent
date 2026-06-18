"""Извлекает полный Redux/hydration state из HTML emex.ru и сохраняет как JSON."""
import re, sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

with open(r'data/analytics/emex_debug/products_CAF100493C_CHAMPION.html', encoding='utf-8') as f:
    html = f.read()

# Данные находятся в inline JS — ищем паттерн window.__STATE__ или похожие
# Или в styled-component block (скрипт [8] = 178K)

# Найдём большой JSON-объект: ищем все строки типа "key":"value" с вложенностью
# Подход: найти первый '{' после 'window' или после конкретного начала state

# Найдём где начинается Redux state в большом скрипт-блоке
# Ищем ключ "details" (мы знаем что он там есть)
idx = html.find(r'\"details\":{\"makes\"')
if idx < 0:
    # попробуем без экранирования
    idx = html.find('"details":{"makes"')
print(f'details.makes at: {idx}')

# Идём назад, ищем начало большого объекта (скорее всего там есть "catalog":)
# Нам нужен весь state — найдём его начало
search_from = max(0, idx - 80000)
big_part = html[search_from:idx+100000]

# Разэкранируем
decoded = big_part.replace('\\"', '"').replace('\\/', '/').replace('\\n', ' ').replace('\\t', ' ').replace('\\r', ' ')

# Найдём детали
idx2 = decoded.find('"details":{"makes"')
if idx2 >= 0:
    # Вытащим весь объект details
    start = idx2
    brace_depth = 0
    in_str = False
    pos = start
    while pos < len(decoded):
        c = decoded[pos]
        if c == '"' and (pos == 0 or decoded[pos-1] != '\\'):
            in_str = not in_str
        if not in_str:
            if c == '{': brace_depth += 1
            elif c == '}':
                brace_depth -= 1
                if brace_depth == 0:
                    break
        pos += 1
    details_str = decoded[start:pos+1]
    print(f'details object len: {len(details_str)}')
    try:
        details = json.loads(details_str)
        print('JSON OK! Top keys:', list(details.keys()))
        # Сохраним
        with open('data/analytics/emex_debug/details_CAF100493C.json', 'w', encoding='utf-8') as f:
            json.dump(details, f, ensure_ascii=False, indent=2)
        print('Сохранено: details_CAF100493C.json')
        # Покажем структуру
        def show_keys(d, prefix='', depth=0):
            if depth > 3: return
            if isinstance(d, dict):
                for k, v in d.items():
                    t = type(v).__name__
                    extra = f' (len={len(v)})' if isinstance(v, (list, dict)) else f' = {str(v)[:60]}'
                    print(f'  {prefix}{k}: {t}{extra}')
                    if isinstance(v, (dict, list)) and depth < 2:
                        show_keys(v, prefix+'  ', depth+1)
            elif isinstance(d, list) and len(d) > 0:
                print(f'  {prefix}[0]: ', end='')
                show_keys(d[0], prefix+'  ', depth+1)
        show_keys(details)
    except json.JSONDecodeError as e:
        print(f'JSON error: {e}')
        print(details_str[:500])
else:
    print('details.makes not found in decoded chunk')
    # Покажем что есть
    print(decoded[5000:6000])
