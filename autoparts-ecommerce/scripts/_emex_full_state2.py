"""Извлекает и парсит полный state из HTML emex.ru."""
import re, sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

with open(r'data/analytics/emex_debug/products_CAF100493C_CHAMPION.html', encoding='utf-8') as f:
    html = f.read()

# Найдём большой inline JS-строковый литерал с данными
# Подход: найти "details":{"makes" и потом найти внешний {} объект

idx = html.find(r'\"details\":{\"makes\"')
if idx < 0:
    print('Not found')
    exit(1)

# Ищем всю большую строку — она начинается с " и заканчивается на "
# Нам нужен родительский JS строковый литерал
# Отступаем назад и ищем паттерн типа:  someKey="[BIGSTRING]"  или  window.X = "[BIGSTRING]"

# Ищем назад открывающий символ этой строки JS
# Большая строка скорее всего начинается с: ="{"  или :"{"
search_start = max(0, idx - 80000)
chunk = html[search_start:idx+100000]

# Разэкранируем весь фрагмент
decoded = chunk.replace('\\"', '"').replace('\\/', '/').replace('\\n', ' ').replace('\\t', ' ')

# Найдём начало большого JSON объекта — ищем {"catalog": или {"checkout": или {"details":
# Нам нужен самый внешний {} который содержит "details"
idx_details = decoded.find('"details":{"makes"')
print(f'idx_details in decoded: {idx_details}')

# Идём назад ищем { который является родителем "details"
depth = 0
start_pos = idx_details
for i in range(idx_details, 0, -1):
    c = decoded[i]
    if c == '}': depth += 1
    elif c == '{':
        depth -= 1
        if depth == -1:
            start_pos = i
            print(f'Found outer {{ at position {i}')
            break

# Теперь у нас есть start_pos. Вытащим JSON с этой позиции
json_str = decoded[start_pos:]
# Ищем конец объекта
depth = 0
in_str = False
i = 0
while i < len(json_str):
    c = json_str[i]
    if c == '"' and (i == 0 or json_str[i-1] != '\\'):
        in_str = not in_str
    if not in_str:
        if c == '{': depth += 1
        elif c == '}':
            depth -= 1
            if depth == 0:
                json_str = json_str[:i+1]
                break
    i += 1

print(f'JSON string length: {len(json_str)}')
print(f'First 200 chars: {json_str[:200]}')
print(f'Last 200 chars: {json_str[-200:]}')

try:
    state = json.loads(json_str)
    print(f'\nJSON parsed OK! Top keys: {list(state.keys())}')
    with open('data/analytics/emex_debug/full_state_CAF100493C.json', 'w', encoding='utf-8') as f:
        json.dump(state, f, ensure_ascii=False, indent=2)
    print('Сохранено: full_state_CAF100493C.json')
except json.JSONDecodeError as e:
    print(f'JSON error at char {e.pos}: {e}')
    print(f'Context: {json_str[max(0,e.pos-100):e.pos+100]}')
    # Сохраним сырой для отладки
    with open('data/analytics/emex_debug/raw_json_attempt.txt', 'w', encoding='utf-8') as f:
        f.write(json_str[:50000])
