"""Извлекает полный state из HTML emex.ru — ищет всю строку целиком."""
import re, sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

with open(r'data/analytics/emex_debug/products_CAF100493C_CHAMPION.html', encoding='utf-8') as f:
    html = f.read()

# Полная разэкранировка всего HTML (медленно, но надёжно)
decoded = html.replace('\\"', '"').replace('\\/', '/').replace('\\n', ' ').replace('\\t', ' ').replace('\\r', ' ')

# Найдём "applePay":{"isAvailable" — начало Redux state
idx_state = decoded.find('"applePay":{"isAvailable"')
if idx_state < 0:
    print('applePay not found, trying basket')
    idx_state = decoded.find('"basket":{"isLoadingReevaluate"')
print(f'State start hint at: {idx_state}')

# Идём назад до открывающей {
depth = 0
start_pos = idx_state
for i in range(idx_state, max(0, idx_state-100), -1):
    c = decoded[i]
    if c == '}': depth += 1
    elif c == '{':
        depth -= 1
        if depth == -1:
            start_pos = i
            break
print(f'Outer {{ at: {start_pos}, char: {repr(decoded[start_pos:start_pos+30])}')

# Найдём конец объекта
depth = 0
in_str = False
i = start_pos
while i < len(decoded):
    c = decoded[i]
    if c == '"' and (i == 0 or decoded[i-1] != '\\'):
        in_str = not in_str
    if not in_str:
        if c == '{': depth += 1
        elif c == '}':
            depth -= 1
            if depth == 0:
                end_pos = i
                break
    i += 1

print(f'JSON span: {start_pos}..{end_pos}  len={end_pos-start_pos}')
json_str = decoded[start_pos:end_pos+1]

try:
    state = json.loads(json_str)
    print(f'JSON OK! Top keys: {list(state.keys())}')
    with open('data/analytics/emex_debug/full_state_CAF100493C.json', 'w', encoding='utf-8') as f:
        json.dump(state, f, ensure_ascii=False, indent=2)
    print('Saved: full_state_CAF100493C.json')

    # Показываем keys внутри details
    details = state.get('details', {})
    print(f'\ndetails keys: {list(details.keys())}')
    makes = details.get('makes', {})
    print(f'makes keys: {list(makes.keys())}')
    print(f'makes.header: {makes.get("header")}')
    makes_list = makes.get('list', [])
    print(f'makes.list count: {len(makes_list)}')
    print('First 5 makes:')
    for item in makes_list[:5]:
        print(f'  {item.get("make"):20}  {item.get("num")}  {item.get("name")}')

except json.JSONDecodeError as e:
    print(f'JSON error at {e.pos}: {e}')
    print(f'Context: ...{json_str[max(0,e.pos-100):e.pos+100]}...')
