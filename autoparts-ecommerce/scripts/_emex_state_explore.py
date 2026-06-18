"""Исследует полный state emex.ru — что из данных уже есть в initial load."""
import sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

with open('data/analytics/emex_debug/full_state_CAF100493C.json', encoding='utf-8') as f:
    state = json.load(f)

details = state['details']

# Показываем все ненулевые поля в details
print('=== details non-empty fields ===')
for k, v in details.items():
    if v is None: continue
    if isinstance(v, bool):
        if v: print(f'  {k}: True')
        continue
    if isinstance(v, (int, float)):
        if v != 0: print(f'  {k}: {v}')
        continue
    if isinstance(v, str):
        if v: print(f'  {k}: {v[:100]}')
        continue
    if isinstance(v, list):
        if v and not (len(v) == 1 and v[0] is None):
            print(f'  {k}: list[{len(v)}]')
            for item in v[:3]:
                print(f'    {json.dumps(item, ensure_ascii=False)[:200]}')
        continue
    if isinstance(v, dict):
        non_empty = {kk: vv for kk, vv in v.items() if vv}
        if non_empty:
            print(f'  {k}: dict keys={list(non_empty.keys())[:8]}')
            if k == 'makes':
                print(f'    header: {v.get("header")}')
                lst = v.get('list', [])
                print(f'    list[{len(lst)}]:')
                for item in lst:
                    print(f'      {item.get("make"):20} {item.get("num")} {item.get("name")}')

print()
print('=== suggestions (если есть) ===')
suggestions = details.get('suggestions', [])
print(f'type: {type(suggestions).__name__}')
if isinstance(suggestions, list):
    print(f'count: {len(suggestions)}')
    for s in suggestions[:5]:
        print(f'  {s}')
elif isinstance(suggestions, dict):
    print(json.dumps(suggestions, ensure_ascii=False, indent=2)[:1000])

print()
print('=== promoEnrichData ===')
ped = details.get('promoEnrichData', {})
print(json.dumps(ped, ensure_ascii=False, indent=2)[:1000])

print()
print('=== productPopupInfo ===')
ppi = details.get('productPopupInfo', {})
print(json.dumps(ppi, ensure_ascii=False, indent=2)[:1000])

# Ищем в search/части каталога
print()
print('=== Другие топ-уровневые секции с данными ===')
for key, val in state.items():
    if key == 'details': continue
    if isinstance(val, dict):
        non_empty_keys = [k for k, v in val.items() if v and v != [] and v != {} and v is not None]
        if non_empty_keys:
            print(f'  {key}: {non_empty_keys[:10]}')
