"""Анализирует extracted state: OEM, аналоги, применяемость."""
import sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

with open('data/analytics/emex_debug/full_state_CAF100493C.json', encoding='utf-8') as f:
    state = json.load(f)

details = state['details']

print('=== details.originals ===')
originals = details.get('originals', {})
print(f'type: {type(originals).__name__}')
print(json.dumps(originals, ensure_ascii=False, indent=2)[:2000])

print('\n=== details.analogs ===')
analogs = details.get('analogs', {})
print(f'type: {type(analogs).__name__}')
print(json.dumps(analogs, ensure_ascii=False, indent=2)[:2000])

print('\n=== details.replacements ===')
replacements = details.get('replacements', {})
print(f'type: {type(replacements).__name__}')
print(json.dumps(replacements, ensure_ascii=False, indent=2)[:2000])

print('\n=== details.tags ===')
tags = details.get('tags', [])
print(json.dumps(tags, ensure_ascii=False, indent=2)[:500])

print('\n=== details.name/num/make/description ===')
for k in ['name', 'num', 'make', 'description', 'keywords']:
    print(f'  {k}: {details.get(k)}')

print('\n=== parts (раздел применяемости) ===')
parts = state.get('parts', {})
print(f'parts keys: {list(parts.keys())}')
print(json.dumps(parts, ensure_ascii=False, indent=2)[:3000])

print('\n=== maintenance (ТО каталог) ===')
maint = state.get('maintenance', {})
print(f'maintenance keys: {list(maint.keys())}')
