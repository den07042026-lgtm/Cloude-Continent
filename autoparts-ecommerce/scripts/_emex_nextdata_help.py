"""Анализирует структуру help.json для понимания initialState формата."""
import sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

with open('data/analytics/emex_debug/nextdata_help.json', encoding='utf-8') as f:
    d = json.load(f)

print('TOP KEYS:', list(d.keys()))
print()

pp = d.get('pageProps', {})
print('pageProps keys:', list(pp.keys()))

init_state = pp.get('initialState', {})
if not init_state:
    init_state = d.get('initialState', {})
print('initialState keys:', list(init_state.keys())[:20])

# Смотрим details
details = init_state.get('details', {})
if details:
    print('\ninitialState.details keys:', list(details.keys())[:20])
    makes = details.get('makes', {})
    if makes:
        print(f'  makes.header: {makes.get("header")}')
        print(f'  makes.list: {len(makes.get("list", []))} items')
    originals = details.get('originals', [])
    print(f'  originals: {originals}')
    analogs = details.get('analogs', [])
    print(f'  analogs: {analogs}')
else:
    print('\nNo details in initialState')
    print('Full initialState:', json.dumps(init_state, ensure_ascii=False, indent=2)[:2000])

# Смотрим initialProps
init_props = d.get('initialProps', {})
print('\ninitialProps:', json.dumps(init_props, ensure_ascii=False, indent=2)[:1000])
