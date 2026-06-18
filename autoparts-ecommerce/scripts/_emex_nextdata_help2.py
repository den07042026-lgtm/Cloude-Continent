import sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

with open('data/analytics/emex_debug/nextdata_help.json', encoding='utf-8') as f:
    d = json.load(f)

print('TOP KEYS:', list(d.keys()))
pp = d.get('pageProps', {})
init_state_raw = pp.get('initialState', '')
print(f'initialState type: {type(init_state_raw).__name__}  len: {len(str(init_state_raw))}')
if isinstance(init_state_raw, str):
    print(f'First 200: {init_state_raw[:200]}')
    # Попробуем распарсить если это JSON-строка
    try:
        init_state = json.loads(init_state_raw)
        print('Parsed as JSON! Keys:', list(init_state.keys())[:20])
    except:
        print('Not JSON string')
elif isinstance(init_state_raw, dict):
    print('Keys:', list(init_state_raw.keys())[:20])

# Посмотрим initialState на верхнем уровне
top_state = d.get('initialState', '')
print(f'\ntop initialState type: {type(top_state).__name__}  len: {len(str(top_state))}')
if isinstance(top_state, str):
    print(f'First 200: {top_state[:200]}')
