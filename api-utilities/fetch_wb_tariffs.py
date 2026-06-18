import requests, json, sys, time
from datetime import date
sys.stdout.reconfigure(encoding='utf-8')

token = 'eyJhbGciOiJFUzI1NiIsImtpZCI6IjIwMjYwMzAydjEiLCJ0eXAiOiJKV1QifQ.eyJhY2MiOjEsImVudCI6MSwiZXhwIjoxNzk0NjkwODAyLCJpZCI6IjAxOWUzMDBmLWNlNjUtNzg5OC1hYmVlLTZmMTZlMzQ5MjU2OCIsImlpZCI6NDUwOTYyNjgsIm9pZCI6MjUwMTM5MzA1LCJzIjoxNjEyNiwic2lkIjoiMDY4ZTk2NDktYzhiYS00ODg0LWE0NmQtNzE2MDVmMmRmMTFkIiwidCI6ZmFsc2UsInVpZCI6NDUwOTYyNjh9.U-7tU37MCYXY0gVQt92ONiLbOWsNgKXYsJWQr3T_eOKqJNldIMJjrZm5U5v9_Hq3jDM-ope_bfDwSeepMRQjWA'
headers = {'Authorization': token}
today = date.today().isoformat()

# Ждём 70 минут чтобы гарантированно сбросить часовой лимит
WAIT_MINUTES = 70
print(f'Ждём {WAIT_MINUTES} минут...')
time.sleep(WAIT_MINUTES * 60)

for name, url in [
    ('box',    'https://common-api.wildberries.ru/api/v1/tariffs/box'),
    ('return', 'https://common-api.wildberries.ru/api/v1/tariffs/return'),
]:
    print(f'\n=== {name} ({time.strftime("%H:%M:%S")}) ===')
    r = requests.get(url, headers=headers, params={'date': today}, timeout=30)
    print(f'Status: {r.status_code}')
    if r.status_code == 200:
        data = r.json()
        out = f'C:/Users/Admin/wb_tariffs_{name}.json'
        with open(out, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f'Сохранено: {out}')
        print(json.dumps(data, ensure_ascii=False)[:2000])
    else:
        print(r.text[:400])
    # Ждём 5 минут между запросами
    if name == 'box':
        print('Ждём 5 мин перед следующим запросом...')
        time.sleep(310)

print('\nГотово.')
