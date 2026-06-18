import requests, json, sys, time
sys.stdout.reconfigure(encoding='utf-8')

token = 'os.environ.get("WB_API_KEY", "")'
headers = {'Authorization': token}

print('Ждём 5 минут для полного сброса rate limit...')
time.sleep(300)

print(f'Отправляем запрос в {time.strftime("%H:%M:%S")}')
r = requests.get('https://common-api.wildberries.ru/api/v1/tariffs/commission',
                 headers=headers, timeout=30)
print(f'Status: {r.status_code}')
print(f'Headers: {dict(r.headers)}')

if r.status_code == 200:
    data = r.json()
    report = data.get('report', [])
    print(f'Категорий: {len(report)}')
    out = 'C:/Users/Admin/wb_commissions.json'
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f'Сохранено в {out}')
    if report:
        print('Поля:', list(report[0].keys()))
        for row in report[:5]:
            print(json.dumps(row, ensure_ascii=False))
else:
    print('Ответ:', r.text[:500])
