import requests, json

headers = {
    'Client-Id': '3315303',
    'Api-Key': '6ee3b7b7-eef3-4ce5-ac54-0b830613c55a',
    'Content-Type': 'application/json',
}

ids = [4540935699]

r = requests.post('https://api-seller.ozon.ru/v3/product/info/list',
    headers=headers, json={"product_id": ids})
data = r.json()
item = data['items'][0]
print(json.dumps(item, indent=2, ensure_ascii=False)[:3000])
