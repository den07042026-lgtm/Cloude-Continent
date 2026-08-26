import sys, openpyxl, requests
sys.stdout.reconfigure(encoding='utf-8')

# 1. Ищем в прайсе
wb = openpyxl.load_workbook(r'C:/Users/Admin/Desktop/mikado_price_34 (3).xlsx', read_only=True, data_only=True)
ws = wb.active
rows = list(ws.iter_rows(values_only=True))
header = [str(v).strip().lower() if v else '' for v in rows[0]]

code_idx  = header.index('code')
price_idx = header.index('priceout')
qty_idx   = header.index('qty')

target = 'sg-4537'
found = []
for row in rows[1:]:
    code = str(row[code_idx]).strip().lower() if row[code_idx] else ''
    if code == target:
        price = float(str(row[price_idx] or 0))
        qty   = int(float(str(row[qty_idx] or 0)))
        found.append((code, price, qty, row))

wb.close()
print(f"В прайсе Mikado: {found}")

if not found:
    print("Артикул не найден")
    exit()

purchase = found[0][1]
qty      = found[0][2]
print(f"\nЗакупочная цена: {purchase} ₽")
print(f"Наличие:         {qty} шт.")

# 2. Получаем volume_weight из Ozon
headers = {
    'Client-Id': '3315303',
    'Api-Key':   '6ee3b7b7-eef3-4ce5-ac54-0b830613c55a',
    'Content-Type': 'application/json',
}
offer_id = target + '-con'
r = requests.post('https://api-seller.ozon.ru/v3/product/list',
    headers=headers,
    json={'filter': {'offer_id': [offer_id]}, 'last_id': '', 'limit': 10})
items = r.json().get('result', {}).get('items', [])
print(f"\nOzon поиск по offer_id '{offer_id}': {items}")

product_id = items[0]['product_id'] if items else None
volume_weight = None
current_price = None

if product_id:
    r2 = requests.post('https://api-seller.ozon.ru/v3/product/info/list',
        headers=headers, json={'product_id': [product_id]})
    info = r2.json().get('items', [{}])[0]
    volume_weight = info.get('volume_weight')
    current_price = info.get('price')
    print(f"Текущая цена на Ozon: {current_price} ₽")
    print(f"Объёмный вес: {volume_weight} кг")

# 3. Расчёт по формуле калькулятора
import math

OTHER    = 20
ACQ_PCT  = 0.015
TAX_PCT  = 0.06
RET_RATE = 0.03
REVERSE  = 80
LOG_FBS  = [(0.5,75),(1,90),(2,115),(5,155),(10,210),(15,265),(20,315),(25,365),(30,420),(50,620)]
FBS_TIERS= [100, 300, 1500, 5000, 10000]
FBS_RATES= [0.14, 0.20, 0.44, 0.44, 0.44, 0.44]

def log_cost(w):
    for lim, cost in LOG_FBS:
        if w <= lim: return cost
    return LOG_FBS[-1][1] + math.ceil(w - LOG_FBS[-1][0]) * 15

def fbs_rate(sell):
    for t, r in zip(FBS_TIERS, FBS_RATES):
        if sell < t: return r
    return FBS_RATES[-1]

def calc(purchase, sell, logistics):
    comm      = sell * fbs_rate(sell)
    acq       = sell * ACQ_PCT
    ret_loss  = RET_RATE * (logistics + REVERSE)
    proceeds  = sell - comm - acq - logistics
    tax       = max(0.0, proceeds) * TAX_PCT
    total_cost= purchase + comm + acq + logistics + ret_loss + OTHER + tax
    profit    = sell - total_cost
    markup    = profit / (purchase + OTHER) * 100
    return dict(comm=comm, acq=acq, ret_loss=ret_loss, tax=tax,
                total_cost=total_cost, profit=profit, markup=markup)

def optimal_markup(p):
    if p < 500:  return 0.30
    if p < 1200: return 0.25
    if p < 2500: return 0.22
    if p < 3500: return 0.20
    return 0.17

vw = float(volume_weight) if volume_weight else 0.5
logistics = log_cost(vw)
target_markup = optimal_markup(purchase)

# Ищем цену Оптимум
opt_price = None
for s in range(50, 500_001):
    r_ = calc(purchase, s, logistics)
    if r_['profit'] / (purchase + OTHER) >= target_markup - 1e-6:
        opt_price = s
        break

print(f"\n{'='*50}")
print(f"РАСЧЁТ ПО ФОРМУЛЕ ОПТИМУМ")
print(f"{'='*50}")
print(f"Закупка:           {purchase:.0f} ₽")
print(f"Упаковка:          {OTHER} ₽")
print(f"Объёмный вес:      {vw} кг  →  логистика {logistics:.0f} ₽")
print(f"Целевая наценка:   {target_markup*100:.0f}% (тир для {purchase:.0f} ₽)")
print(f"Цена Оптимум:      {opt_price} ₽")

if opt_price:
    res = calc(purchase, opt_price, logistics)
    print(f"\nРАСКЛАД ПРИ ЦЕНЕ {opt_price} ₽:")
    print(f"  Комиссия Ozon ({fbs_rate(opt_price)*100:.0f}%):  {res['comm']:.0f} ₽")
    print(f"  Эквайринг (1.5%):     {res['acq']:.0f} ₽")
    print(f"  Логистика FBS:        {logistics:.0f} ₽")
    print(f"  Возвраты (3%):        {res['ret_loss']:.0f} ₽")
    print(f"  Упаковка:             {OTHER} ₽")
    print(f"  Налог УСН 6%:         {res['tax']:.0f} ₽")
    print(f"  Итого расходов:       {res['total_cost']:.0f} ₽")
    print(f"  ─────────────────────────────")
    print(f"  ПРИБЫЛЬ:              {res['profit']:.0f} ₽")
    print(f"  Наценка:              {res['markup']:.1f}%")
    print(f"  (наценка на закупка+упаковка = {purchase+OTHER:.0f} ₽)")
