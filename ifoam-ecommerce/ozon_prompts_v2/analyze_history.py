import pandas as pd

f = 'C:/Users/1/Downloads/DataExport_2026-03-24_09-04-22.xlsx'
df = pd.read_excel(f, sheet_name='Chart data', header=0)

# Переименуем по индексам
df.columns = [
    'Период',          # 0
    'Продано',         # 1
    'Маржа_pct',       # 2
    'Прибыль',         # 3
    'Приход',          # 4
    'Продажи_gross',   # 5
    'Скидки_баллы',    # 6
    'Себест_total',    # 7
    'Комиссия_ozon',   # 8
    'Col9',            # 9
    'Col10',           # 10
    'Col11',           # 11
    'Комиссия2',       # 12
    'Хранение',        # 13
    'Доставка',        # 14
    'Col15',           # 15
    'Col16',           # 16
    'Col17',           # 17
    'Себест_за_шт',    # 18
]

total_row = df[df['Период'].str.startswith('За')].iloc[0]
months = df[~df['Период'].str.startswith('За')].copy().sort_values('Период').reset_index(drop=True)

def per_unit(val, sold):
    return round(val / sold) if sold > 0 else 0

months['Ср_цена']       = months.apply(lambda r: per_unit(r['Продажи_gross'], r['Продано']), axis=1)
months['Приход_шт']     = months.apply(lambda r: per_unit(r['Приход'],         r['Продано']), axis=1)
months['Прибыль_шт']    = months.apply(lambda r: per_unit(r['Прибыль'],        r['Продано']), axis=1)
months['Комис_шт']      = months.apply(lambda r: per_unit(abs(r['Комиссия_ozon']), r['Продано']), axis=1)
months['Комис_pct']     = months.apply(
    lambda r: round(abs(r['Комиссия_ozon']) / r['Продажи_gross'] * 100, 1) if r['Продано'] > 0 else 0, axis=1)

out = []
out.append('=== ИСТОРИЯ ПРОДАЖ: Orange Standard 5кг (Art. 119105) ===')
out.append(f'{"Период":<27} {"Шт":>4} {"Цена":>7} {"Комис%":>7} {"Хранение":>9} {"Прибыль/шт":>12} {"Маржа%":>8}')
out.append('-'*80)

for _, r in months.iterrows():
    period = str(r['Период'])[:24]
    sold = int(r['Продано'])
    stor = f"{float(r['Хранение']):.0f}"
    if sold > 0:
        price  = f"{int(r['Ср_цена']):,}"
        comm   = f"{r['Комис_pct']:.0f}%"
        prof   = f"{int(r['Прибыль_шт']):+,}"
        margin = f"{r['Маржа_pct']*100:.1f}%"
        note   = ' <- убыток' if r['Прибыль_шт'] < 0 else ''
    else:
        price = comm = prof = '-'
        margin = '0%'
        note = ' <- только хранение'
    out.append(f'{period:<27} {sold:>4} {price:>7} {comm:>7} {stor:>9} {prof:>12} {margin:>8}{note}')

out.append('-'*80)
tot_sold = int(total_row['Продано'])
out.append(f'ИТОГО: {tot_sold} шт | Выручка {int(total_row["Продажи_gross"]):,} | Приход {int(total_row["Приход"]):,} | Прибыль {int(total_row["Прибыль"]):,}')
out.append(f'Средняя цена: {int(total_row["Продажи_gross"]/tot_sold):,} руб/шт | '
           f'Среднемесячно: {tot_sold/14:.1f} шт/мес')

# Зависимость цены и объёма
out.append('')
out.append('=== ЗАВИСИМОСТЬ: ЦЕНА -> ОБЪЁМ -> ПРИБЫЛЬ/ШТ ===')
selling = months[months['Продано'] > 0].sort_values('Ср_цена')
for _, r in selling.iterrows():
    period = str(r['Период'])[:10]
    bar = '█' * int(r['Продано'])
    sign = '+' if r['Прибыль_шт'] >= 0 else ''
    out.append(f'  {period}: {int(r["Ср_цена"]):>5,} руб -> {int(r["Продано"])} шт {bar:<5} '
               f'прибыль {sign}{int(r["Прибыль_шт"]):,} руб/шт  комиссия {r["Комис_pct"]:.0f}%')

# Выводы
out.append('')
out.append('=== КЛЮЧЕВЫЕ ВЫВОДЫ ===')
best_vol    = selling.loc[selling['Продано'].idxmax()]
best_margin = selling.loc[selling['Прибыль_шт'].idxmax()]
loss_months = selling[selling['Прибыль_шт'] < 0]
no_sales    = len(months[months['Продано'] == 0])

out.append(f'Лучший месяц по объёму:  {str(best_vol["Период"])[:10]} — {int(best_vol["Продано"])} шт '
           f'@ {int(best_vol["Ср_цена"]):,} руб -> прибыль {int(best_vol["Прибыль_шт"]):+,} руб/шт')
out.append(f'Лучший месяц по марже:   {str(best_margin["Период"])[:10]} — '
           f'{int(best_margin["Прибыль_шт"]):+,} руб/шт @ {int(best_margin["Ср_цена"]):,} руб')
if not loss_months.empty:
    out.append(f'Убыточные продажи:       при цене {int(loss_months["Ср_цена"].min()):,}–{int(loss_months["Ср_цена"].max()):,} руб')
out.append(f'Месяцев без продаж:      {no_sales} из {len(months)} (только хранение — кровоточит)')
out.append(f'Себестоимость:           {int(total_row["Себест_за_шт"])} руб/шт (из данных)')

# Хранение суммарно
total_storage = months[months['Продано'] == 0]['Хранение'].sum()
out.append(f'Потери на хранение (пустые месяцы): {abs(total_storage):.0f} руб')

text = '\n'.join(out)
with open('C:/Users/1/Downloads/history_analysis.txt', 'w', encoding='utf-8') as fout:
    fout.write(text)
print(text)
