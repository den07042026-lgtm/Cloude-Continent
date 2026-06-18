import sys, io, openpyxl
from collections import Counter
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

wb = openpyxl.load_workbook('data/analytics/wb_top500_combined.xlsx')
ws = wb.worksheets[0]

categories = Counter()
for row in ws.iter_rows(min_row=2, max_row=501, values_only=True):
    name = str(row[3]).lower() if row[3] else ''
    if 'фильтр' in name or 'filter' in name:
        if 'масл' in name or 'oil' in name:
            categories['Фильтр масляный'] += 1
        elif 'воздуш' in name or 'air' in name:
            categories['Фильтр воздушный'] += 1
        elif 'салон' in name or 'cabin' in name:
            categories['Фильтр салонный'] += 1
        elif 'топлив' in name or 'fuel' in name:
            categories['Фильтр топливный'] += 1
        else:
            categories['Фильтр прочий'] += 1
    elif 'амортиз' in name or 'shock' in name:
        categories['Амортизатор'] += 1
    elif 'колодк' in name:
        categories['Колодки тормозные'] += 1
    elif 'тормозной диск' in name or 'brake disc' in name:
        categories['Диск тормозной'] += 1
    elif 'сайлентблок' in name or 'silent' in name:
        categories['Сайлентблок'] += 1
    elif 'шаровая' in name:
        categories['Шаровая опора'] += 1
    elif 'стойк' in name:
        categories['Стойка стабилизатора'] += 1
    elif 'рычаг' in name:
        categories['Рычаг подвески'] += 1
    elif 'масло' in name:
        categories['Масло'] += 1
    elif 'антифриз' in name or 'coolant' in name:
        categories['Антифриз'] += 1
    elif 'свеч' in name or 'spark' in name:
        categories['Свечи зажигания'] += 1
    elif 'ремень' in name or 'belt' in name:
        categories['Ремень'] += 1
    elif 'подшипник' in name:
        categories['Подшипник'] += 1
    elif 'наконечник' in name:
        categories['Рулевой наконечник'] += 1
    elif 'дворник' in name or 'wiper' in name:
        categories['Дворники'] += 1
    elif 'катушк' in name:
        categories['Катушка зажигания'] += 1
    elif 'термостат' in name:
        categories['Термостат'] += 1
    elif 'помпа' in name or 'насос' in name:
        categories['Помпа/насос'] += 1
    else:
        categories['Прочее'] += 1

total_filters = sum(v for k, v in categories.items() if 'Фильтр' in k)
print('Категории в Топ-500:')
for cat, n in categories.most_common():
    bar = '#' * (n // 3)
    print(f'  {cat:<28} {n:>3}  {bar}')
print(f'\n  Итого фильтров: {total_filters} из 500 ({total_filters/5:.0f}%)')
print(f'  Итого подвеска: {categories["Амортизатор"]+categories["Сайлентблок"]+categories["Шаровая опора"]+categories["Стойка стабилизатора"]+categories["Рычаг подвески"]}')
