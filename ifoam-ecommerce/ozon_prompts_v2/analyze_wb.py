import pandas as pd

f = 'C:/Users/1/Downloads/DataExport_2026-03-24_09-37-33.xlsx'
df = pd.read_excel(f, sheet_name='Chart data', header=0)
df.columns = ['Период','Маркетплейс','Продано','Маржа_pct','Прибыль','Приход',
              'Продажи_gross','Скидки','Себест_total','Комиссия_mp','C10','C11','C12',
              'Комиссия2','Хранение','Доставка','C16','C17','C18','Себест_за_шт']

ozon = df[df['Маркетплейс']=='Ozon'].copy().sort_values('Период').reset_index(drop=True)
wb   = df[df['Маркетплейс']=='Wildberries'].copy().sort_values('Период').reset_index(drop=True)

def per_unit(val, sold):
    return round(val / sold) if sold and sold > 0 else 0

def enrich(d):
    d['Ср_цена']    = d.apply(lambda r: per_unit(r['Продажи_gross'], r['Продано']), axis=1)
    d['Прибыль_шт'] = d.apply(lambda r: per_unit(r['Прибыль'], r['Продано']), axis=1)
    d['Комис_pct']  = d.apply(
        lambda r: round(abs(r['Комиссия_mp'])/r['Продажи_gross']*100, 1) if r['Продано']>0 else 0, axis=1)
    return d

ozon = enrich(ozon)
wb   = enrich(wb)

out = []
out.append('=== СРАВНЕНИЕ OZON vs WILDBERRIES: Orange Standard 5кг ===')
out.append(f'{"Период":<12} {"WB шт":>6} {"WB цена":>8} {"WB марж%":>9} {"WB приб/шт":>11} | {"OZ шт":>6} {"OZ цена":>8} {"OZ марж%":>9} {"OZ приб/шт":>11}')
out.append('-'*90)

periods = sorted(set(list(ozon['Период']) + list(wb['Период'])))
for p in periods:
    wb_r  = wb[wb['Период']==p]
    oz_r  = ozon[ozon['Период']==p]

    if not wb_r.empty:
        r = wb_r.iloc[0]
        wb_sold  = int(r['Продано'])
        wb_price = int(r['Ср_цена']) if wb_sold > 0 else 0
        wb_marg  = f"{r['Маржа_pct']*100:.1f}%" if wb_sold > 0 else '-'
        wb_prof  = f"{int(r['Прибыль_шт']):+,}" if wb_sold > 0 else '-'
    else:
        wb_sold = 0; wb_price = 0; wb_marg = '-'; wb_prof = '-'

    if not oz_r.empty:
        r = oz_r.iloc[0]
        oz_sold  = int(r['Продано'])
        oz_price = int(r['Ср_цена']) if oz_sold > 0 else 0
        oz_marg  = f"{r['Маржа_pct']*100:.1f}%" if oz_sold > 0 else '-'
        oz_prof  = f"{int(r['Прибыль_шт']):+,}" if oz_sold > 0 else '-'
    else:
        oz_sold = 0; oz_price = 0; oz_marg = '-'; oz_prof = '-'

    period_s = str(p)[:10]
    out.append(f'{period_s:<12} {wb_sold:>6} {wb_price:>8,} {wb_marg:>9} {wb_prof:>11} | '
               f'{oz_sold:>6} {oz_price:>8,} {oz_marg:>9} {oz_prof:>11}')

# Итоги
out.append('-'*90)
out.append(f'{"ИТОГО WB":<12} {int(wb["Продано"].sum()):>6} {int(wb["Продажи_gross"].sum()/wb["Продано"].sum()):>8,} '
           f'{"":>9} {int(wb["Прибыль"].sum()/wb["Продано"].sum()):>+11,} | '
           f'{int(ozon["Продано"].sum()):>6} {int(ozon["Продажи_gross"].sum()/ozon["Продано"].sum()):>8,} '
           f'{"":>9} {int(ozon["Прибыль"].sum()/ozon["Продано"].sum()):>+11,}')

out.append('')
out.append('=== КЛЮЧЕВЫЕ ВЫВОДЫ ===')
out.append(f'WB: {int(wb["Продано"].sum())} шт за {len(wb)} мес = {wb["Продано"].mean():.0f} шт/мес avg')
out.append(f'OZ: {int(ozon["Продано"].sum())} шт за {len(ozon)} мес = {ozon["Продано"].mean():.1f} шт/мес avg')
out.append(f'WB объём в {wb["Продано"].mean()/ozon["Продано"].mean():.0f}x больше чем Ozon')
wb_sell = wb[wb['Продано']>0]
out.append(f'WB цена: {int(wb_sell["Ср_цена"].min()):,}–{int(wb_sell["Ср_цена"].max()):,} руб (ср. {int(wb_sell["Ср_цена"].mean()):,})')
oz_sell = ozon[ozon['Продано']>0]
out.append(f'OZ цена: {int(oz_sell["Ср_цена"].min()):,}–{int(oz_sell["Ср_цена"].max()):,} руб (ср. {int(oz_sell["Ср_цена"].mean()):,})')
best_wb = wb_sell.loc[wb_sell['Прибыль_шт'].idxmax()]
out.append(f'Лучший мес WB: {str(best_wb["Период"])[:10]} @ {int(best_wb["Ср_цена"]):,} руб -> '
           f'{int(best_wb["Продано"])} шт, маржа {best_wb["Маржа_pct"]*100:.1f}%, прибыль {int(best_wb["Прибыль_шт"]):+,}/шт')
worst_wb = wb_sell.loc[wb_sell['Прибыль_шт'].idxmin()]
out.append(f'Худший мес WB: {str(worst_wb["Период"])[:10]} @ {int(worst_wb["Ср_цена"]):,} руб -> '
           f'{int(worst_wb["Продано"])} шт, маржа {worst_wb["Маржа_pct"]*100:.1f}%, прибыль {int(worst_wb["Прибыль_шт"]):+,}/шт')

text = '\n'.join(out)
with open('C:/Users/1/Downloads/wb_vs_ozon.txt', 'w', encoding='utf-8') as fout:
    fout.write(text)
print('OK')
