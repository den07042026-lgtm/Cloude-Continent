"""Dry-run обогатителя — показывает что будет заполнено без изменения Excel."""
import sys, io, os, openpyxl
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, os.path.abspath('scripts'))
import emex_enricher as enr

EXCEL_FILE = r'C:\Users\Admin\Desktop\Топ-500 ВБ\Топ-500 ВБ.xlsx'
wb = openpyxl.load_workbook(EXCEL_FILE)
ws = wb.worksheets[0]

print('row | article              | brand         | NEW_applic                         | emex_makes')
print('-' * 120)

for row_idx in range(2, 16):  # первые 14 строк
    article  = ws.cell(row_idx, 1).value
    name     = ws.cell(row_idx, 2).value
    brand    = ws.cell(row_idx, 3).value
    cur_app  = ws.cell(row_idx, 8).value
    cur_alts = ws.cell(row_idx, 9).value

    if not article:
        continue

    article_s = str(article).strip()
    brand_s   = str(brand).strip() if brand else ''
    name_s    = str(name).strip() if name else ''

    # Применяемость
    extracted = enr.extract_applicability_from_name(name_s)
    if extracted and (not cur_app or len(str(cur_app)) < len(extracted)):
        new_applic = f'NEW: {extracted[:50]}'
    else:
        new_applic = f'keep: {str(cur_app or "")[:50]}'

    # Альтернативные (только если нет и есть бренд)
    if not cur_alts and brand_s:
        data  = enr.fetch_emex_data(article_s, brand_s)
        makes = data.get('makes', [])
        alts  = enr.format_alternatives(makes, brand_s, article_s)
        makes_cnt = len(makes)
        err_note = f'  ERR:{data.get("error","")[:20]}' if 'error' in data else ''
    else:
        alts = str(cur_alts or '')[:50]
        makes_cnt = 0
        err_note = '(already)'

    print(f'{row_idx-1:3d} | {article_s[:20]:20s} | {brand_s[:13]:13s} | {new_applic[:50]:50s} | {makes_cnt:2d} {alts[:40]}{err_note}')
