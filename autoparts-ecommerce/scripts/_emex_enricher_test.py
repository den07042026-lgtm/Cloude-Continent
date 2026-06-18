import sys, io, os, openpyxl
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, os.path.abspath('scripts'))
import emex_enricher as enr

EXCEL_FILE = r'C:\Users\Admin\Desktop\Топ-500 ВБ\Топ-500 ВБ.xlsx'
wb = openpyxl.load_workbook(EXCEL_FILE)
ws = wb.worksheets[0]

lines = []
lines.append('=== TEST extract_applicability_from_name ===')
for n in [
    "FILTР VOZDUSHNY VW PASSAT B8 14- TIGUAN 16- 2.0TDI (AMD) AMDFA576",
    "FILTР MASLYANY BMW E46 E39 E60 (AMD) AMDFL293",
    "Filtr vozdushny CHAMPION Ford Focus II/III Mazda 3/5 2004-2019",
    "KLYUCH SVECHNY 16MM TRUBCHATY",
    "SAYYENTBLOK VAZ-2101 DAAZ 8SHT",
]:
    r = enr.extract_applicability_from_name(n)
    lines.append(f'IN:  {n}')
    lines.append(f'OUT: {r}')
    lines.append('')

lines.append('=== TEST emex fetch (3 rows) ===')
for row_idx in range(2, 5):
    article = str(ws.cell(row_idx, 1).value).strip()
    brand   = str(ws.cell(row_idx, 3).value).strip()
    lines.append(f'#{row_idx-1}: {article} / {brand}')
    data  = enr.fetch_emex_data(article, brand)
    makes = data.get('makes', [])
    alts  = enr.format_alternatives(makes, brand, article)
    lines.append(f'  makes: {len(makes)}')
    lines.append(f'  alts:  {alts[:120]}')
    if 'error' in data:
        lines.append(f'  error: {data["error"]}')

print('\n'.join(lines))
