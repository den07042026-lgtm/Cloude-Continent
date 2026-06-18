"""
normalize_alts.py
Нормализует артикулы в столбце "Альтернативные артикулы товара":
  xnfl-cn1036 → cn1036
  xzk-if-3083k → if-3083k
Использует прайс Mikado как словарь Prodnum → Code.
Коды не из Mikado (без префикса) остаются без изменений.
"""
import sys, openpyxl
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8')

BASE_FILE  = Path(r"C:\Users\Admin\Desktop\Топ ВБ 1306\Топ-500 ВБ 1306.xlsx")
PRICE_FILE = Path(r"C:\Users\Admin\Documents\Ecommerce\mikado_price_34.xlsx")
ALT_COL    = 9   # столбец "Альтернативные артикулы товара" (1-based)

# ── Строим словарь Prodnum (lower) → Code ─────────────────────────────────────
print("Загружаю прайс Mikado...")
prodnum_to_code = {}
wb_p = openpyxl.load_workbook(PRICE_FILE, read_only=True, data_only=True)
ws_p = wb_p.active
first = next(ws_p.iter_rows(min_row=1, max_row=1, values_only=True))
headers = [str(v).strip() if v else "" for v in first]
pi = next((i for i, h in enumerate(headers) if h.lower() == "prodnum"), None)
ci = next((i for i, h in enumerate(headers) if h.lower() == "code"),    None)
if pi is None or ci is None:
    print("Колонки Prodnum/Code не найдены в прайсе!"); sys.exit(1)
for row in ws_p.iter_rows(min_row=2, values_only=True):
    p, c = row[pi], row[ci]
    if p and c:
        prodnum_to_code[str(p).strip().lower()] = str(c).strip()
wb_p.close()
print(f"  Загружено {len(prodnum_to_code)} пар")


def normalize(code: str) -> str:
    """Заменяет Prodnum на Code из прайса, иначе возвращает как есть."""
    return prodnum_to_code.get(code.strip().lower(), code.strip())


# ── Читаем и правим базовый файл ──────────────────────────────────────────────
print("Открываю базовый файл...")
wb = openpyxl.load_workbook(BASE_FILE)
ws = wb.active

changed_rows = 0
for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
    cell = ws.cell(row_idx, ALT_COL)
    val  = cell.value
    if not val:
        continue

    codes     = [c.strip() for c in str(val).split(";") if c.strip()]
    new_codes = [normalize(c) for c in codes]

    if new_codes != codes:
        cell.value = "; ".join(new_codes)
        changed_rows += 1
        print(f"  Строка {row_idx}: {'; '.join(codes)}")
        print(f"         → {'; '.join(new_codes)}")

print(f"\nИзменено строк: {changed_rows}")
wb.save(BASE_FILE)
wb.close()
print("Готово!")
