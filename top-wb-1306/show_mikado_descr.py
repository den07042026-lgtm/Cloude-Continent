import sys, openpyxl
sys.stdout.reconfigure(encoding="utf-8")

wb = openpyxl.load_workbook(r"C:\Users\Admin\Desktop\Топ ВБ 1306\Топ-500 ВБ 1306.xlsx", read_only=True)
ws = wb.active
for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
    supplier = row[4] if len(row) > 4 else None
    name = row[1] if len(row) > 1 else None
    descr = row[9] if len(row) > 9 else None
    if supplier and str(supplier).strip().lower() == "mikado" and name and str(name).strip():
        print(f"Строка {i}: {str(name)[:65]}")
        print(f"  Описание: {str(descr or '(ПУСТО)')[:400]}")
        print()
wb.close()
