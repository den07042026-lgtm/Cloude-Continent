import sys, openpyxl
sys.stdout.reconfigure(encoding="utf-8")

wb = openpyxl.load_workbook(r"C:\Users\Admin\Desktop\Топ ВБ 1306\Топ-500 ВБ 1306.xlsx", read_only=True)
ws = wb.active

print("=== СТРОКИ С ОПИСАНИЕМ (не пустое col10) ===")
count = 0
for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
    descr = row[9] if len(row) > 9 else None
    if descr and str(descr).strip():
        name = row[1] or ""
        supplier = row[4] or ""
        print(f"Строка {i}: [{supplier}] {str(name)[:50]}")
        print(f"  Описание: {str(descr)[:300]}")
        print()
        count += 1
        if count >= 6:
            break

print("\n=== СТРОКИ MIKADO БЕЗ ОПИСАНИЯ (target rows) ===")
for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
    supplier = row[4] if len(row) > 4 else None
    name = row[1] if len(row) > 1 else None
    descr = row[9] if len(row) > 9 else None
    if (supplier and str(supplier).strip().lower() == "mikado"
        and name and str(name).strip()
        and (not descr or not str(descr).strip())):
        print(f"Строка {i}: {str(name)[:60]}")
        print(f"  params={str(row[5] or '')[:120]}")
        print(f"  oem   ={str(row[6] or '')[:100]}")
        print(f"  compat={str(row[7] or '')[:100]}")
        print(f"  alts  ={str(row[8] or '')[:100]}")
        print()
wb.close()
