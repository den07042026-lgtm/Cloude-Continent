import sys, openpyxl
sys.stdout.reconfigure(encoding="utf-8")
wb = openpyxl.load_workbook(r"C:\Users\Admin\Desktop\Топ ВБ 1306\Топ-500 ВБ 1306.xlsx", read_only=True)
ws = wb.active
print("Первые 30 непустых наименований:")
count = 0
for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
    name = row[1] if len(row) > 1 else None
    if name and str(name).strip():
        print(f"  [{i}] {str(name)}")
        count += 1
        if count >= 30:
            break
wb.close()
