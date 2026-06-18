import sys, openpyxl
sys.stdout.reconfigure(encoding="utf-8")

wb = openpyxl.load_workbook(r"C:\Users\Admin\Desktop\Топ ВБ 1306\Топ-500 ВБ 1306.xlsx", read_only=True)
ws = wb.active

print("Все примеры с несколькими изображениями:")
for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
    img = row[14] if len(row) > 14 else None
    if img and "," in str(img):
        print(f"  [{i}] {str(img)[:120]}")

print("\nПервые 10 заполненных ячеек столбца Изображения:")
count = 0
for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
    img = row[14] if len(row) > 14 else None
    if img and str(img).strip():
        print(f"  [{i}] {str(img)[:120]}")
        count += 1
        if count >= 10: break

wb.close()
