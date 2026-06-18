import sys, openpyxl
sys.stdout.reconfigure(encoding="utf-8")

wb = openpyxl.load_workbook(r"C:\Users\Admin\Desktop\Топ ВБ 1306\Топ-500 ВБ 1306.xlsx", read_only=True)
ws = wb.active

# Показываем заголовки
headers = [ws.cell(1, c).value for c in range(1, 20)]
print("Заголовки:", headers)

print("\n=== Примеры строк со значением в столбце 15 (Изображения) ===")
count = 0
for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
    kod  = row[0] if len(row) > 0 else None
    img  = row[14] if len(row) > 14 else None
    if img and str(img).strip():
        print(f"  [{i}] Код={kod}  Изображения={str(img)[:100]}")
        count += 1
        if count >= 5:
            break

print("\n=== Первые 20 строк с пустым столбцом 15 ===")
count = 0
for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
    kod  = row[0] if len(row) > 0 else None
    name = row[1] if len(row) > 1 else None
    img  = row[14] if len(row) > 14 else None
    if kod and (not img or not str(img).strip()):
        print(f"  [{i}] Код={kod}  Наим={str(name or '')[:50]}")
        count += 1
        if count >= 20:
            break

wb.close()
