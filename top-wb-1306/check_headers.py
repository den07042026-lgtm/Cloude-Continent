import sys, os, openpyxl
sys.stdout.reconfigure(encoding="utf-8")
folder = r"C:\Users\Admin\Desktop\Топ ВБ 1306"

wb1 = openpyxl.load_workbook(folder + r"\Топ-500 ВБ 1306.xlsx", read_only=True)
ws1 = wb1.active
row1 = [ws1.cell(1, c).value for c in range(1, 20)]
print("ТЕКУЩИЙ файл, строка 1:")
for i, v in enumerate(row1, 1):
    print(f"  col{i}: {repr(v)}")
wb1.close()

print()
backup = None
for name in os.listdir(folder):
    if "BACKUP" in name.upper() or "backup" in name.lower():
        backup = os.path.join(folder, name)
        break
if backup:
    wb2 = openpyxl.load_workbook(backup, read_only=True)
    ws2 = wb2.active
    row2 = [ws2.cell(1, c).value for c in range(1, 20)]
    print(f"BACKUP ({os.path.basename(backup)}), строка 1:")
    for i, v in enumerate(row2, 1):
        print(f"  col{i}: {repr(v)}")
    wb2.close()
else:
    print("BACKUP файл не найден, список файлов в папке:")
    for name in os.listdir(folder):
        if name.endswith(".xlsx"):
            print(f"  {name}")
