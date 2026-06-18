import sys, openpyxl
sys.stdout.reconfigure(encoding="utf-8")
wb = openpyxl.load_workbook(r"C:\Users\Admin\Desktop\Топ ВБ 1306\Топ-500 ВБ 1406.xlsx", read_only=True)
ws = wb.active
# Ищем строки, где в изображениях есть наши пропавшие префиксы
missing_prefixes = ("ag01", "vsp0", "lecar", "brr", "brb", "fap", "crl", "amdfa", "amdfl", "sb20", "sb21", "sb31", "sb32", "sb33", "атс", "бво", "вв", "щдр")
for r in range(2, ws.max_row + 1):
    img = ws.cell(r, 15).value
    art = ws.cell(r, 1).value
    sup = ws.cell(r, 2).value if ws.max_column >= 2 else ""
    if img:
        img_low = str(img).lower()
        for pref in missing_prefixes:
            if pref in img_low:
                print(f"Арт={art}  Пост={sup}  Фото={img[:60]}")
                break
wb.close()
