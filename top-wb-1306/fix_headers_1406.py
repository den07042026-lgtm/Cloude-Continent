import sys, zipfile, re, os, openpyxl
from copy import copy
sys.stdout.reconfigure(encoding="utf-8")

folder  = r"C:\Users\Admin\Desktop\Топ ВБ 1306"
src     = folder + r"\Топ-500 ВБ 1306_BACKUP.xlsx"   # откуда берём заголовки
dst     = folder + r"\Топ-500 ВБ 1406.xlsx"           # куда вставляем

# ── 1. Копируем строку заголовков через openpyxl ─────────────────────────
wb_src = openpyxl.load_workbook(src)
ws_src = wb_src.active

wb_dst = openpyxl.load_workbook(dst)
ws_dst = wb_dst.active

for col in range(1, ws_src.max_column + 1):
    sc = ws_src.cell(1, col)
    dc = ws_dst.cell(1, col)
    dc.value = sc.value
    if sc.has_style:
        dc.font      = sc.font.copy()
        dc.fill      = sc.fill.copy()
        dc.border    = sc.border.copy()
        dc.alignment = sc.alignment.copy()
        dc.number_format = sc.number_format
    print(f"  col{col}: {sc.value!r}")

if ws_src.row_dimensions[1].height:
    ws_dst.row_dimensions[1].height = ws_src.row_dimensions[1].height

wb_dst.save(dst)
wb_dst.close()
wb_src.close()
print("Заголовки скопированы.\n")

# ── 2. Правим topLeftCell в XML ───────────────────────────────────────────
tmp = dst + ".tmp.xlsx"
with zipfile.ZipFile(dst, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "xl/worksheets/sheet1.xml":
            text = data.decode("utf-8")
            # Сбрасываем горизонтальный сдвиг
            text = re.sub(
                r'<sheetView\b([^>]*?)topLeftCell="[^"]*"([^>]*)>',
                lambda m: f'<sheetView{m.group(1)}topLeftCell="A1"{m.group(2)}>',
                text
            )
            # pane оставляем на A2 (freeze)
            text = re.sub(
                r'(<pane\b[^>]*?)topLeftCell="[^"]*"',
                r'\1topLeftCell="A2"',
                text
            )
            data = text.encode("utf-8")
        zout.writestr(item, data)

os.replace(tmp, dst)
print("XML-прокрутка сброшена в A1.")
print("Готово!")
