import sys, openpyxl
sys.stdout.reconfigure(encoding="utf-8")

for fname in ["Топ-500 ВБ 1306.xlsx", "Топ-500 ВБ 1306_BACKUP.xlsx"]:
    wb = openpyxl.load_workbook(rf"C:\Users\Admin\Desktop\Топ ВБ 1306\{fname}")
    ws = wb.active
    rd = ws.row_dimensions[1]
    print(f"=== {fname} ===")
    print(f"  Строка 1 — высота: {rd.height}, скрыта: {rd.hidden}")
    print(f"  freeze_panes: {ws.freeze_panes}")
    # Проверим высоту первых 5 строк
    for r in range(1, 6):
        rr = ws.row_dimensions[r]
        print(f"  row{r}: height={rr.height}, hidden={rr.hidden}")
    wb.close()
    print()
