import sys, openpyxl
sys.stdout.reconfigure(encoding="utf-8")

path = r"C:\Users\Admin\Desktop\Топ ВБ 1306\Топ-500 ВБ 1306.xlsx"
wb = openpyxl.load_workbook(path)
ws = wb.active

print(f"Было freeze_panes: {ws.freeze_panes}")

# Стандартное закрепление: только строка заголовков (row 1) всегда видна
ws.freeze_panes = "A2"

print(f"Стало freeze_panes: {ws.freeze_panes}")

wb.save(path)
wb.close()
print("Готово!")
