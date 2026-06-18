import sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")

path = r"C:\Users\Admin\Desktop\Топ ВБ 1306\Топ-500 ВБ 1306.xlsx"

with zipfile.ZipFile(path) as z:
    # Читаем XML листа
    with z.open("xl/worksheets/sheet1.xml") as f:
        content = f.read().decode("utf-8")

# Выводим sheetViews и первые rows
import re

# sheetViews блок
sv = re.search(r'<sheetViews>.*?</sheetViews>', content, re.DOTALL)
print("=== sheetViews ===")
print(sv.group(0) if sv else "НЕ НАЙДЕНО")

# Первые 3 строки данных
rows = re.findall(r'<row[^>]*>.*?</row>', content, re.DOTALL)
print(f"\n=== Первые 3 строки из {len(rows)} ===")
for r in rows[:3]:
    print(r[:300])
    print()
