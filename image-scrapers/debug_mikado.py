import openpyxl
from pathlib import Path

EXCEL = r"C:\Users\Admin\Desktop\Топ-500 ВБ\Топ-500 ВБ.xlsx"
IMGS  = Path(r"C:\Users\Admin\Desktop\На сортировку 08.05\images")

wb = openpyxl.load_workbook(EXCEL, read_only=True)
ws = wb.active

articles = []
for row in ws.iter_rows(min_row=2, values_only=True):
    if row[4] and row[0] and "микадо" in str(row[4]).lower():
        articles.append(str(row[0]).strip().lower())
    if len(articles) >= 10:
        break
wb.close()

all_files = list(IMGS.glob("*"))
print(f"Artikuli Mikado ({len(articles)}):")
for art in articles:
    found = [f.name for f in all_files if art in f.stem.lower()]
    status = found[:3] if found else ["NOT FOUND"]
    print(f"  [{art}] -> {status}")
