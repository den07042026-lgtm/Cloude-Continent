import sys, zipfile, shutil, re, os
sys.stdout.reconfigure(encoding="utf-8")

src  = r"C:\Users\Admin\Desktop\Топ ВБ 1306\Топ-500 ВБ 1306.xlsx"
tmp  = src + ".tmp.xlsx"

with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)

        if item.filename == "xl/worksheets/sheet1.xml":
            text = data.decode("utf-8")

            # Показываем ДО
            sv_before = re.search(r'<sheetViews>.*?</sheetViews>', text, re.DOTALL)
            print("ДО:")
            print(sv_before.group(0) if sv_before else "не найдено")

            # Сбрасываем горизонтальный сдвиг листа: topLeftCell="H1" → "A1"
            # и сдвиг в pane: оставляем freeze на строке 1 (A2)
            text = re.sub(
                r'<sheetView\b([^>]*?)topLeftCell="[^"]*"([^>]*)>',
                lambda m: f'<sheetView{m.group(1)}topLeftCell="A1"{m.group(2)}>',
                text
            )
            # Сбрасываем topLeftCell внутри <pane> на A2 (начало прокручиваемой зоны)
            text = re.sub(
                r'(<pane\b[^>]*?)topLeftCell="[^"]*"',
                r'\1topLeftCell="A2"',
                text
            )

            sv_after = re.search(r'<sheetViews>.*?</sheetViews>', text, re.DOTALL)
            print("\nПОСЛЕ:")
            print(sv_after.group(0) if sv_after else "не найдено")

            data = text.encode("utf-8")

        zout.writestr(item, data)

# Заменяем оригинал
os.replace(tmp, src)
print("\nФайл сохранён!")
