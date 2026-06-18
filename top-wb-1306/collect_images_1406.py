import sys, shutil, openpyxl
from pathlib import Path
sys.stdout.reconfigure(encoding="utf-8")

XLSX = Path(r"C:\Users\Admin\Desktop\Топ ВБ 1306\Топ-500 ВБ 1406.xlsx")
DST  = Path(r"C:\Users\Admin\Desktop\Топ ВБ 1406\Изображения")

SEARCH_DIRS = [
    r"C:\Users\Admin\Desktop\Топ ВБ 1306\Изображения Автопитер",
    r"C:\Users\Admin\Desktop\Топ ВБ 1306\images",
    r"C:\Users\Admin\Desktop\Топ-500 ВБ\Изображения Автопитер",
    r"C:\Users\Admin\Desktop\Топ-500 ВБ\Изображения Автолига",
    r"C:\Users\Admin\Desktop\Топ-500 ВБ\Изображения Микадо",
    r"C:\Users\Admin\Desktop\Топ-500 ВБ\СТ изображения",
    r"C:\Users\Admin\Desktop\На сортировку 08.05\images",
    r"C:\Users\Admin\Desktop\На сортировку 24.04(2)\images",
    r"D:\Изображения Автопитер",
    r"C:\Users\Admin\Documents\Autoparts_Ecommerce",
    r"C:\Users\Admin\Downloads",
    r"C:\Users\Admin\Yandex.Disk\Озон",                              # рекурсивно — 53 447 фото
    r"C:\Users\Admin\Yandex.Disk-SharedResources-kalugin.don",
]

IMG_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".jfif"}

# ── 1. Читаем нужные имена файлов из Excel ───────────────────────────────────
print("Читаю Excel...")
wb = openpyxl.load_workbook(XLSX, read_only=True)
ws = wb.active
needed = set()
for r in range(2, ws.max_row + 1):
    v = ws.cell(r, 15).value
    if v:
        raw = str(v).replace(";", ",")
        for f in raw.split(","):
            f = f.strip()
            if f:
                needed.add(f.lower())
wb.close()
print(f"Нужно файлов: {len(needed)}")

# ── 2. Индексируем все файлы из папок-источников ────────────────────────────
print("\nИндексирую источники...")
index = {}   # имя_нижний_регистр → Path
for d in SEARCH_DIRS:
    p = Path(d)
    if not p.exists():
        print(f"  Пропуск (нет папки): {d}")
        continue
    cnt = 0
    for f in p.rglob("*"):
        if f.is_file() and f.suffix.lower() in IMG_EXTS:
            key = f.name.lower()
            if key not in index:   # первое вхождение приоритетно
                index[key] = f
            cnt += 1
    print(f"  {p.name}: {cnt} файлов")

print(f"Всего проиндексировано: {len(index)} уникальных имён")

# ── 3. Копируем совпадения ───────────────────────────────────────────────────
DST.mkdir(parents=True, exist_ok=True)
print(f"\nКопирую в: {DST}")

copied  = 0
missing = []

for name in sorted(needed):
    if name in index:
        src_file = index[name]
        dst_file = DST / src_file.name
        if not dst_file.exists():
            shutil.copy2(src_file, dst_file)
            copied += 1
        else:
            copied += 1   # уже есть
    else:
        missing.append(name)

print(f"\nСкопировано: {copied}/{len(needed)}")
if missing:
    print(f"Не найдено: {len(missing)}")
    for m in sorted(missing):
        print(f"  {m}")
else:
    print("Все файлы найдены!")
