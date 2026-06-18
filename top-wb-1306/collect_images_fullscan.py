import sys, shutil, openpyxl
from pathlib import Path
sys.stdout.reconfigure(encoding="utf-8")

XLSX = Path(r"C:\Users\Admin\Desktop\Топ ВБ 1406\Топ-500 ВБ 1406.xlsx")
DST  = Path(r"C:\Users\Admin\Desktop\Топ ВБ 1406\Изображения")
IMG_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".jfif"}

# ── 1. Читаем нужные файлы из Excel ──────────────────────────────────────────
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

# Уже скопированные пропускаем
DST.mkdir(parents=True, exist_ok=True)
existing = {f.name.lower() for f in DST.iterdir() if f.is_file()}
still_needed = {n for n in needed if n not in existing}
print(f"Уже в папке: {len(needed) - len(still_needed)}, ещё нужно найти: {len(still_needed)}")

if not still_needed:
    print("Все файлы уже скопированы!")
    sys.exit(0)

# ── 2. Полный скан всех дисков ────────────────────────────────────────────────
SCAN_ROOTS = ["C:\\", "D:\\"]
SKIP_DIRS = {
    "windows", "program files", "program files (x86)",
    "$recycle.bin", "system volume information",
    "programdata", "appdata\\local\\temp",
    ".git", "node_modules", "__pycache__",
}

print(f"\nПолный скан дисков: {SCAN_ROOTS}")
print("(может занять 2-5 минут...)")

index = {}
scanned = 0

for root_str in SCAN_ROOTS:
    root = Path(root_str)
    if not root.exists():
        continue
    for f in root.rglob("*"):
        try:
            if not f.is_file():
                continue
            # Пропускаем системные папки
            parts_lower = [p.lower() for p in f.parts]
            if any(skip in parts_lower for skip in SKIP_DIRS):
                continue
            if f.suffix.lower() not in IMG_EXTS:
                continue
            key = f.name.lower()
            if key in still_needed and key not in index:
                index[key] = f
            scanned += 1
            if scanned % 100000 == 0:
                print(f"  просканировано {scanned:,} файлов, найдено совпадений: {len(index)}")
        except (PermissionError, OSError):
            continue

print(f"\nВсего просканировано: {scanned:,} файлов")
print(f"Найдено совпадений:   {len(index)}")

# ── 3. Копируем ───────────────────────────────────────────────────────────────
copied = 0
missing = []

for name in sorted(still_needed):
    if name in index:
        dst_file = DST / index[name].name
        if not dst_file.exists():
            shutil.copy2(index[name], dst_file)
        copied += 1
        print(f"  СКОПИРОВАН: {name}  ←  {index[name]}")
    else:
        missing.append(name)

print(f"\nДополнительно скопировано: {copied}")
print(f"Итого в папке: {len(existing) + copied}")
print(f"Не найдено нигде: {len(missing)}")
if missing:
    for m in sorted(missing):
        print(f"  {m}")
