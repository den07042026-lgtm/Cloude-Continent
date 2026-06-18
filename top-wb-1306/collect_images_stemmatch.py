import sys, shutil, re, openpyxl
from pathlib import Path
sys.stdout.reconfigure(encoding="utf-8")

XLSX = Path(r"C:\Users\Admin\Desktop\Топ ВБ 1406\Топ-500 ВБ 1406.xlsx")
DST  = Path(r"C:\Users\Admin\Desktop\Топ ВБ 1406\Изображения")
IMG_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".jfif"}

# ── 1. Читаем нужные файлы из Excel ──────────────────────────────────────────
wb = openpyxl.load_workbook(XLSX, read_only=True)
ws = wb.active
needed = set()
for r in range(2, ws.max_row + 1):
    v = ws.cell(r, 15).value
    if v:
        for f in str(v).replace(";", ",").split(","):
            f = f.strip()
            if f: needed.add(f.lower())
wb.close()

DST.mkdir(parents=True, exist_ok=True)
existing = {f.name.lower() for f in DST.iterdir() if f.is_file()}
still_needed = {n for n in needed if n not in existing}
print(f"Уже есть: {len(needed) - len(still_needed)}, ищем: {len(still_needed)}")

if not still_needed:
    print("Всё найдено!")
    sys.exit(0)

# Стемы нужных файлов: убираем _N и расширение
def get_stem(name: str) -> str:
    s = Path(name).stem.lower()
    return re.sub(r"_\d+$", "", s)

needed_stems = {get_stem(n): n for n in still_needed}  # stem → исходное имя
print(f"Уникальных стемов для поиска: {len(needed_stems)}")

# ── 2. Полный скан без ограничений ───────────────────────────────────────────
SCAN_ROOTS = ["C:\\", "D:\\"]
SKIP_PARTS = {"windows", "$recycle.bin", "system volume information"}

print(f"\nПолный скан (стем-совпадение)...")

exact_idx = {}   # точное имя → Path
stem_idx  = {}   # стем → Path

scanned = 0
for root_str in SCAN_ROOTS:
    root = Path(root_str)
    if not root.exists():
        continue
    for f in root.rglob("*"):
        try:
            if not f.is_file():
                continue
            parts_lower = {p.lower() for p in f.parts}
            if parts_lower & SKIP_PARTS:
                continue
            if f.suffix.lower() not in IMG_EXTS:
                continue
            scanned += 1
            key = f.name.lower()
            stem = get_stem(key)
            if key in still_needed and key not in exact_idx:
                exact_idx[key] = f
            if stem in needed_stems and stem not in stem_idx:
                stem_idx[stem] = f
            if scanned % 100_000 == 0:
                print(f"  {scanned:,} файлов... точных: {len(exact_idx)}, по стему: {len(stem_idx)}")
        except (PermissionError, OSError):
            continue

print(f"\nПросканировано: {scanned:,} файлов")
print(f"Точных совпадений: {len(exact_idx)}")
print(f"По стему: {len(stem_idx)}")

# ── 3. Копируем ───────────────────────────────────────────────────────────────
copied = 0
missing = []

for name in sorted(still_needed):
    dst_file = DST / name
    if dst_file.exists():
        copied += 1
        continue
    if name in exact_idx:
        shutil.copy2(exact_idx[name], dst_file)
        print(f"  ТОЧНЫЙ: {name}  ←  {exact_idx[name]}")
        copied += 1
    else:
        stem = get_stem(name)
        if stem in stem_idx:
            shutil.copy2(stem_idx[stem], dst_file)
            print(f"  СТЕМ:   {name}  ←  {stem_idx[stem].name}")
            copied += 1
        else:
            missing.append(name)

print(f"\nДополнительно скопировано: {copied}")
print(f"Итого в папке: {len(existing) + copied}")
print(f"Не найдено нигде: {len(missing)}")
for m in sorted(missing):
    print(f"  {m}")
