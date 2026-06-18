"""
Проверяем: сколько WB-названий содержат OEM-подобные коды,
и как хорошо они совпадают с Автолигой.
"""
import sqlite3, re, xlrd
from pathlib import Path
from collections import defaultdict

DB_PATH       = Path("data/analytics/wb_index.db")
AUTOLIGA_PATH = Path(r"C:\Users\Admin\Desktop\PriceALVLG4.xls")

# ─── Паттерны OEM-кодов в названиях WB ──────────────────────────────────────
OEM_PATTERNS = [
    # Числовой ВАЗ-стиль: 2108-3703-010-05 или 21080370301005
    re.compile(r"\b\d{4}[-\s]?\d{3,4}[-\s]?\d{2,4}[-\s]?\d{1,4}\b"),
    # Длинный числовой без разделителей: 21080370301005
    re.compile(r"\b\d{10,16}\b"),
    # Буквенно-цифровые коды: AH015702, W71273, 1K0513029JA
    re.compile(r"\b[A-Z]{1,4}\d{3,}[A-Z0-9]*\b"),
    re.compile(r"\b[A-Z0-9]{2,}[-/][A-Z0-9]{2,}(?:[-/][A-Z0-9]+)*\b"),
]


def extract_oem_candidates(text: str) -> list[str]:
    text = text.upper()
    found = set()
    for pat in OEM_PATTERNS:
        for m in pat.finditer(text):
            code = re.sub(r"[^A-Z0-9]", "", m.group())
            if 4 <= len(code) <= 20:
                found.add(code)
    return list(found)


def normalize(s: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", s.upper())


# ─── Загружаем OEM Автолиги ──────────────────────────────────────────────────
print("Загружаем Автолигу...")
al_oems: set[str] = set()
wb_al = xlrd.open_workbook(str(AUTOLIGA_PATH), encoding_override="cp1251")
ws = wb_al.sheet_by_index(0)
for r in range(9, ws.nrows):
    row = ws.row_values(r)
    if len(row) < 8:
        continue
    oem = str(row[2]).strip()
    if oem.endswith(".0"):
        oem = oem[:-2]
    norm = normalize(oem)
    if len(norm) >= 4:
        al_oems.add(norm)
print(f"  Уникальных OEM в Автолиге: {len(al_oems)}")

# ─── Извлекаем OEM из названий WB ───────────────────────────────────────────
print("Анализируем WB названия...")
conn = sqlite3.connect(str(DB_PATH))
products = conn.execute("SELECT nm_id, name, brand FROM wb_products").fetchall()

matched_nm = 0
total_codes = 0
oem_to_nms: dict[str, list[int]] = defaultdict(list)

for nm_id, name, brand in products:
    if not name:
        continue
    candidates = extract_oem_candidates(name)
    total_codes += len(candidates)
    for code in candidates:
        oem_to_nms[code].append(nm_id)
        if code in al_oems:
            matched_nm += 1

matches_in_al = sum(1 for code in oem_to_nms if code in al_oems)
print(f"  Всего WB товаров: {len(products)}")
print(f"  Уникальных OEM-кодов извлечено: {len(oem_to_nms)}")
print(f"  Из них совпадают с Автолигой: {matches_in_al}")
print(f"  WB товаров с совпадением: {matched_nm}")

# ─── Примеры совпадений ──────────────────────────────────────────────────────
print("\nПримеры WB названий с OEM из Автолиги:")
shown = 0
for nm_id, name, brand in products:
    if not name or shown >= 15:
        break
    for code in extract_oem_candidates(name):
        if code in al_oems:
            print(f"  nm={nm_id}  [{code}]  {brand} | {name[:60]}")
            shown += 1
            break

conn.close()
