"""Обновляет краткие наименования в Топ-500 ВБ_new.xlsx."""
import sys, io, openpyxl
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

FILE = r"C:\Users\Admin\Desktop\Топ-500 ВБ\Топ-500 ВБ_new.xlsx"

# Словарь: артикул (lower) → расширенное наименование
NAMES = {
    # Фильтры воздушные Champion (строки 19-21)
    "caf100689p": "Фильтр воздушный CHAMPION Opel Astra G/H, Vectra B/C, Zafira A/B, Chevrolet Lacetti/Cruze",
    "caf100715p": "Фильтр воздушный CHAMPION VW Polo/Fox/Lupo, Skoda Fabia I/II, Seat Ibiza/Cordoba 1.0–1.4",
    "caf100724p": "Фильтр воздушный CHAMPION Renault Clio II/III/Modus, Dacia Logan/Sandero, Nissan Micra K12/Note",
    # Катушка (строка 22) — уже неплохое название, делаем чуть чище
    "217003705010": "Катушка зажигания Lada Priora ВАЗ-2170/2171/2172 1.6L 16-клапанный",
    # Фильтры салонные Champion (строки 24-31)
    "ccf0046":  "Фильтр салонный CHAMPION Renault Kangoo I (1998–2007), Renault Clio II (1998–2005)",
    "ccf0070":  "Фильтр салонный CHAMPION Opel Astra H (2004–2010), Zafira B (2005–2011), Corsa D",
    "ccf0093c": "Фильтр салонный угольный CHAMPION VW Touareg I (2004–2017), Audi Q7 4L (2007–2015), Porsche Cayenne (2003–2018)",
    "ccf0152":  "Фильтр салонный CHAMPION Ford Fiesta VI (2008–2017) 1.25–1.6/1.4TDCi/1.6TDCi",
    "ccf0153":  "Фильтр салонный CHAMPION Renault Koleos I (2008–2016) 2.0/2.5/2.0dCi",
    "ccf0327":  "Фильтр салонный CHAMPION Ford Mondeo III/IV (2004–2007), Jaguar X-Type (2005–2009)",
    "ccf0417":  "Фильтр салонный CHAMPION к-т 2 шт. VW Touareg I (2004–2010), Audi Q7 4L, Porsche Cayenne 955/957",
    # Свеча (строка 31)
    "oe059/t10": "Свеча зажигания медная CHAMPION L92YC — мотоциклы, ATV, садовая техника (M14×1.25, аналог NGK BP5HS)",
}

wb = openpyxl.load_workbook(FILE)
ws = wb.worksheets[0]

updated = 0
for row_idx in range(2, 502):
    article_raw = ws.cell(row_idx, 1).value
    if article_raw is None:
        continue
    article_key = str(article_raw).strip().lower()
    if article_key in NAMES:
        old = ws.cell(row_idx, 2).value
        ws.cell(row_idx, 2).value = NAMES[article_key]
        print(f"  Строка {row_idx-1:>3}: [{old}]")
        print(f"          -> [{NAMES[article_key]}]")
        updated += 1

wb.save(FILE)
print(f"\nОбновлено строк: {updated}")
