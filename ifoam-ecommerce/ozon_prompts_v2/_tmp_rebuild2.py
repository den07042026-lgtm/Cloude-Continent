# -*- coding: utf-8 -*-
import sys, pandas as pd
sys.path.insert(0, "C:/Users/1/Downloads/ozon_prompts_v2")
from save_helpers import get_product_folder, make_xlsx_competitors

folder = get_product_folder("ifoamHOME", "КондиционерBalmy_РайскиеЦветы", "770758")
bs_file = str(folder / "analytics_report_2026-03-26_15_49.xlsx")
print(f"Bestsellers: {bs_file}")

# Читаем: заголовки в строке 4 (индекс 4), пропускаем строку 5
df = pd.read_excel(bs_file, header=4, skiprows=[5])
COLS = {
    0: "Название", 1: "Ссылка", 2: "Продавец", 3: "Бренд",
    4: "Категория_1", 5: "Категория_3", 6: "Признак",
    7: "Заказано_руб", 8: "Динамика_%", 9: "Заказано_шт",
    10: "Средняя_цена", 11: "Мин_цена", 12: "Доля_выкупа_%",
    13: "Упущенные_продажи_руб", 14: "Дней_без_остатка",
    15: "Ср_доставка_часы", 16: "Ср_продажи_руб_день",
    17: "Ср_продажи_шт_день", 18: "Остаток_шт", 19: "Схема",
    20: "Объём_л", 21: "Показы_всего", 22: "Показы_поиск_каталог",
    23: "Просмотры_карточки", 24: "Конверсия_показ_заказ_%",
    25: "В_корзину_поиск_%", 26: "В_корзину_карточка_%",
    27: "Скидка_акции", 28: "Доля_оборота_акции_%",
    29: "Дней_в_акциях", 30: "Дней_с_продвижением",
    31: "ДРР_%", 32: "Дата_создания",
}
df.columns = [COLS.get(i, f"col_{i}") for i in range(len(df.columns))]
df = df[df["Название"].notna() & (df["Название"].astype(str) != "nan")]
print(f"Всего строк в файле: {len(df)}")

# Посмотрим уникальные категории чтобы убедиться что файл правильный
cats = df["Категория_3"].dropna().value_counts().head(10)
print("\nТоп категории в файле:")
for cat, cnt in cats.items():
    print(f"  {cnt:4d}  {cat}")

MOY_OBEM = 1.0

df["nl"] = df["Название"].astype(str).str.lower()
df["Объём_л"] = pd.to_numeric(df["Объём_л"], errors="coerce")

EXCLUDE_FORMAT = [
    "гранул", "капсул", "саше", "листовой", "лист ", "таблетк",
    "полоск", "стик", "порошок", "шарик",
]
mask_bad_format = df["nl"].str.contains("|".join(EXCLUDE_FORMAT), na=False)
mask_volume = df["Объём_л"].between(MOY_OBEM * 0.75, MOY_OBEM * 2.0)
INCLUDE_PROPS = [
    "парфюм", "parfum", "аромат", "fragrance",
    "концентрат", "concentrate", "ультраконцентрат",
]
mask_props = df["nl"].str.contains("|".join(INCLUDE_PROPS), na=False)

filtered = df[~mask_bad_format & mask_volume & mask_props].copy()
filtered = filtered.sort_values("Заказано_руб", ascending=False).head(30).reset_index(drop=True)
filtered["номер"] = filtered.index + 1

print(f"\nКонкурентов после фильтрации: {len(filtered)}")
print("\nСписок конкурентов:")
for _, row in filtered.iterrows():
    vol = f"{row['Объём_л']:.2f}л" if pd.notna(row['Объём_л']) else "?"
    price = f"{row['Средняя_цена']:.0f}" if pd.notna(row['Средняя_цена']) else "?"
    name = str(row['Название'])[:70]
    print(f"  {row['номер']:2.0f}. {name} | {vol} | {price}руб")
