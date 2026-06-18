# -*- coding: utf-8 -*-
import sys, pandas as pd, glob
sys.path.insert(0, "C:/Users/1/Downloads/ozon_prompts_v2")
from save_helpers import get_product_folder, make_xlsx_competitors, make_docx

folder = get_product_folder("ifoamHOME", "КондиционерBalmy_РайскиеЦветы", "770758")

# Читаем bestsellers
bs_files = sorted(glob.glob("C:/Users/1/Downloads/analytics_report_*.xlsx"))
bs_file = bs_files[-1]
print(f"Bestsellers: {bs_file}")

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

MOY_OBEM = 1.0  # наш объём в литрах

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
filtered["№"] = filtered.index + 1

# Цена за объём нашего товара (1л)
filtered["Цена_экв_1л"] = (filtered["Средняя_цена"] * MOY_OBEM / filtered["Объём_л"]).round(0)
filtered.loc[filtered["Объём_л"].isna(), "Цена_экв_1л"] = filtered["Средняя_цена"]

filtered["Реальная_цена"] = filtered["Мин_цена"].fillna(filtered["Средняя_цена"])
filtered["Рейтинг"] = "[→ Промпт 6]"
filtered["Отзывов"] = "[→ Промпт 6]"

export_cols = [
    "№", "Название", "Бренд", "Продавец",
    "Объём_л", "Средняя_цена", "Мин_цена", "Реальная_цена",
    "Цена_экв_1л",  # <-- цена за 1л как у нашего товара
    "Заказано_руб", "Заказано_шт", "Динамика_%",
    "Ср_продажи_шт_день", "Ср_продажи_руб_день",
    "Доля_выкупа_%", "Упущенные_продажи_руб", "Дней_без_остатка",
    "Конверсия_показ_заказ_%", "В_корзину_поиск_%", "В_корзину_карточка_%",
    "Показы_всего", "Просмотры_карточки",
    "Скидка_акции", "Доля_оборота_акции_%", "Дней_в_акциях",
    "Дней_с_продвижением", "ДРР_%",
    "Схема", "Остаток_шт", "Дата_создания",
    "Рейтинг", "Отзывов",
    "Ссылка",
]
out = filtered[[c for c in export_cols if c in filtered.columns]]

# Основной файл
make_xlsx_competitors(folder, "03_competitors_770758.xlsx", out)
print(f"Сохранено: 03_competitors_770758.xlsx ({len(filtered)} конкурентов)")

# Альтернативный без Tesori
mask_tesori = (
    filtered["Название"].astype(str).str.lower().str.contains("tesori", na=False) |
    filtered["Бренд"].astype(str).str.lower().str.contains("tesori", na=False)
)
filtered_alt = filtered[~mask_tesori].copy()
filtered_alt = filtered_alt.reset_index(drop=True)
filtered_alt["№"] = filtered_alt.index + 1
out_alt = filtered_alt[[c for c in export_cols if c in filtered_alt.columns]]

make_xlsx_competitors(folder, "03_competitors_770758_alt.xlsx", out_alt)
print(f"Сохранено: 03_competitors_770758_alt.xlsx ({len(filtered_alt)} конкурентов)")

# Вывод информации о колонке
print("\nПример значений Цена_экв_1л (цена за 1л как у нашего товара):")
preview = filtered[["Название","Объём_л","Средняя_цена","Цена_экв_1л"]].head(10)
for _, row in preview.iterrows():
    print(f"  {str(row['Название'])[:50]:50s} | {row['Объём_л']:.2f}л | {row['Средняя_цена']:.0f}₽ → {row['Цена_экв_1л']:.0f}₽/1л")
