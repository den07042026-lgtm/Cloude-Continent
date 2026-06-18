# -*- coding: utf-8 -*-
import sys, pandas as pd
sys.path.insert(0, "C:/Users/1/Downloads/ozon_prompts_v2")
from save_helpers import get_product_folder

folder = get_product_folder("ifoamHOME", "КондиционерBalmy_РайскиеЦветы", "770758")
bs_file = str(folder / "analytics_report_2026-03-26_15_49.xlsx")

# Читаем RAW заголовки из строки 4 (индекс 4)
df_raw = pd.read_excel(bs_file, header=None)
headers_row = df_raw.iloc[4].tolist()
print("=== РЕАЛЬНЫЕ ЗАГОЛОВКИ (строка 4) ===")
for i, h in enumerate(headers_row):
    print(f"  col {i:2d}: {h}")

# Также смотрим несколько строк данных для col 18-22
print("\n=== ДАННЫЕ col 18-22 (строки 6-12) ===")
for row_i in range(6, 13):
    row = df_raw.iloc[row_i]
    vals = [str(row[i])[:30] for i in range(18, 23)]
    print(f"  row {row_i}: {vals}")
