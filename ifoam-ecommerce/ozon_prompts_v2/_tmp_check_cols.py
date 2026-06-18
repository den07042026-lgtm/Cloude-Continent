# -*- coding: utf-8 -*-
import pandas as pd

folder = 'C:/Users/1/Downloads/ifoamHOME_КондиционерBalmy_РайскиеЦветы_770758/'
df = pd.read_excel(folder + '03_competitors_770758.xlsx', header=None)
headers = df.iloc[3].tolist()
print("Заголовки:")
for i, h in enumerate(headers):
    print(f"  {i}: {h}")

print("\nДанные строки 4 (первый конкурент):")
row = df.iloc[4].tolist()
for i, v in enumerate(row):
    print(f"  {i} [{headers[i]}]: {v}")
