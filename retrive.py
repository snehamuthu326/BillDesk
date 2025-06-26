import pandas as pd
import json

df = pd.read_excel(r"D:\Internship\Costing_sheet.xlsx", sheet_name="Sheet1", header=0)

product_rate_map = {}

for index, row in df.iterrows():
    product_rate_map[row['Product']] = row['Rate']

print(json.dumps(product_rate_map, indent = 2))

