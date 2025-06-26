import pandas as pd
import json


price_sheet_path = r"D:\Internship\Costing_sheet.xlsx"
df = pd.read_excel(price_sheet_path, sheet_name="new")

# Create product-rate dictionary
product_rates = {row['Product']: row['Rate'] for _, row in df.iterrows()}

print(json.dumps(product_rates, indent = 2))

#-------------------------------------


"""

df = pd.read_excel(r"D:\Internship\Costing_sheet.xlsx", sheet_name="new", header=0)

product = {}

for index, row in df.iterrows():
    product[row['Product']] = row['Rate']

#print(json.dumps(product, indent = 2))

print(product['Coir 80D'])

"""