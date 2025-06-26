import pandas as pd
import re

# Read cost sheet with proper path
cost_df = pd.read_excel(r"D:\Internship\Product cost.xlsx")
cost_df.columns = cost_df.columns.str.strip().str.lower()

cost_dict = {}
for _, row in cost_df.iterrows():
    item = str(row['item']).strip().lower()
    rate_per_cu_inch = row['rate per']  # Double-check actual column name, apply .str.lower() accordingly
    cost_dict[item] = rate_per_cu_inch

# Read price list with proper path
comb_df = pd.read_excel(r"D:\Internship\Price list.xlsx", sheet_name='Price List', header=1)
comb_df.columns = comb_df.columns.str.strip().str.lower()

print("Columns in Price List Sheet:", comb_df.columns.tolist())  # Verify columns

# Print item names to verify
for _, row in comb_df.iterrows():
    item = str(row['item']).strip().lower()
    print(item)

# Filter product columns
product_columns = [col for col in comb_df.columns if isinstance(col, str) and (
    'coir' in col.lower() or 'foam' in col.lower() or 'spring' in col.lower() or 'topper' in col.lower() or 'quilt' in col.lower()
)]

# Price calculation function
def calculate_price(formula, surface_area):
    total_cost = 0
    layers = re.findall(r'([A-Za-z &]+)\s*(\d*)\"?', formula)
    
    for layer, thickness in layers:
        item_name = layer.strip().lower()
        thickness = float(thickness) if thickness else 0

        if 'quilt' in item_name or 'cloth' in item_name:
            area_cost = surface_area * cost_dict.get(item_name, 0)
            total_cost += area_cost
        else:
            volume_cost = surface_area * thickness * cost_dict.get(item_name, 0)
            total_cost += volume_cost
            
    return round(total_cost)

# Apply calculation
for col in product_columns:
    formula = col
    comb_df[col] = comb_df.apply(lambda row: calculate_price(formula, row['surface']), axis=1)

# Save updated sheet
comb_df.to_excel(r'D:\Internship\updated_price_list.xlsx', index=False)
print("Updated price sheet generated as 'updated_price_list.xlsx'")
