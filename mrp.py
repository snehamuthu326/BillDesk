import itertools
import pandas as pd
import json

# ------------------ Read Price Sheet ------------------

price_sheet_path = r"D:\Internship\Costing_Sheet.xlsx"
df = pd.read_excel(price_sheet_path, sheet_name="new")

# Create product-rate dictionary
product_rates = {row['Product']: row['Rate'] for _, row in df.iterrows()}
#print(json.dumps(product_rates, indent = 2))

# Production Costs
labour_rate = df.loc[df['Production'] == 'Labour', 'ProRate'].values[0]
transport_rate = df.loc[df['Production'] == 'Transport (Default: 350)', 'ProRate'].values[0]
indirect_expense_rate = df.loc[df['Production'] == 'Indirect & Office expense (Default: 7)', 'ProRate'].values[0]
wastage_rate = df.loc[df['Production'] == 'Wastage (Default: 3)', 'ProRate'].values[0]

# Retail Markups
margin_percent = df.loc[df['Retail'] == 'Margin 25%', 'ReRate'].values[0]
tax_percent = df.loc[df['Retail'] == 'Tax 18%', 'ReRate'].values[0]
working_cap_percent = df.loc[df['Retail'] == 'Working Capital Interest (Default: 5)', 'ReRate'].values[0]
dealer_margin_percent = df.loc[df['Retail'] == 'Dealer Margin', 'ReRate'].values[0]

# Packaging Costs
pvc_packing_rate = product_rates["PVC Packing"]
flat_packing_cost = product_rates["Thread, Cornershoe, Label"]

# ------------------ Material Options with Thickness ------------------

core_options = [
    ("Coir 80D", 1), ("Coir 80D", 2), ("Coir 80D", 4),
    ("Coir 90D", 2),
    ("Coir 100D", 2),
    ("Topper", 1),
    ("Natural Latex", 2),
    ("Memory foam", 2),
    ("Srilanka Latex Rebond", 2),
    ("Foam - Rebonded", 2),
    ("Bonnel (only 5) Spring", 5),
    ("Pocketed (only 5) Spring", 5),
    ("EP Foam", 2),
    ("PU Foam", 1), ("PU Foam", 2)
]

foam_options = [
    ("None", 0),
    ("Single Foam", 1),
    ("Single foam + Single foam", 2),
    ("Double foam + double foam", 4)
]

fabric_options = [
    ("Fabric Regular (120 GSM)", 0.5),
    ("Fabric Premium (250 GSM)", 0.5),
    ("Fabric Ultra Premium (350 GSM)", 0.5)
]

quilting_thickness = 1  # Fixed for all

# ------------------ Size Options ------------------

length_options = [72, 75, 78, 84]
width_options = [30, 36, 42, 44, 48, 60, 72]

# ------------------ Generate Material Combinations ------------------

mattress_combinations = []

for core, core_thick in core_options:

    if "Coir" in core:
        # Coir Only
        for fabric, fabric_thick in fabric_options:
            combo_name = f"{core} {core_thick}\" | None | {fabric}"
            mattress_combinations.append((combo_name, core, core_thick, "None", 0, fabric, fabric_thick))

        # Coir + Foam
        for foam, foam_thick in foam_options[1:]:  # Skip "None"
            for fabric, fabric_thick in fabric_options:
                combo_name = f"{core} {core_thick}\" | {foam} {foam_thick}\" | {fabric}"
                mattress_combinations.append((combo_name, core, core_thick, foam, foam_thick, fabric, fabric_thick))

    elif core in ["Topper", "Natural Latex", "Memory foam", "Srilanka Latex Rebond", "Foam - Rebonded"]:
        for foam, foam_thick in foam_options:
            for fabric, fabric_thick in fabric_options:
                combo_name = f"{core} {core_thick}\" | {foam} {foam_thick}\" | {fabric}"
                mattress_combinations.append((combo_name, core, core_thick, foam, foam_thick, fabric, fabric_thick))

    elif "Spring" in core:
        for foam, foam_thick in foam_options:
            for fabric, fabric_thick in fabric_options:
                combo_name = f"{core} {core_thick}\" | {foam} {foam_thick}\" | {fabric}"
                mattress_combinations.append((combo_name, core, core_thick, foam, foam_thick, fabric, fabric_thick))

    elif core in ["EP Foam", "PU Foam"]:
        for foam, foam_thick in foam_options:
            for fabric, fabric_thick in fabric_options:
                combo_name = f"{core} {core_thick}\" | {foam} {foam_thick}\" | {fabric}"
                mattress_combinations.append((combo_name, core, core_thick, foam, foam_thick, fabric, fabric_thick))

print(f"Total mattress combinations generated: {len(mattress_combinations)}")

# ------------------ Generate Size Combinations ------------------

size_combinations = list(itertools.product(length_options, width_options))

# ------------------ Create Final Matrix ------------------

columns = ["Length", "Width"] + [f"{combo[0]} | MRP" for combo in mattress_combinations] + [f"{combo[0]} | Dealer Price" for combo in mattress_combinations]
data = []

for length, width in size_combinations:
    surface_area = length * width
    row = [length, width]

    for combo in mattress_combinations:
        core, core_thick = combo[1], combo[2]
        foam, foam_thick = combo[3], combo[4]
        fabric, fabric_thick = combo[5], combo[6]

        # Rates from price sheet
        core_rate = product_rates[core]
        foam_rate = product_rates[foam] if foam != "None" else 0
        fabric_rate = product_rates[fabric]
        quilting_rate = product_rates["Quilting"]

        # ------------------ Cost Calculation ------------------
        core_cost = surface_area * core_thick * core_rate
        foam_cost = surface_area * foam_thick * foam_rate * 2
        fabric_cost = surface_area * fabric_thick * fabric_rate * 2
        quilting_cost = surface_area * quilting_thickness * quilting_rate * 2
        pvc_cost = surface_area * pvc_packing_rate
        material_cost = core_cost + foam_cost + fabric_cost + quilting_cost + pvc_cost + flat_packing_cost

        thickness = core_thick + foam_thick + fabric_thick + quilting_thickness
        # Add production costs
        total_cost = material_cost + (labour_rate*thickness*surface_area) + transport_rate + (indirect_expense_rate*material_cost/100) + (wastage_rate*material_cost/100)

        # ------------------ MRP Calculation ------------------
        mrp = total_cost * (1 + margin_percent / 100)
        mrp += mrp * (tax_percent / 100)
        mrp += mrp * (working_cap_percent / 100)
        mrp = round(mrp, 2)

        # Dealer Price
        dealer_price = mrp + (mrp * dealer_margin_percent / 100)
        dealer_price = round(dealer_price, 2)

        row.append(mrp)
        #row.append(dealer_price) 
        
        #two rows one mrp and another dealer price atlast

    data.append(row)

# ------------------ Export to Excel ------------------

df_final = pd.DataFrame(data, columns=columns)
output_path = r"D:\Internship\Final_Mattress_Matrix_With_MRP_Dealer.xlsx"
df_final.to_excel(output_path, index=False)

print(f"Matrix with MRP & Dealer Price saved to: {output_path}")
