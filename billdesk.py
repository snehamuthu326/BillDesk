import itertools
import pandas as pd
import tkinter as tk
from tkinter import messagebox
import os
import subprocess

# ------------------ Read Price Sheet ------------------

price_sheet_path = r"D:\Internship\Costing_Sheet.xlsx"
output_path = r"D:\Internship\Final_Mattress_Matrix_With_MRP_Dealer.xlsx"

df = pd.read_excel(price_sheet_path, sheet_name="new")

# Create product-rate dictionary
product_rates = {row['Product']: row['Rate'] for _, row in df.iterrows()}

# Production & Retail Rates
labour_rate = df.loc[df['Production'] == 'Labour', 'ProRate'].values[0]
transport_rate = df.loc[df['Production'] == 'Transport (Default: 350)', 'ProRate'].values[0]
indirect_expense_rate = df.loc[df['Production'] == 'Indirect & Office expense (Default: 7)', 'ProRate'].values[0]
wastage_rate = df.loc[df['Production'] == 'Wastage (Default: 3)', 'ProRate'].values[0]

margin_percent = df.loc[df['Retail'] == 'Margin 25%', 'ReRate'].values[0]
tax_percent = df.loc[df['Retail'] == 'Tax 18%', 'ReRate'].values[0]
working_cap_percent = df.loc[df['Retail'] == 'Working Capital Interest (Default: 5)', 'ReRate'].values[0]
dealer_margin_percent = df.loc[df['Retail'] == 'Dealer Margin', 'ReRate'].values[0]

# Packaging
pvc_packing_rate = product_rates["PVC Packing"]
flat_packing_cost = product_rates["Thread, Cornershoe, Label"]

# ------------------ Material Options ------------------

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

quilting_thickness = 1  # Fixed

# ------------------ Mattress Combinations ------------------

mattress_combinations = []

for core, core_thick in core_options:
    if "Coir" in core:
        for fabric, fabric_thick in fabric_options:
            mattress_combinations.append((core, core_thick, "None", 0, fabric, fabric_thick))
        for foam, foam_thick in foam_options[1:]:
            for fabric, fabric_thick in fabric_options:
                mattress_combinations.append((core, core_thick, foam, foam_thick, fabric, fabric_thick))
    elif core in ["Topper", "Natural Latex", "Memory foam", "Srilanka Latex Rebond", "Foam - Rebonded", "EP Foam", "PU Foam"] or "Spring" in core:
        for foam, foam_thick in foam_options:
            for fabric, fabric_thick in fabric_options:
                mattress_combinations.append((core, core_thick, foam, foam_thick, fabric, fabric_thick))

# ------------------ Size Options ------------------

length_options = [72, 75, 78, 84]
width_options = [30, 36, 42, 44, 48, 60, 72]

# ------------------ Calculation Function ------------------

def calculate_prices(length, width):
    surface_area = length * width
    mrp_list = []
    dealer_list = []

    for core, core_thick, foam, foam_thick, fabric, fabric_thick in mattress_combinations:
        core_rate = product_rates[core]
        foam_rate = product_rates[foam] if foam != "None" else 0
        fabric_rate = product_rates[fabric]
        quilting_rate = product_rates["Quilting"]

        core_cost = surface_area * core_thick * core_rate
        foam_cost = surface_area * foam_thick * foam_rate * 2
        fabric_cost = surface_area * fabric_thick * fabric_rate * 2
        quilting_cost = surface_area * quilting_thickness * quilting_rate * 2
        pvc_cost = surface_area * pvc_packing_rate

        material_cost = core_cost + foam_cost + fabric_cost + quilting_cost + pvc_cost + flat_packing_cost
        thickness = core_thick + foam_thick + fabric_thick + quilting_thickness

        total_cost = material_cost + (labour_rate * thickness * surface_area) + transport_rate
        total_cost += (indirect_expense_rate * material_cost / 100) + (wastage_rate * material_cost / 100)

        mrp = total_cost * (1 + margin_percent / 100)
        mrp += mrp * (tax_percent / 100) + mrp * (working_cap_percent / 100)
        mrp = round(mrp, 2)

        dealer_price = round(mrp + (mrp * dealer_margin_percent / 100), 2)

        mrp_list.append(mrp)
        dealer_list.append(dealer_price)

    return mrp_list, dealer_list

# ------------------ Matrix Generation ------------------

def generate_matrix():
    columns = ["Length", "Width"]
    columns += [f"{core} {core_thick}\" | {foam} {foam_thick}\" | {fabric} | MRP" 
                for core, core_thick, foam, foam_thick, fabric, fabric_thick in mattress_combinations]
    columns += [f"{core} {core_thick}\" | {foam} {foam_thick}\" | {fabric} | Dealer Price" 
                for core, core_thick, foam, foam_thick, fabric, fabric_thick in mattress_combinations]

    data = []

    for length, width in itertools.product(length_options, width_options):
        mrp_list, dealer_list = calculate_prices(length, width)
        row = [length, width] + mrp_list + dealer_list
        data.append(row)

    df = pd.DataFrame(data, columns=columns)
    df.to_excel(output_path, index=False)

# ------------------ GUI Functions ------------------

def add_custom_size():
    try:
        length = float(length_entry.get())
        width = float(width_entry.get())
        if length <= 0 or width <= 0:
            raise ValueError
    except ValueError:
        messagebox.showerror("Invalid Input", "Enter valid positive numbers for Length and Width.")
        return

    mrp_list, dealer_list = calculate_prices(length, width)
    row = [length, width] + mrp_list + dealer_list

    df = pd.read_excel(output_path)
    df.loc[len(df)] = row
    df.to_excel(output_path, index=False)

    messagebox.showinfo("Success", "Custom size added to Matrix and Excel updated.")

def view_matrix():
    try:
        os.startfile(output_path)
    except Exception:
        subprocess.Popen(["start", output_path], shell=True)

def live_bill():
    try:
        length = float(live_length_entry.get())
        width = float(live_width_entry.get())
        if length <= 0 or width <= 0:
            raise ValueError
    except ValueError:
        messagebox.showerror("Invalid Input", "Enter valid positive numbers for Length and Width.")
        return

    mrp_list, dealer_list = calculate_prices(length, width)
    msg = f"MRP for First Combination: ₹{mrp_list[0]}\nDealer Price: ₹{dealer_list[0]}"
    messagebox.showinfo("Live Bill", msg)

# ------------------ Main Execution ------------------

generate_matrix()

root = tk.Tk()
root.title("Mattress Billing & MRP Generator")

tk.Label(root, text="Custom Size (Length x Width in inches):").pack(pady=5)

length_entry = tk.Entry(root)
length_entry.pack()

width_entry = tk.Entry(root)
width_entry.pack()

tk.Button(root, text="Add to Matrix", command=add_custom_size).pack(pady=5)
tk.Button(root, text="View MRP Matrix", command=view_matrix).pack(pady=5)

tk.Label(root, text="Live Bill (Any Size):").pack(pady=10)

live_length_entry = tk.Entry(root)
live_length_entry.pack()

live_width_entry = tk.Entry(root)
live_width_entry.pack()

tk.Button(root, text="Generate Live Bill", command=live_bill).pack(pady=5)

root.mainloop()
