#Complete code for custom matrix

import itertools
import pandas as pd
import tkinter as tk
from tkinter import messagebox, END, ttk
import os
import subprocess

# ------------------ Read Price Sheet ------------------
price_sheet_path = r"D:\Projects\Internship\Costing_Sheet.xlsx"
output_path = r"D:\Projects\Internship\Final_Mattress_Matrix_With_MRP_Dealer.xlsx"

df = pd.read_excel(price_sheet_path, sheet_name="new")
product_rates = {row['Product']: row['Rate'] for _, row in df.iterrows()}
labour_rate = df.loc[df['Production'] == 'Labour', 'ProRate'].values[0]
transport_rate = df.loc[df['Production'] == 'Transport (Default: 350)', 'ProRate'].values[0]
indirect_expense_rate = df.loc[df['Production'] == 'Indirect & Office expense (Default: 7)', 'ProRate'].values[0]
wastage_rate = df.loc[df['Production'] == 'Wastage (Default: 3)', 'ProRate'].values[0]
margin_percent = df.loc[df['Retail'] == 'Margin 25%', 'ReRate'].values[0]
tax_percent = df.loc[df['Retail'] == 'Tax 18%', 'ReRate'].values[0]
working_cap_percent = df.loc[df['Retail'] == 'Working Capital Interest (Default: 5)', 'ReRate'].values[0]
dealer_margin_percent = df.loc[df['Retail'] == 'Dealer Margin', 'ReRate'].values[0]
pvc_packing_rate = product_rates["PVC Packing"]
flat_packing_cost = product_rates["Thread, Cornershoe, Label"]

core_options = [("Coir 80D", 1), ("Coir 80D", 2), ("Coir 80D", 4), ("Coir 90D", 2), ("Coir 100D", 2),
    ("Topper", 1), ("Natural Latex", 2), ("Memory foam", 2), ("Srilanka Latex Rebond", 2), ("Foam - Rebonded", 2),
    ("Bonnel (only 5) Spring", 5), ("Pocketed (only 5) Spring", 5), ("EP Foam", 2), ("PU Foam", 1), ("PU Foam", 2)]

foam_options = [("None", 0), ("Single Foam", 1), ("Single foam + Single foam", 2), ("Double foam + double foam", 4)]

fabric_options = [("Fabric Regular (120 GSM)", 0.5), ("Fabric Premium (250 GSM)", 0.5), ("Fabric Ultra Premium (350 GSM)", 0.5)]

quilting_thickness = 1
mattress_combinations = []

for core, core_thick in core_options:
    if "Coir" in core:
        for fabric, fabric_thick in fabric_options:
            mattress_combinations.append([(core, core_thick), ("None", 0), (fabric, fabric_thick)])
        for foam, foam_thick in foam_options[1:]:
            for fabric, fabric_thick in fabric_options:
                mattress_combinations.append([(core, core_thick), (foam, foam_thick), (fabric, fabric_thick)])
    elif core in ["Topper", "Natural Latex", "Memory foam", "Srilanka Latex Rebond", "Foam - Rebonded", "EP Foam", "PU Foam"] or "Spring" in core:
        for foam, foam_thick in foam_options:
            for fabric, fabric_thick in fabric_options:
                mattress_combinations.append([(core, core_thick), (foam, foam_thick), (fabric, fabric_thick)])

length_options = [72, 75, 78, 84]
width_options = [30, 36, 42, 44, 48, 60, 72]
custom_sizes = []
custom_products_list = []
current_product_layers = []

root = tk.Tk()
root.title("Mattress Billing & Custom Product Generator")
root.geometry("550x800")

main_canvas = tk.Canvas(root, bg="#f0f4f7")
main_scrollbar = tk.Scrollbar(root, orient="vertical", command=main_canvas.yview)
main_scrollbar.pack(side="right", fill="y")
main_canvas.pack(side="left", fill="both", expand=True)
main_canvas.configure(yscrollcommand=main_scrollbar.set)

frame = tk.Frame(main_canvas, bg="#f0f4f7")
main_canvas.create_window((0, 0), window=frame, anchor="nw")

def configure_scroll(event):
    main_canvas.configure(scrollregion=main_canvas.bbox("all"))

frame.bind("<Configure>", configure_scroll)

style = ttk.Style()
style.configure("TButton", font=("Arial", 10, "bold"), padding=5)
style.configure("TLabel", font=("Arial", 10))

section_title = tk.Label(frame, text="Customise Your Costing Sheet", bg="#f0f4f7", font=("Arial", 11, "bold"))
section_title.pack(pady=5)

material_frame = tk.Frame(frame, bg="#f0f4f7")
material_frame.pack(pady=5)
tk.Label(material_frame, text="Select Material:", bg="#f0f4f7", font=("Arial", 10)).pack(side="left", padx=5)
material_var = tk.StringVar(value=list(product_rates.keys())[0])
ttk.Combobox(material_frame, textvariable=material_var, values=list(product_rates.keys()), state="readonly").pack(side="left", padx=5)

thickness_frame = tk.Frame(frame, bg="#f0f4f7")
thickness_frame.pack(pady=5)
tk.Label(thickness_frame, text="Thickness (inches):", bg="#f0f4f7", font=("Arial", 10)).pack(side="left", padx=5)
thickness_entry = tk.Entry(thickness_frame)
thickness_entry.pack(side="left", padx=5)

item_btn_frame = tk.Frame(frame, bg="#f0f4f7")
item_btn_frame.pack(pady=5)
ttk.Button(item_btn_frame, text="Add Item", command=lambda: add_layer()).pack(side="left", padx=5)
ttk.Button(item_btn_frame, text="Delete Selected Item", command=lambda: delete_layer()).pack(side="left", padx=5)

layer_listbox = tk.Listbox(frame, height=5, bg="white", font=("Arial", 10))
layer_listbox.pack(pady=5, fill="x")

final_btn = ttk.Button(frame, text="Finalize Product", command=lambda: finalize_product())
final_btn.pack(pady=5)

product_frame = tk.Frame(frame, bg="#f0f4f7")
product_frame.pack(pady=5, fill="x")
product_listbox = tk.Listbox(product_frame, height=5, bg="#e6f2ff", font=("Arial", 10))
product_listbox.pack(side="left", fill="x", expand=True)
ttk.Button(product_frame, text="Delete Selected Product", command=lambda: delete_custom_product()).pack(side="right", padx=5)

size_title = tk.Label(frame, text="Add Custom Sizes", bg="#f0f4f7", font=("Arial", 10, "bold"))
size_title.pack(pady=5)

size_frame = tk.Frame(frame, bg="#f0f4f7")
size_frame.pack(pady=5)
tk.Label(size_frame, text="Length:", bg="#f0f4f7", font=("Arial", 10)).pack(side="left", padx=5)
custom_length_entry = tk.Entry(size_frame, width=5)
custom_length_entry.pack(side="left", padx=5)
tk.Label(size_frame, text="Width:", bg="#f0f4f7", font=("Arial", 10)).pack(side="left", padx=5)
custom_width_entry = tk.Entry(size_frame, width=5)
custom_width_entry.pack(side="left", padx=5)

size_btn_frame = tk.Frame(frame, bg="#f0f4f7")
size_btn_frame.pack(pady=5)
ttk.Button(size_btn_frame, text="Add Size", command=lambda: add_custom_size()).pack(side="left", padx=5)
ttk.Button(size_btn_frame, text="Delete Selected Size", command=lambda: delete_custom_size()).pack(side="left", padx=5)

custom_size_listbox = tk.Listbox(frame, height=4, bg="#e6f2ff", font=("Arial", 10))
custom_size_listbox.pack(pady=5, fill="x")

generate_btn = ttk.Button(frame, text="View Customised Costing Sheet", command=lambda: generate_matrix())
generate_btn.pack(pady=10)

def add_layer():
    material = material_var.get()
    try:
        thickness = float(thickness_entry.get())
        if thickness <= 0 or material not in product_rates:
            raise ValueError
    except ValueError:
        messagebox.showerror("Error", "Enter valid material and positive thickness.")
        return
    current_product_layers.append((material, thickness))
    layer_listbox.insert(END, f"{material} - {thickness}\"")
    thickness_entry.delete(0, END)

def delete_layer():
    sel = layer_listbox.curselection()
    if not sel:
        messagebox.showerror("Error", "Select an item to delete.")
        return
    index = sel[0]
    layer_listbox.delete(index)
    del current_product_layers[index]

def finalize_product():
    if not current_product_layers:
        messagebox.showerror("Error", "Add at least one layer.")
        return
    custom_products_list.append(current_product_layers.copy())
    desc = " + ".join([f"{mat} {thk}\"" for mat, thk in current_product_layers])
    product_listbox.insert(END, desc)
    current_product_layers.clear()
    layer_listbox.delete(0, END)

def delete_custom_product():
    sel = product_listbox.curselection()
    if not sel:
        messagebox.showerror("Error", "Select a product to delete.")
        return
    index = sel[0]
    product_listbox.delete(index)
    del custom_products_list[index]

def add_custom_size():
    try:
        length = float(custom_length_entry.get())
        width = float(custom_width_entry.get())
        if length <= 0 or width <= 0:
            raise ValueError
    except ValueError:
        messagebox.showerror("Error", "Enter valid positive Length and Width.")
        return
    custom_sizes.append((length, width))
    custom_size_listbox.insert(END, f"{length} x {width} inches")
    custom_length_entry.delete(0, END)
    custom_width_entry.delete(0, END)

def delete_custom_size():
    sel = custom_size_listbox.curselection()
    if not sel:
        messagebox.showerror("Error", "Select a size to delete.")
        return
    index = sel[0]
    custom_size_listbox.delete(index)
    del custom_sizes[index]

def calculate_total_cost(length, width, layers):
    area = length * width
    material_cost = sum(area * thk * product_rates.get(mat, 0) for mat, thk in layers)
    total_thickness = sum(thk for _, thk in layers) + quilting_thickness
    quilting_rate = product_rates.get("Quilting", 0)
    material_cost += area * quilting_thickness * quilting_rate * 2
    material_cost += area * pvc_packing_rate + flat_packing_cost
    total_cost = material_cost + (labour_rate * total_thickness * area) + transport_rate
    total_cost += (indirect_expense_rate * material_cost / 100) + (wastage_rate * material_cost / 100)
    mrp = total_cost * (1 + margin_percent / 100)
    mrp += mrp * (tax_percent / 100) + mrp * (working_cap_percent / 100)
    dealer_price = mrp + (mrp * dealer_margin_percent / 100)
    return round(mrp, 2), round(dealer_price, 2)

def generate_matrix():
    columns = ["Length", "Width"]
    for prod in custom_products_list:
        desc = " + ".join([f"{mat} {thk}\"" for mat, thk in prod])
        columns += [f"{desc} | Net Rate", f"{desc} | MRP"]
    for combo in mattress_combinations:
        desc = " + ".join([f"{mat} {thk}\"" for mat, thk in combo])
        columns += [f"{desc} | Net Rate", f"{desc} | MRP"]
    all_lengths = sorted(set(length_options + [l for l, _ in custom_sizes]))
    all_widths = sorted(set(width_options + [w for _, w in custom_sizes]))
    data = []
    for length, width in itertools.product(all_lengths, all_widths):
        row = [length, width]
        for prod in custom_products_list:
            mrp, dealer_price = calculate_total_cost(length, width, prod)
            row += [mrp, dealer_price]
        for combo in mattress_combinations:
            mrp, dealer_price = calculate_total_cost(length, width, combo)
            row += [mrp, dealer_price]
        data.append(row)
    pd.DataFrame(data, columns=columns).to_excel(output_path, index=False)
    try:
        os.startfile(output_path)
    except:
        subprocess.Popen(["start", output_path], shell=True)

root.mainloop()

"""
import itertools
import pandas as pd
import tkinter as tk
from tkinter import messagebox, Listbox, END, ttk
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

# ------------------ Standard Mattress Combinations ------------------

core_options = [
    ("Coir 80D", 1), ("Coir 80D", 2), ("Coir 80D", 4),
    ("Coir 90D", 2), ("Coir 100D", 2),
    ("Topper", 1), ("Natural Latex", 2), ("Memory foam", 2),
    ("Srilanka Latex Rebond", 2), ("Foam - Rebonded", 2),
    ("Bonnel (only 5) Spring", 5), ("Pocketed (only 5) Spring", 5),
    ("EP Foam", 2), ("PU Foam", 1), ("PU Foam", 2)
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

mattress_combinations = []

for core, core_thick in core_options:
    if "Coir" in core:
        for fabric, fabric_thick in fabric_options:
            mattress_combinations.append([(core, core_thick), ("None", 0), (fabric, fabric_thick)])
        for foam, foam_thick in foam_options[1:]:
            for fabric, fabric_thick in fabric_options:
                mattress_combinations.append([(core, core_thick), (foam, foam_thick), (fabric, fabric_thick)])
    elif core in ["Topper", "Natural Latex", "Memory foam", "Srilanka Latex Rebond", "Foam - Rebonded", "EP Foam", "PU Foam"] or "Spring" in core:
        for foam, foam_thick in foam_options:
            for fabric, fabric_thick in fabric_options:
                mattress_combinations.append([(core, core_thick), (foam, foam_thick), (fabric, fabric_thick)])

# ------------------ Size Options ------------------

length_options = [72, 75, 78, 84]
width_options = [30, 36, 42, 44, 48, 60, 72]

# ------------------ Custom Product Storage ------------------

custom_products_list = []  # List of custom products, each is a list of (material, thickness)
current_product_layers = []

# ------------------ Price Calculation ------------------

def calculate_total_cost(length, width, layers):
    surface_area = length * width

    material_cost = 0
    total_thickness = 0

    for material, thickness in layers:
        rate = product_rates.get(material, 0)
        material_cost += surface_area * thickness * rate
        total_thickness += thickness

    quilting_rate = product_rates.get("Quilting", 0)
    pvc_cost = surface_area * pvc_packing_rate

    material_cost += surface_area * quilting_thickness * quilting_rate * 2
    material_cost += pvc_cost + flat_packing_cost
    total_thickness += quilting_thickness

    total_cost = material_cost + (labour_rate * total_thickness * surface_area) + transport_rate
    total_cost += (indirect_expense_rate * material_cost / 100) + (wastage_rate * material_cost / 100)

    mrp = total_cost * (1 + margin_percent / 100)
    mrp += mrp * (tax_percent / 100) + mrp * (working_cap_percent / 100)
    mrp = round(mrp, 2)

    dealer_price = round(mrp + (mrp * dealer_margin_percent / 100), 2)

    return mrp, dealer_price

# ------------------ Matrix Generation ------------------

def generate_matrix():
    columns = ["Length", "Width"]

    for idx, product in enumerate(custom_products_list, 1):
        desc = " + ".join([f"{mat} {thk}\"" for mat, thk in product])
        columns += [f"{desc} | Net Rate", f"{desc} | MRP"]

    for combo in mattress_combinations:
        desc = " + ".join([f"{mat} {thk}\"" for mat, thk in combo])
        columns += [f"{desc} | Net Rate", f"{desc} | MRP"]

    data = []

    for length, width in itertools.product(length_options, width_options):
        row = [length, width]

        for product in custom_products_list:
            mrp, dealer_price = calculate_total_cost(length, width, product)
            row += [mrp, dealer_price]

        for combo in mattress_combinations:
            mrp, dealer_price = calculate_total_cost(length, width, combo)
            row += [mrp, dealer_price]

        data.append(row)

    df_out = pd.DataFrame(data, columns=columns)
    df_out.to_excel(output_path, index=False)

    try:
        os.startfile(output_path)
    except Exception:
        subprocess.Popen(["start", output_path], shell=True)

# ------------------ GUI ------------------

root = tk.Tk()
root.title("Mattress Billing & Multi Custom Product Generator")
root.configure(bg="#f0f4f7")
root.geometry("450x650")

style = ttk.Style()
style.configure("TButton", font=("Arial", 10, "bold"), padding=5)
style.configure("TLabel", font=("Arial", 10))

frame = tk.Frame(root, bg="#f0f4f7")
frame.pack(pady=10, padx=10, fill="both", expand=True)

# Layer Adding Section
tk.Label(frame, text="Customise Your Costing Sheet", bg="#f0f4f7", font=("Arial", 11, "bold")).pack(pady=5)
# Title
tk.Label(frame, text="Select Items for Custom Products Below", bg="#f0f4f7", font=("Arial", 10, "bold")).pack(pady=5)

# Frame for Material Selection
material_frame = tk.Frame(frame, bg="#f0f4f7")
material_frame.pack(pady=5)

tk.Label(material_frame, text="Select Material:", bg="#f0f4f7", font=("Arial", 10)).pack(side="left", padx=5)
material_var = tk.StringVar(value=list(product_rates.keys())[0])
ttk.Combobox(material_frame, textvariable=material_var, values=list(product_rates.keys()), state="readonly").pack(side="left", padx=5)

# Frame for Thickness Input
thickness_frame = tk.Frame(frame, bg="#f0f4f7")
thickness_frame.pack(pady=5)

tk.Label(thickness_frame, text="Enter Thickness (inches):", bg="#f0f4f7", font=("Arial", 10)).pack(side="left", padx=5)
thickness_entry = tk.Entry(thickness_frame)
thickness_entry.pack(side="left", padx=5)



ttk.Button(frame, text="Add Items", command=lambda: add_layer()).pack(pady=5)

layer_listbox = Listbox(frame, height=6, bg="white", font=("Arial", 10))
layer_listbox.pack(pady=5, fill="x")

ttk.Button(frame, text="Finalize Current Product", command=lambda: finalize_product()).pack(pady=5)

product_listbox = Listbox(frame, height=6, bg="#e6f2ff", font=("Arial", 10))
product_listbox.pack(pady=10, fill="x")

frame2 = tk.Frame(frame, bg="#f0f4f7")
frame2.pack(pady=5)

ttk.Button(frame2, text="View Cutomised Costing Sheet", command=lambda: generate_matrix()).pack(side="left", padx=5)

# ------------------ Functions ------------------

def add_layer():
    material = material_var.get()
    try:
        thickness = float(thickness_entry.get())
        if thickness <= 0 or material not in product_rates:
            raise ValueError
    except ValueError:
        messagebox.showerror("Error", "Enter valid material and positive thickness.")
        return

    current_product_layers.append((material, thickness))
    layer_listbox.insert(END, f"{material} - {thickness}\"")
    thickness_entry.delete(0, END)

def finalize_product():
    if not current_product_layers:
        messagebox.showerror("Error", "Add at least one layer.")
        return
    custom_products_list.append(current_product_layers.copy())
    desc = " + ".join([f"{mat} {thk}\"" for mat, thk in current_product_layers])
    product_listbox.insert(END, f"{desc}")
    current_product_layers.clear()
    layer_listbox.delete(0, END)

    
root.mainloop()
"""