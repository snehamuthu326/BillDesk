import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd

# Load Excel data
file_path = "Costing Sheet for Mattress Dated on 19-08-2022 .xlsx - Internship.xlsx"
df = pd.read_excel(file_path, sheet_name="Format 25.11.2023", header=None)

# Extract default data from Excel
default_core = df.iloc[4, 1].strip()  # Coir 80D
default_length = int(df.iloc[1, 2])   # 75
default_width = int(df.iloc[1, 3])    # 36
#default_thickness = int(df.iloc[1, 1])  # 6
value = df.iloc[1, 1]
if pd.notna(value):
    default_thickness = int(value)
else:
    default_thickness = 4  # or some default value like 6

default_discount = float(df.iloc[3, 2])  # 40
default_mrp = float(df.iloc[3, 0])  # MRP in Excel

# Extract core cost table
core_costs = {}
for i in range(4, 15):
    material = df.iloc[i, 1]
    cost = df.iloc[i, 3]
    if pd.notna(material) and pd.notna(cost):
        core_costs[material.strip()] = float(cost)

# --------------------------- GUI ---------------------------

def handle_user_choice(choice):
    if choice == "No":
        # User accepts defaults; move to step 3 (PDF logic placeholder)
        area = default_length * default_width
        cost = core_costs[default_core]
        mrp = area * cost
        net = mrp - (mrp * (default_discount / 100))

        messagebox.showinfo("Default Bill Summary", 
            f"Core: {default_core}\n"
            f"Size: {default_length} x {default_width} x {default_thickness} in\n"
            f"Area: {area} in²\n"
            f"Cost per inch²: ₹{cost:.3f}\n"
            f"MRP: ₹{mrp:.2f}\n"
            f"Discount: {default_discount}%\n"
            f"Net Price: ₹{net:.2f}\n\n"
            f"➡ Ready to proceed to Step 3: Generate PDF"
        )
        # We’ll plug in PDF generation here later

    else:
        # Go to customization form
        root.destroy()
        launch_custom_form()

def launch_custom_form():
    form = tk.Tk()
    form.title("Customize Product Configuration")

    # Core selection
    tk.Label(form, text="Core Type").grid(row=0, column=0)
    core_var = tk.StringVar(value=default_core)
    core_menu = ttk.Combobox(form, textvariable=core_var, values=list(core_costs.keys()))
    core_menu.grid(row=0, column=1)

    # Size inputs
    def make_entry(label, val, row):
        tk.Label(form, text=label).grid(row=row, column=0)
        entry = tk.Entry(form)
        entry.insert(0, str(val))
        entry.grid(row=row, column=1)
        return entry

    length_entry = make_entry("Length (in)", default_length, 1)
    width_entry = make_entry("Width (in)", default_width, 2)
    thickness_entry = make_entry("Thickness (in)", default_thickness, 3)
    discount_entry = make_entry("Discount (%)", default_discount, 4)

    # Calculate button
    def calculate_custom():
        try:
            core = core_var.get()
            length = float(length_entry.get())
            width = float(width_entry.get())
            thickness = float(thickness_entry.get())
            discount = float(discount_entry.get())

            area = length * width
            cost = core_costs[core]
            mrp = area * cost
            net = mrp - (mrp * (discount / 100))

            messagebox.showinfo("Custom Bill", 
                f"Core: {core}\nSize: {length} x {width} x {thickness} in\n"
                f"Area: {area} in²\nCost per inch²: ₹{cost:.3f}\n"
                f"MRP: ₹{mrp:.2f}\nDiscount: {discount}%\nNet Price: ₹{net:.2f}"
            )
        except:
            messagebox.showerror("Error", "Invalid input values!")

    tk.Button(form, text="Calculate MRP", command=calculate_custom).grid(row=5, columnspan=2, pady=10)

    form.mainloop()

# Main window to ask user
root = tk.Tk()
root.title("Mattress Billing App")

tk.Label(root, text="Default Product Info", font=("Arial", 12, "bold")).pack(pady=10)
tk.Label(root, text=(
    f"Core: {default_core}\n"
    f"Size: {default_length} x {default_width} x {default_thickness} inches\n"
    f"Discount: {default_discount}%"
)).pack(pady=5)

tk.Label(root, text="Do you want to change product configuration?").pack(pady=10)

tk.Button(root, text="No", width=20, command=lambda: handle_user_choice("No")).pack(pady=5)
tk.Button(root, text="Yes", width=20, command=lambda: handle_user_choice("Yes")).pack(pady=5)

root.mainloop()
