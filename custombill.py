import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, Canvas, Frame, Scrollbar
from fpdf import FPDF, XPos, YPos
import os

# ------------------ Read Price Sheet Function ------------------

def read_price_sheet():
    price_sheet_path = r"D:\Projects\Internship\Costing_Sheet.xlsx"
    df = pd.read_excel(price_sheet_path, sheet_name="new")

    product_rates = {row['Product']: row['Rate'] for _, row in df.iterrows()}

    labour_rate = df.loc[df['Production'] == 'Labour', 'ProRate'].values[0]
    transport_rate = df.loc[df['Production'] == 'Transport (Default: 350)', 'ProRate'].values[0]
    indirect_expense_rate = df.loc[df['Production'] == 'Indirect & Office expense (Default: 7)', 'ProRate'].values[0]
    wastage_rate = df.loc[df['Production'] == 'Wastage (Default: 3)', 'ProRate'].values[0]

    margin_percent = df.loc[df['Retail'] == 'Margin 25%', 'ReRate'].values[0]
    tax_percent = df.loc[df['Retail'] == 'Tax 18%', 'ReRate'].values[0]
    working_cap_percent = df.loc[df['Retail'] == 'Working Capital Interest (Default: 5)', 'ReRate'].values[0]

    pvc_packing_rate = product_rates["PVC Packing"]
    flat_packing_cost = product_rates["Thread, Cornershoe, Label"]

    return product_rates, labour_rate, transport_rate, indirect_expense_rate, wastage_rate, margin_percent, tax_percent, working_cap_percent, pvc_packing_rate, flat_packing_cost

# ------------------ MRP Calculation ------------------

def calculate_mrp(length, width, core_layers, foam_layers, fabric_layers, dealer_margin_percent):
    (product_rates, labour_rate, transport_rate, indirect_expense_rate, wastage_rate,
     margin_percent, tax_percent, working_cap_percent, pvc_packing_rate, flat_packing_cost) = read_price_sheet()

    surface_area = length * width
    quilting_thickness = 1
    quilting_rate = product_rates["Quilting"]

    total_core_cost = sum(surface_area * thickness * product_rates[material] for material, thickness in core_layers)
    total_core_thickness = sum(thickness for _, thickness in core_layers)

    total_foam_cost = sum(surface_area * thickness * product_rates[material] * 2 for material, thickness in foam_layers if material.lower() != "none")
    total_foam_thickness = sum(thickness for material, thickness in foam_layers if material.lower() != "none")

    total_fabric_cost = sum(surface_area * thickness * product_rates[material] * 2 for material, thickness in fabric_layers)
    total_fabric_thickness = sum(thickness for _, thickness in fabric_layers)

    quilting_cost = surface_area * quilting_thickness * quilting_rate * 2
    pvc_cost = surface_area * pvc_packing_rate

    material_cost = total_core_cost + total_foam_cost + total_fabric_cost + quilting_cost + pvc_cost + flat_packing_cost
    total_thickness = total_core_thickness + total_foam_thickness + total_fabric_thickness + quilting_thickness

    total_cost = material_cost + (labour_rate * total_thickness * surface_area) + transport_rate
    total_cost += (indirect_expense_rate * material_cost / 100) + (wastage_rate * material_cost / 100)

    mrp = total_cost * (1 + margin_percent / 100)
    mrp += mrp * (tax_percent / 100)
    mrp += mrp * (working_cap_percent / 100)
    mrp = round(mrp, 2)

    dealer_price = mrp + (mrp * dealer_margin_percent / 100) if dealer_margin_percent else mrp
    dealer_price = round(dealer_price, 2)

    return mrp, dealer_price

# ------------------ PDF Bill Export ------------------

def export_pdf(details, mrp, dealer_price):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("helvetica", size=12)

    pdf.cell(200, 10, text="Mattress Bill", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
    pdf.ln(10)

    for key, value in details.items():
        pdf.cell(200, 10, text=f"{key}: {value}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.ln(5)
    pdf.cell(200, 10, text=f"Total MRP: Rs.{mrp}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    #pdf.cell(200, 10, text=f"Dealer Price: Rs.{dealer_price}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    output_path = r"D:\\Internship\\Mattress_Bill.pdf"
    pdf.output(output_path)
    os.startfile(output_path)

# ------------------ Tkinter GUI ------------------

def run_gui():
    product_rates, *_ = read_price_sheet()

    root = tk.Tk()
    root.title("Mattress Bill Generator")
    root.geometry("500x500")

    canvas = Canvas(root)
    scrollbar = Scrollbar(root, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)

    scrollbar.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    scrollable_frame = Frame(canvas)
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

    tk.Label(scrollable_frame, text="Length (inches):").grid(row=0, column=0)
    length_entry = tk.Entry(scrollable_frame)
    length_entry.grid(row=0, column=1)

    tk.Label(scrollable_frame, text="Width (inches):").grid(row=1, column=0)
    width_entry = tk.Entry(scrollable_frame)
    width_entry.grid(row=1, column=1)

    #tk.Label(scrollable_frame, text="Dealer Margin %:").grid(row=2, column=0)
    #dealer_margin_entry = tk.Entry(scrollable_frame)
    #dealer_margin_entry.grid(row=2, column=1)

    material_options = [m for m in product_rates.keys() if m not in ["PVC Packing", "Thread, Cornershoe, Label", "Quilting"]]

    layers = {"Core": [], "Foam": [], "Fabric": []}
    row_counter = 3

    def add_layer(layer_type):
        nonlocal row_counter
        material_var = tk.StringVar()
        material_dropdown = ttk.Combobox(scrollable_frame, textvariable=material_var, values=material_options, state="readonly")
        material_dropdown.grid(row=row_counter, column=0)
        thickness_entry = tk.Entry(scrollable_frame)
        thickness_entry.grid(row=row_counter, column=1)
        layers[layer_type].append((material_var, thickness_entry))
        row_counter += 1
        update_generate_button()

    generate_button = tk.Button(scrollable_frame, text="Generate Bill", bg="lightgreen")

    def update_generate_button():
        generate_button.grid(row=row_counter, column=0, columnspan=2)

    def generate_bill():
        try:
            length = float(length_entry.get())
            width = float(width_entry.get())
            dealer_margin = float(dealer_margin_entry.get() or 0)

            core_layers = [(var.get(), float(entry.get())) for var, entry in layers["Core"]]
            foam_layers = [(var.get(), float(entry.get())) for var, entry in layers["Foam"]]
            fabric_layers = [(var.get(), float(entry.get())) for var, entry in layers["Fabric"]]

            mrp, dealer_price = calculate_mrp(length, width, core_layers, foam_layers, fabric_layers, dealer_margin)

            details = {
                "Length": f"{length}\"",
                "Width": f"{width}\"",
                "Added Layers": str(core_layers),
                #"Foam Layers": str(foam_layers),
                #"Fabric Layers": str(fabric_layers)
            }

            messagebox.showinfo("Bill", f"MRP: Rs.{mrp}\nDealer Price: Rs.{dealer_price}")

            export_pdf(details, mrp, dealer_price)

        except Exception as e:
            messagebox.showerror("Error", f"Invalid input: {e}")

    generate_button.config(command=generate_bill)
    update_generate_button()

    tk.Button(scrollable_frame, text="Add Layer", command=lambda: add_layer("Core")).grid(row=3, column=2)
    #tk.Button(scrollable_frame, text="Add Foam Layer", command=lambda: add_layer("Foam")).grid(row=4, column=2)
    #tk.Button(scrollable_frame, text="Add Fabric Layer", command=lambda: add_layer("Fabric")).grid(row=5, column=2)

    root.mainloop()

run_gui()
