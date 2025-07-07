from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
from fpdf import FPDF
import os

app = Flask(__name__)

# ---------------- Read Price Sheet ----------------

def read_price_sheet():
    df = pd.read_excel("Costing_Sheet.xlsx", sheet_name="new")
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

# ---------------- MRP Calculation ----------------

def calculate_mrp(length, width, material, thickness, dealer_margin_percent):
    (product_rates, labour_rate, transport_rate, indirect_expense_rate, wastage_rate,
     margin_percent, tax_percent, working_cap_percent, pvc_packing_rate, flat_packing_cost) = read_price_sheet()

    surface_area = length * width
    quilting_thickness = 1
    quilting_rate = product_rates["Quilting"]

    material_cost = surface_area * thickness * product_rates[material]
    quilting_cost = surface_area * quilting_thickness * quilting_rate * 2
    pvc_cost = surface_area * pvc_packing_rate

    total_material_cost = material_cost + quilting_cost + pvc_cost + flat_packing_cost
    total_thickness = thickness + quilting_thickness

    total_cost = total_material_cost + (labour_rate * total_thickness * surface_area) + transport_rate
    total_cost += (indirect_expense_rate * total_material_cost / 100) + (wastage_rate * total_material_cost / 100)

    mrp = total_cost * (1 + margin_percent / 100)
    mrp += mrp * (tax_percent / 100)
    mrp += mrp * (working_cap_percent / 100)
    mrp = round(mrp, 2)

    dealer_price = mrp + (mrp * dealer_margin_percent / 100) if dealer_margin_percent else mrp
    dealer_price = round(dealer_price, 2)

    return mrp, dealer_price

# ---------------- PDF Export ----------------

def export_pdf(details, mrp):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("helvetica", size=12)

    pdf.cell(200, 10, text="Mattress Bill", ln=True, align="C")
    pdf.ln(10)

    for key, value in details.items():
        pdf.cell(200, 10, txt=f"{key}: {value}", ln=True)

    pdf.ln(5)
    pdf.cell(200, 10, txt=f"Total MRP: Rs.{mrp}", ln=True)

    output_path = "static/Mattress_Bill.pdf"
    pdf.output(output_path)
    return output_path

# ---------------- Routes ----------------

@app.route('/')
def index():
    product_rates, *_ = read_price_sheet()
    material_options = [m for m in product_rates.keys() if m not in ["PVC Packing", "Thread, Cornershoe, Label", "Quilting"]]
    return render_template("home.html", materials=material_options)

@app.route('/calculate', methods=['POST'])
def calculate():
    data = request.json
    length = float(data['length'])
    width = float(data['width'])
    material = data['material']
    thickness = float(data['thickness'])
    dealer_margin = float(data['dealer_margin'])

    mrp, dealer_price = calculate_mrp(length, width, material, thickness, dealer_margin)

    details = {
        "Length": f"{length}\"",
        "Width": f"{width}\"",
        "Material": material,
        "Thickness": f"{thickness}\""
    }

    pdf_path = export_pdf(details, mrp)

    return jsonify({"mrp": mrp, "dealer_price": dealer_price, "pdf_path": pdf_path})

# ---------------- Run ----------------

if __name__ == '__main__':
    app.run(debug=True)
