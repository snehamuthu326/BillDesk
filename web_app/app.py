from flask import Flask, render_template, request, redirect, send_file
import pandas as pd
import itertools
import os

app = Flask(__name__)

# ----------------- Read Price Sheet -----------------
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
quilting_thickness = 1

custom_products_list = []
custom_sizes = []

length_options = [72, 75, 78, 84]
width_options = [30, 36, 42, 44, 48, 60, 72]

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        material = request.form.get("material")
        thickness = request.form.get("thickness")
        length = request.form.get("length")
        width = request.form.get("width")

        if material and thickness:
            custom_products_list.append([(material, float(thickness))])

        if length and width:
            custom_sizes.append((float(length), float(width)))

        return redirect("/")

    return render_template("index.html",
                           product_rates=product_rates,
                           custom_products=custom_products_list,
                           custom_sizes=custom_sizes)

@app.route("/generate")
def generate():
    if not custom_products_list:
        return "Please add at least one custom product.", 400

    columns = ["Length", "Width"]
    for prod in custom_products_list:
        desc = " + ".join([f"{mat} {thk}\"" for mat, thk in prod])
        columns += [f"{desc} | Net Rate", f"{desc} | MRP"]

    all_lengths = sorted(set(length_options + [l for l, _ in custom_sizes]))
    all_widths = sorted(set(width_options + [w for _, w in custom_sizes]))
    data = []

    for length, width in itertools.product(all_lengths, all_widths):
        row = [length, width]
        for prod in custom_products_list:
            mrp, dealer_price = calculate_total_cost(length, width, prod)
            row += [mrp, dealer_price]
        data.append(row)

    pd.DataFrame(data, columns=columns).to_excel(output_path, index=False)
    return send_file(output_path, as_attachment=True)

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

if __name__ == "__main__":
    app.run(debug=True)
