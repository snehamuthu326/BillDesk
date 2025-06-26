import itertools
import pandas as pd

# ------------------ MATERIAL OPTIONS WITH THICKNESS ------------------

# Core materials with thickness variants
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

# Foam options with thickness
foam_options = [
    ("None", 0),
    ("Single Foam", 1),
    ("Single foam + Single foam", 2),
    ("Double foam + double foam", 4)
]

# Fabric options with thickness (generally fixed)
fabric_options = [
    ("Fabric Regular (120 GSM)", 0.5),
    ("Fabric Premium (250 GSM)", 0.5),
    ("Fabric Ultra Premium (350 GSM)", 0.5)
]

# Compulsory quilting thickness
compulsory_thickness = 0.5

# ------------------ SIZE OPTIONS ------------------

length_options = [72, 75]
width_options = [30, 36, 42, 48, 60]

# ------------------ GENERATE MATERIAL COMBINATIONS ------------------

mattress_combinations = []

for core, core_thick in core_options:
    
    # 1. Coir Only if applicable
    if "Coir" in core:
        for fabric, fabric_thick in fabric_options:
            total_thickness = core_thick + fabric_thick + compulsory_thickness
            combo_name = f"{core} {core_thick}\" | None | {fabric} | Total {total_thickness}\""
            mattress_combinations.append(combo_name)

        # 2. Coir + Foam
        for foam, foam_thick in foam_options[1:]:  # Exclude "None"
            for fabric, fabric_thick in fabric_options:
                total_thickness = core_thick + foam_thick + fabric_thick + compulsory_thickness
                combo_name = f"{core} {core_thick}\" | {foam} {foam_thick}\" | {fabric} | Total {total_thickness}\""
                mattress_combinations.append(combo_name)

    # 3. Latex/Foam Rebond Core with Optional Foam
    elif core in ["Topper", "Natural Latex", "Memory foam", "Srilanka Latex Rebond", "Foam - Rebonded"]:
        for foam, foam_thick in foam_options:
            for fabric, fabric_thick in fabric_options:
                total_thickness = core_thick + foam_thick + fabric_thick + compulsory_thickness
                combo_name = f"{core} {core_thick}\" | {foam} {foam_thick}\" | {fabric} | Total {total_thickness}\""
                mattress_combinations.append(combo_name)

    # 4. Spring Core with Optional Foam
    elif "Spring" in core:
        for foam, foam_thick in foam_options:
            for fabric, fabric_thick in fabric_options:
                total_thickness = core_thick + foam_thick + fabric_thick + compulsory_thickness
                combo_name = f"{core} {core_thick}\" | {foam} {foam_thick}\" | {fabric} | Total {total_thickness}\""
                mattress_combinations.append(combo_name)

    # 5. Foam Core with Optional Foam
    elif core in ["EP Foam", "PU Foam"]:
        for foam, foam_thick in foam_options:
            for fabric, fabric_thick in fabric_options:
                total_thickness = core_thick + foam_thick + fabric_thick + compulsory_thickness
                combo_name = f"{core} {core_thick}\" | {foam} {foam_thick}\" | {fabric} | Total {total_thickness}\""
                mattress_combinations.append(combo_name)

# ------------------ SIZE COMBINATIONS ------------------

size_combinations = list(itertools.product(length_options, width_options))

# ------------------ FINAL MATRIX GENERATION ------------------

columns = ["Length", "Width"] + mattress_combinations
data = []

for length, width in size_combinations:
    row = [length, width] + [""] * len(mattress_combinations)
    data.append(row)

df = pd.DataFrame(data, columns=columns)

# ------------------ EXPORT TO EXCEL ------------------

output_path = r"D:\Internship\Mattress_Matrix_With_Thickness_Corrected.xlsx"
df.to_excel(output_path, index=False)

print(f"Matrix generated with {len(size_combinations)} size rows and {len(mattress_combinations)} combinations.")
print(f"Saved to: {output_path}")
