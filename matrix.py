import itertools
import pandas as pd

# Define product options
coir_core = ["Coir 80D", "Coir 90D", "Coir 100D"]
latex_foam_core = ["Topper", "Natural Latex", "Memory foam", "Srilanka Latex Rebond", "Foam - Rebonded"]
spring_core = ["Bonnel (only 5) Spring", "Pocketed (only 5) Spring"]
foam_core = ["EP Foam", "PU Foam"]

foam_options = ["None", "Single Foam", "Single foam + Single foam", "Double foam + double foam"]
fabric_options = ["Fabric Regular (120 GSM)", "Fabric Premium (250 GSM)", "Fabric Ultra Premium (350 GSM)"]

compulsory = "Quilting"

# Mattress sizes
length_options = [72, 75]
width_options = [30, 36, 42, 48, 60]
thickness_options = [4, 5, 6, 8]

# 1. Generate all mattress combinations (Core + Foam + Fabric)
mattress_combinations = []

# Coir Only
for core, fabric in itertools.product(coir_core, fabric_options):
    combo = f"{core} | None | {fabric} | {compulsory}"
    mattress_combinations.append(combo)

# Coir + Foam
for core, foam, fabric in itertools.product(coir_core, foam_options[1:], fabric_options):
    combo = f"{core} | {foam} | {fabric} | {compulsory}"
    mattress_combinations.append(combo)

# Latex/Foam Rebond Core with Optional Foam
for core, foam, fabric in itertools.product(latex_foam_core, foam_options, fabric_options):
    combo = f"{core} | {foam} | {fabric} | {compulsory}"
    mattress_combinations.append(combo)

# Spring Mattress with Optional Foam
for core, foam, fabric in itertools.product(spring_core, foam_options, fabric_options):
    combo = f"{core} | {foam} | {fabric} | {compulsory}"
    mattress_combinations.append(combo)

# PU/EP Foam Core with Optional Foam
for core, foam, fabric in itertools.product(foam_core, foam_options, fabric_options):
    combo = f"{core} | {foam} | {fabric} | {compulsory}"
    mattress_combinations.append(combo)

# 2. Generate all size combinations
size_combinations = list(itertools.product(length_options, width_options, thickness_options))

# 3. Create final DataFrame
columns = ["Length", "Width", "Thickness"] + mattress_combinations

# Initialize blank data
data = []

for size in size_combinations:
    row = list(size) + [""] * len(mattress_combinations)
    data.append(row)

df = pd.DataFrame(data, columns=columns)

# 4. Save to Excel
output_path = r"D:\Internship\Mattress_Matrix.xlsx"
df.to_excel(output_path, index=False)

print(f"Excel generated with {len(size_combinations)} rows and {len(mattress_combinations)} mattress combinations.")
print(f"Saved to: {output_path}")
