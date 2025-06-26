import itertools
import pandas as pd

# Define the options
coir_core = ["Coir 80D", "Coir 90D", "Coir 100D"]
latex_foam_core = ["Topper", "Natural Latex", "Memory foam", "Srilanka Latex Rebond", "Foam - Rebonded"]
spring_core = ["Bonnel (only 5) Spring", "Pocketed (only 5) Spring"]
foam_core = ["EP Foam", "PU Foam"]

foam_options = ["None", "Single Foam", "Single foam + Single foam", "Double foam + double foam"]
fabric_options = ["Fabric Regular (120 GSM)", "Fabric Premium (250 GSM)", "Fabric Ultra Premium (350 GSM)"]

compulsory = "Quilting"

# Store all combinations
combinations = []

# 1. Coir Only
for core, fabric in itertools.product(coir_core, fabric_options):
    combinations.append([core, "None", fabric, compulsory])

# 2. Coir + Foam
for core, foam, fabric in itertools.product(coir_core, foam_options[1:], fabric_options):
    combinations.append([core, foam, fabric, compulsory])

# 3. Latex / Foam Rebond Core with Optional Foam
for core, foam, fabric in itertools.product(latex_foam_core, foam_options, fabric_options):
    combinations.append([core, foam, fabric, compulsory])

# 4. Spring Mattress with Optional Foam
for core, foam, fabric in itertools.product(spring_core, foam_options, fabric_options):
    combinations.append([core, foam, fabric, compulsory])

# 5. PU Foam / EP Foam Core with Optional Foam
for core, foam, fabric in itertools.product(foam_core, foam_options, fabric_options):
    combinations.append([core, foam, fabric, compulsory])

# Convert to DataFrame
df = pd.DataFrame(combinations, columns=["Core", "Foam", "Fabric", "Compulsory"])

# Export to Excel
output_path = r"D:\Internship\Mattress_Combinations.xlsx"
df.to_excel(output_path, index=False)

print(f"Total combinations generated: {len(combinations)}")
print(f"File saved at: {output_path}")
