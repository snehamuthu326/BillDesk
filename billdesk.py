import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd

file_path = "Costing_Sheet.xlsx"
df = pd.read_excel(file_path, sheet_name="Working 25.11.2023", header=None)

core_costs = {}
for i in range(4, 15):
    material = df.iloc[i, 1]
    cost = df.iloc[i, 3]
    if pd.notna(material) and pd.notna(cost):
        core_costs[material.strip()] = float(cost)
