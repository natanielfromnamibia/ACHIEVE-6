import csv
import json
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# Load rules
with open("rules.json", "r", encoding="utf-8") as f:
    R = json.load(f)

# Read input.csv (we change input.xlsx to csv – easier with csv module)
with open("input.csv", "r", encoding="utf-8") as f:
    reader = csv.DictReader(f)
    rows = list(reader)
    r = rows[0]  # first row only

score = 0

# Income
for lim, pts in R["scoring"]["income"]:
    if float(r["monthly_income"]) >= lim:
        score += pts
        break

# Years
for lim, pts in R["scoring"]["years"]:
    if float(r["employment_years"]) >= lim:
        score += pts
        break

# Age
for lo, hi, pts in R["scoring"]["age"]:
    if lo <= int(r["age"]) <= hi:
        score += pts
        break

# DTI
for lim, pts in R["scoring"]["dti"]:
    if float(r["debt_to_income"]) <= lim:
        score += pts
        break

# Defaults
if int(r["defaults_last_24m"]) == 0:
    score += 20

# Color
color = R["colors"]["decline"]
if score >= 85: color = R["colors"]["approve"]
elif score >= 65: color = R["colors"]["review"]
elif score >= 45: color = R["colors"]["caution"]

# Save to Excel
out = "ACHIEVE6_RESULT.xlsx"
wb = load_workbook("input.xlsx")  # open original for formatting
ws = wb.active
ws["H2"] = score  # FINAL_SCORE in column H row 2

fill = PatternFill("solid", fgColor=color)
for cell in ws[ws.max_row]:
    cell.fill = fill

wb.save(out)

print(f"ACHIEVE-6 → SCORE {score} → {out} READY")