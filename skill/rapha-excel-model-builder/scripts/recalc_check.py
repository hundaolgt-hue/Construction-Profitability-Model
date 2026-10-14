"""Recalculate an .xlsx with LibreOffice (via the xlsx skill's recalc.py) and print errors + named outputs.

usage: python recalc_check.py model.xlsx [NAME1,NAME2,...]
Writes model_calc.xlsx next to the input; never modifies the input.
"""
import glob
import json
import shutil
import subprocess
import sys

from openpyxl import load_workbook

src = sys.argv[1]
names = sys.argv[2].split(",") if len(sys.argv) > 2 else []
calc = src.replace(".xlsx", "_calc.xlsx")
shutil.copy(src, calc)
recalc = glob.glob("/root/.claude/skills/**/xlsx/scripts/recalc.py", recursive=True)
if not recalc:
    sys.exit("xlsx skill recalc.py not found – install libreoffice-calc and the xlsx skill")
out = subprocess.run([sys.executable, recalc[0], calc, "300"], capture_output=True, text=True)
print(out.stdout.strip())
wb = load_workbook(calc, data_only=True)
for n in names:
    d = wb.defined_names[n]
    sh, ref = list(d.destinations)[0]
    print(f"{n:16s}", wb[sh][ref.replace('$', '').split(':')[0]].value)
