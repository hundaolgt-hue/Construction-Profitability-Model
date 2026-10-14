"""Google Sheets compatibility audit.

Flags: HYPERLINK() formulas, Data Tables, conditional-format rules that use named ranges or other sheets,
data validations sourced from names, defined names that look like cell references, unsupported functions."""
import re
import sys
from openpyxl import load_workbook
from openpyxl.worksheet.formula import DataTableFormula

wb = load_workbook(sys.argv[1])
names = set(wb.defined_names.keys())
issues = []
FUNCS_BAD = ("TABLE(", "GETPIVOTDATA(", "CELL(", "INFO(", "INDIRECT(")
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            if isinstance(v, DataTableFormula):
                issues.append(f"{ws.title}!{c.coordinate} Data Table")
            elif isinstance(v, str) and v.startswith("="):
                if "HYPERLINK(" in v.upper():
                    issues.append(f"{ws.title}!{c.coordinate} HYPERLINK formula")
                for fb in FUNCS_BAD:
                    if fb in v.upper():
                        issues.append(f"{ws.title}!{c.coordinate} uses {fb}")
    for rng, rules in ws.conditional_formatting._cf_rules.items():
        for rule in rules:
            for f in rule.formula or []:
                body = re.sub(r'"[^"]*"', "", f)
                toks = set(re.findall(r"[A-Za-z_][A-Za-z0-9_]*", body)) - {"AND", "OR", "NOT", "LEFT", "COLUMN", "ROW", "IF"}
                used = [t for t in toks if t in names]
                if used or "!" in body:
                    issues.append(f"{ws.title} CF {rng.sqref}: {f}")
    for dv in ws.data_validations.dataValidation:
        if dv.type == "list" and dv.formula1 and "!" not in dv.formula1 and not dv.formula1.startswith('"'):
            issues.append(f"{ws.title} DV {dv.sqref}: {dv.formula1}")
        if dv.formula1 and str(dv.formula1).startswith("="):
            issues.append(f"{ws.title} DV {dv.sqref}: list source must not start with '=' ({dv.formula1})")
for n in names:
    if re.fullmatch(r"[A-Za-z]{1,3}\d+(_.*)?", n) or re.fullmatch(r"[Rr]\d*[Cc]\d*", n):
        issues.append(f"name looks like a cell reference: {n}")
links = sum(1 for ws in wb.worksheets for row in ws.iter_rows() for c in row if c.hyperlink is not None and c.hyperlink.location)
print(f"static hyperlinks: {links}")
print("\n".join(issues) if issues else "NO GOOGLE SHEETS ISSUES")
