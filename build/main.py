"""Build the Rapha Development Profitability Model workbook.

usage: python main.py [--blank] [--out path]
"""
import sys
import argparse
from openpyxl.formatting.rule import Rule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles.numbers import NumberFormat
from openpyxl.workbook.properties import CalcProperties
from core import Builder, header, CL
import engine
import fin
import front

ORDER = ["00_Start", "01_Dashboard", "02_Home", "10_Inputs", "11_Scenarios", "12_TaxLaw", "20_Timeline", "21_Area",
         "22_Capex", "23_Construction", "24_Revenue", "25_Opex", "26_Assets", "27_Loan", "28_Tax", "29_VAT",
         "30_WorkingCap", "40_PL", "41_BS", "42_CF", "43_Ratios", "50_Returns", "51_Payback", "52_Sensitivity",
         "60_Checks", "61_Glossary", "Lists"]
VIS = {"00_Start": 13, "01_Dashboard": 19, "02_Home": 9, "10_Inputs": 10, "11_Scenarios": 10, "12_TaxLaw": 7,
       "61_Glossary": 4, "22_Capex": 12, "52_Sensitivity": 16, "60_Checks": 6, "21_Area": 20}


def build(blank=False):
    b = Builder()
    front.build_lists(b)
    inp = front.build_inputs(b, blank)
    sc = front.build_scenarios(b)
    tl_ = front.build_taxlaw(b)
    engine.build_engine(b)
    engine.build_engine2(b)
    fin.build_fin(b)
    fin.build_sens(b)
    fin.build_checks(b)
    front.build_glossary(b)
    front.build_start(b, blank)
    front.build_dashboard(b)
    front.build_home(b)
    # render: headers first, then content (content of front sheets depends on rows reserved elsewhere)
    for name, sh in b.sheets.items():
        header(b, sh, ncols_vis=VIS.get(name))
    render_order = list(b.sheets)
    for n in render_order:
        b.sheets[n].render()
    front.add_validations(b)
    # unit toggle on statements: base format = ETB '000 ; conditional = ETB / ETB m
    for n in ("40_PL", "41_BS", "42_CF", "43_Ratios", "50_Returns", "51_Payback"):
        sh = b.sheets[n]
        krows = [it[1].r for it in sh.items if it[0] == "line" and it[1].fmt == "k"]
        if not krows:
            continue
        rng = " ".join(f"D{r}:R{r}" for r in krows)
        for cond, code, fid in (("UNIT_SEL=1", '#,##0;(#,##0);"–"', 201), ("UNIT_SEL=3", '#,##0.0,,;(#,##0.0,,);"–"', 203)):
            dxf = DifferentialStyle(numFmt=NumberFormat(numFmtId=fid, formatCode=code))
            sh.ws.conditional_formatting.add(rng, Rule(type="expression", dxf=dxf, formula=[cond]))
    for n, end in (("27_Loan", b.LOAN_PRINT_END), ("20_Timeline", b.TL_PRINT_END)):
        ws = b.sheets[n].ws
        ws.print_area = f"A1:R{end}"
        ws.print_title_cols = None
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
    b.finalize_names()
    b.wb._sheets = [b.wb[n] for n in ORDER]
    b.wb.active = 0
    for ws in b.wb.worksheets:
        ws.sheet_view.tabSelected = ws.title == "00_Start"
    b.wb.calculation = CalcProperties(fullCalcOnLoad=True)
    return b


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--blank", action="store_true")
    ap.add_argument("--out", default="../output/model.xlsx")
    a = ap.parse_args()
    b = build(a.blank)
    b.wb.save(a.out)
    print("saved", a.out, "names:", len(b.names))
