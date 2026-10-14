"""Build the Rapha Development Profitability Model workbook (private + commercial tracks, Google-Sheets-ready).

usage: python main.py [--blank] [--out path]
"""
import argparse
from openpyxl.workbook.properties import CalcProperties
from core import Builder, header
import engine
import fin
import front
import private

ORDER = ["00_Start", "P1_Inputs", "P2_Calc", "P3_Estimate", "P4_Checks",
         "01_Dashboard", "10_Inputs", "11_Scenarios", "12_TaxLaw", "13_Spaces", "20_Timeline", "21_Area",
         "22_Capex", "23_Construction", "24_Revenue", "25_Opex", "26_Assets", "27_Loan", "28_Tax", "29_VAT",
         "30_WorkingCap", "40_PL", "41_BS", "42_CF", "43_Ratios", "50_Returns", "51_Payback", "52_Sensitivity",
         "60_Checks", "61_Glossary", "Lists"]
VIS = {"00_Start": 11, "01_Dashboard": 19, "10_Inputs": 12, "11_Scenarios": 10, "12_TaxLaw": 7, "13_Spaces": 19,
       "61_Glossary": 4, "22_Capex": 12, "52_Sensitivity": 18, "60_Checks": 6, "21_Area": 20,
       "P1_Inputs": 9, "P2_Calc": 16, "P3_Estimate": 11, "P4_Checks": 6}


def build(blank=False):
    b = Builder()
    front.build_lists(b)
    front.build_inputs(b, blank)
    private.build_p_inputs(b, blank)
    front.build_scenarios(b)
    front.build_taxlaw(b)
    front.build_spaces(b, blank)
    engine.build_engine(b)
    engine.build_engine2(b)
    fin.build_fin(b)
    fin.build_sens(b)
    fin.build_checks(b)
    private.build_p_calc(b)
    private.build_p_estimate(b)
    private.build_p_checks(b)
    front.build_glossary(b)
    front.build_start(b, blank)
    front.build_dashboard(b)
    for name, sh in b.sheets.items():
        header(b, sh, ncols_vis=VIS.get(name))
    for n in list(b.sheets):
        b.sheets[n].render()
    front.add_validations(b)
    for n, end in (("27_Loan", b.LOAN_PRINT_END), ("20_Timeline", b.TL_PRINT_END)):
        ws = b.sheets[n].ws
        ws.print_area = f"A1:R{end}"
        ws.print_title_cols = None
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
    ws = b.sheets["P2_Calc"].ws
    ws.print_area = f"A1:R{b.sheets['P2_Calc'].r + 2}"
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
