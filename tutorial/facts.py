"""Pull the numbers quoted in the tutorial straight from the finished workbooks (cached values)."""
from openpyxl import load_workbook


def _nm(wb, n):
    d = wb.defined_names[n]
    sh, ref = list(d.destinations)[0]
    return wb[sh][ref.replace("$", "").split(":")[0]].value


def grab(path, names):
    wb = load_workbook(path, data_only=True)
    out = {n: _nm(wb, n) for n in names}
    return out, wb


def etb(v, d=0):
    if v is None:
        return "–"
    return f"{v:,.{d}f}"


def m(v, d=1):
    return f"{v / 1e6:,.{d}f}"


def pct(v, d=1):
    return f"{v * 100:.{d}f}%"


A_NAMES = ["TDC", "HOME_TOTAL", "CONSTR_COST", "CC_M2", "TDC_M2G", "GFA", "LN_INSTAL_M", "DEBT_DRAWN", "LN_INT_TOT", "LN_FINAL_DATE",
           "AFFORD_PCT", "FUND_GAP", "SAVINGS", "FAMILY", "HOME_LOAN", "RATE_EFF", "TENOR_Y", "EST_LOW", "EST_HIGH", "HARD_BASE",
           "ESC_AMT", "CONT_TOT", "DESIGN_TOT", "SUPERV_TOT", "PERMIT_TOT", "FFE_TOT", "IDC_TOT", "CONSTR_EFF", "HH_INCOME",
           "RET_TOT", "PEAK_DEBT", "ARR_TOT", "STAGE_TOT", "CHK_PASS", "CHK_TOTAL"]
B_NAMES = ["TDC", "CONSTR_COST", "LAND_TOT", "IDC_TOT", "FIN_CAP", "GDV", "POC", "EIRR", "PIRR", "PNPV", "ENPV", "EMOIC", "MIN_DSCR",
           "AVG_DSCR", "PEAK_DEBT", "VERDICT", "PB_P", "PB_E", "PB_PD", "PB_ED", "PRESOLD", "SALE_AREA", "LET_AREA", "GFA", "NET_AREA",
           "EQ_DRAWN", "DEBT_DRAWN", "PS_USED", "FUND_GAP", "TV_NET", "NOI_STAB", "YOC", "DEV_MARGIN", "LLCR_O1", "MIN_ICR",
           "CHK_PASS", "CHK_TOTAL", "BE_PRICE_NPV", "BE_PRICE_POC", "BE_OCC", "BE_COST", "BE_RATE", "BE_DELAY", "BLEND_PRICE",
           "HARD_BASE", "HARD_ESC_TOT", "CONT_TOT", "PEAK_FUND", "EQ_COMMIT", "DEBT_CAP", "HAND_TOT", "LN_INSTAL", "LN_INT_TOT",
           "CONSTR_EFF", "TDC_M2G", "SALES_NOM", "RET_VALUE", "PAT_TOT", "N_DS_YEARS", "RATE_EFF", "OCC_STAB", "HURDLE",
           "DSCR_COV", "CASH_MIN_OPS", "TDC_INV", "TDC_IP"]
