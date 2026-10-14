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
    return f"{(v or 0) / 1e6:,.{d}f}"


def pct(v, d=1):
    return f"{(v or 0) * 100:.{d}f}%"


C_NAMES = ["TDC", "GDV", "POC", "EIRR", "PIRR", "PNPV", "EMOIC", "MIN_DSCR", "PEAK_DEBT", "VERDICT", "NET_AREA", "GFA", "SALE_AREA",
           "LET_AREA", "OWN_AREA", "EQ_DRAWN", "DEBT_DRAWN", "PS_USED", "IDC_TOT", "NOI_STAB", "YOC", "BE_PRICE_NPV", "BE_COST",
           "BE_RENT_NPV", "BE_OCC", "BE_RATE", "BE_DELAY", "BE_CAP", "BLEND_PRICE", "PB_P", "PB_E", "CHK_PASS", "CHK_TOTAL", "LN_INSTAL",
           "LN_FINAL_DATE", "CONSTR_EFF", "HARD_BASE", "CONSTR_COST", "LAND_TOT", "FIN_CAP", "TDC_M2G", "PREOP_TOT", "CONSULT_TOT",
           "HOTEL_RENT_TODAY", "HOTEL_AREA", "NCI_PCT", "OCC_STAB", "GPR_TODAY", "RENT_VAT_PCT", "TOP_RISK", "PRESOLD", "NTG",
           "SALES_NOM", "RET_VALUE", "DEV_MARGIN", "ENPV", "INV_SH", "IP_SH", "PPE_SH", "N_LOWCONF", "TV_NET", "RATE_EFF", "AVG_DSCR",
           "LLCR_O1", "PAT_TOT", "EQ_COMMIT", "DEBT_CAP", "PEAK_FUND", "HURDLE", "DSCR_COV", "EXIT_CAP_EFF", "PV_SALESNET", "PV_SPRENT",
           "PV_HOTRENT", "PV_CAPX", "DELAY6_PV", "FLEX_ON"]
P_NAMES = ["P_GFA", "P_BUILD_BASE", "P_BUILD_ESC", "P_CONT_TOT", "P_CONSTR", "P_TOTAL", "P_CC_M2", "P_LOW", "P_HIGH", "P_SAV_USED",
           "P_FAM_USED", "P_LOAN_DRAWN", "P_GAP", "P_B0", "P_INSTAL", "P_IDC_TOT", "P_INT_ALL", "P_PAYOFF", "P_DTI", "P_AFF_TXT",
           "P_LOAN_AFF", "P_RENT_Y1", "P_RENT_NET", "P_RENT_YIELD", "P_RENT_PB", "P_RENT_COVER", "P_PEAK_M", "P_RET_TOT", "P_NPAY",
           "P_MONTHS", "PCHK_PASS", "PCHK_TOTAL", "P_RATE", "P_TENOR", "P_STAGE_TOT", "P_DES_TOT", "P_SUP_TOT", "P_PER_TOT", "P_FUR_TOT",
           "P_ARR_TOT", "P_INCOME", "P_AFF_MAX", "P_SAVINGS", "P_FAMILY", "P_LOAN_MAX", "P_ESC", "P_RENT_SH", "P_RENT_PARTCOST"]
