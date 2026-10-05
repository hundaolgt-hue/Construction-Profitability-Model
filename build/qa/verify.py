"""Independent hand-verification of a recalculated workbook (spec §12.3).

1. one loan repayment (PMT) and one loan year roll-forward
2. one tax year (O5)
3. balance sheet in O5 rebuilt from its components
4. Equity / project IRR vs an independent Python XIRR and numpy_financial.irr
"""
import sys
import datetime as dt
import numpy_financial as npf
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter as CL

path = sys.argv[1]
wb = load_workbook(path, data_only=True)
fwb = load_workbook(path.replace("_calc", "") if "_calc" in path else path)


def nm(n):
    d = wb.defined_names[n]
    sh, ref = list(d.destinations)[0]
    return wb[sh][ref.replace("$", "").split(":")[0]].value


def row_of(sheet, label_start, after=1):
    ws = wb[sheet]
    for r in range(after, ws.max_row + 1):
        v = ws.cell(r, 2).value
        if isinstance(v, str) and v.startswith(label_start):
            return r
    raise KeyError(label_start)


def val(sheet, label, col):
    return wb[sheet].cell(row_of(sheet, label), col).value or 0


ok = True


def check(name, a, b, tol=0.5):
    global ok
    good = abs((a or 0) - (b or 0)) <= tol
    ok &= good
    print(f"  [{'OK ' if good else 'BAD'}] {name}: model={a:,.4f}  independent={b:,.4f}")


# ---------------------------------------------------------------- 1. loan
print("1. Loan schedule")
ln = "27_Loan"
ws = wb[ln]
r_ev, r_open, r_rate, r_rem, r_prin, r_capg, r_bfv = (row_of(ln, s) for s in (
    "Repayment event flag", "Opening balance", "Interest rate (p.a.)", "Remaining payments", "Scheduled principal",
    "Interest capitalised – grace", "Balloon amount"))
freq = nm("FREQ_M")
first = None
for c in range(5, 5 + 288):
    if (ws.cell(r_ev, c).value or 0) == 1:
        first = c
        break
if first is None:
    print("  no scheduled repayment in this case")
else:
    opening = ws.cell(r_open, first).value
    rate = ws.cell(r_rate, first).value
    rem = ws.cell(r_rem, first).value
    rp = rate * freq / 12
    rtype = nm("REPAY_TYPE")
    bfv = ws.cell(r_bfv, first).value or 0
    if rtype == 2:
        indep = (opening - bfv) / rem
    else:
        indep = float(npf.pmt(rp, rem, -opening, bfv)) - opening * rp
    print(f"  first repayment month t={first - 4}, opening {opening:,.0f}, rate {rate:.4%}, payments left {rem}")
    check("scheduled principal (numpy_financial.pmt − interest)", ws.cell(r_prin, first).value, indep, 1)
    # one loan year roll-forward O3 (period 7 = column K)
    col = 11
    a0 = row_of(ln, "Annual roll-up")
    get = lambda lab: ws.cell(row_of(ln, lab, a0), col).value or 0
    opening_y = val(ln, "Opening loan balance", col)
    closing_y = val(ln, "Closing loan balance", col)
    roll = opening_y + get("Drawdowns") + get("IDC capitalised") + get("Interest capitalised (grace)") - get("Scheduled principal") - get("Cash sweep prepayment")
    check("loan roll-forward O3 (opening + draws + cap. interest − principal − sweep)", closing_y, roll, 1)
    # monthly IDC recursion check at month 6
    r_draw, r_int = row_of(ln, "Drawdowns"), row_of(ln, "Interest accrued")
    t = max(range(1, 49), key=lambda m: ws.cell(r_draw, 4 + m).value or 0)
    print(f"  IDC test month t={t}")
    o, d, rr = ws.cell(r_open, 4 + t).value, ws.cell(r_draw, 4 + t).value, ws.cell(r_rate, 4 + t).value
    check("IDC month t = (opening + draw) × rate / 12", ws.cell(r_int, 4 + t).value, (o + d) * rr / 12, 0.01)

# ---------------------------------------------------------------- 2. tax O5
print("2. Tax – year O5 (column M)")
col = 13
tx = "28_Tax"
pbt = val("40_PL", "Profit before tax", col)
adep = val("26_Assets", "Total depreciation", col)
tdb = val(tx, "Less: tax depreciation – buildings", col)
tdp = val(tx, "Less: tax depreciation – plant", col)
ti = pbt + adep - tdb - tdp
check("taxable income before losses", val(tx, "Taxable income before losses", col), ti)
base_str = nm("BASE_STR")
check("building allowance = 5% × structure base", tdb, min(nm("TAXDEP_BLD") * base_str, base_str - val(tx, "  cumulative", col - 1)))
usage = val(tx, "Losses used", col)
cit = nm("CIT") * (max(0, ti) - usage)
check("CIT = rate × (taxable − losses used)", val(tx, "Corporate income tax", col), cit)
patpre = pbt - cit - val(tx, "Deferred tax expense", col)
k = (1 - nm("LR_PCT")) * (1 - nm("DIV_PAYOUT"))
r_ = nm("UPT_R")
upt = r_ * k * patpre / (1 + r_ * k) if patpre > 0 else 0
check("UPT closed form r·k·PAT/(1+r·k)", val(tx, "UPT = r", col), upt)

# ---------------------------------------------------------------- 3. BS O5
print("3. Balance sheet – O5 rebuilt from components")
bs = "41_BS"
assets = sum(val(bs, lab, col) for lab in ("Investment property", "Property, plant", "Deferred tax asset", "Inventory",
                                             "Trade receivables", "VAT receivable", "Contract asset", "Restricted cash",
                                             "Cash and cash"))
liabs = sum(val(bs, lab, col) for lab in ("Bank loan", "Land lease liability", "Deferred tax liability", "Deferred income",
                                            "Retention payable", "Trade payables", "Contract liability", "Income tax payable"))
equity = sum(val(bs, lab, col) for lab in ("Share capital", "Legal reserve", "Retained earnings", "Non-controlling"))
check("assets = liabilities + equity", assets, liabs + equity, 1)
check("BS cash = CF closing cash", val(bs, "Cash and cash", col), val("42_CF", "CLOSING CASH", col), 0.5)

# ---------------------------------------------------------------- 4. IRR
print("4. Returns – independent IRR")
rt = "50_Returns"
ecf = [val(rt, "EQUITY CASH FLOW", c) for c in range(5, 19)]
pcf = [val(rt, "PROJECT CASH FLOW", c) for c in range(5, 19)]
dates = [wb[rt].cell(row_of(rt, "Cash flow date"), c).value for c in range(5, 19)]
d0 = dates[0]


def xnpv(r, cfs):
    return sum(cf / (1 + r) ** ((d - d0).days / 365.0) for cf, d in zip(cfs, dates))


def xirr(cfs):
    lo, hi = -0.99, 10.0
    if xnpv(lo, cfs) * xnpv(hi, cfs) > 0:
        return 0.0
    for _ in range(300):
        mid = (lo + hi) / 2
        if xnpv(lo, cfs) * xnpv(mid, cfs) <= 0:
            hi = mid
        else:
            lo = mid
    return (lo + hi) / 2


check("Equity IRR (XIRR) vs Python bisection XIRR", nm("EIRR"), xirr(ecf), 0.001)
check("Project IRR (XIRR) vs Python bisection XIRR", nm("PIRR"), xirr(pcf), 0.001)
check("Per-period project IRR vs numpy_financial.irr", nm("PIRR_PER"), float(npf.irr(pcf)), 0.001)
dfs = [(1 + nm("DISC_EFF")) ** (-((d - nm("START_DATE")).days + 1) / 365) for d in dates]
check("Project NPV vs Python discounting", nm("PNPV"), sum(c * f for c, f in zip(pcf, dfs)), 1)
print("ALL VERIFIED" if ok else "SOME CHECKS FAILED")
