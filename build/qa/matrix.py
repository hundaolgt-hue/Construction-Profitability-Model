"""Run the model across tracks, project types, data sets and scenarios; report errors, checks and key results."""
import sys
from harness import run_many, values, nm
src = sys.argv[1]
which = sys.argv[2] if len(sys.argv) > 2 else "full"
specs = []
if which in ("full", "modes"):
    for d in (2, 3):
        for m in (1, 2, 3):
            specs.append((f"d{d}_m{m}_s1", dict(track=2, mode=m, data=d, scen=1)))
if which in ("full", "scen"):
    for d, m in ((2, 3), (3, 2)):
        for s in (2, 3, 4):
            specs.append((f"d{d}_m{m}_s{s}", dict(track=2, mode=m, data=d, scen=s)))
if which in ("full", "private"):
    for pd in (1, 2):
        specs.append((f"priv_p{pd}", dict(track=1, pdata=pd)))
if which == "blank":
    for m in (1, 2, 3):
        for s in (1, 3):
            specs.append((f"blank_m{m}_s{s}", dict(track=2, mode=m, data=1, scen=s)))
    specs.append(("blank_priv", dict(track=1, pdata=1)))
res = run_many(src, specs, workers=4, keep_dir=f"/tmp/claude-0/qa_{which}")
keys = ["TDC", "FUND_GAP", "PEAK_DEBT", "EIRR", "PIRR", "POC", "MIN_DSCR", "NCI_PCT", "VERDICT", "CHK_PASS", "CHK_FAIL", "CHK_WARN",
        "P_TOTAL", "P_INSTAL", "PCHK_PASS", "PCHK_FAIL", "PCHK_WARN"]


def f(v):
    return f"{v:,.3f}" if isinstance(v, float) and abs(v) < 10 else (f"{v:,.0f}" if isinstance(v, (int, float)) else str(v))


nfail = 0
for tag, (r, p) in sorted(res.items()):
    if "error" in r:
        print(tag, "RECALC ERROR", r["error"])
        continue
    wb = values(p)
    vals = {k: nm(wb, k) for k in keys}
    flags = []
    for sname in ("60_Checks", "P4_Checks"):
        ws = wb[sname]
        for row in range(5, 50):
            st = ws.cell(row, 6).value
            if st in ("FAIL", "WARN"):
                v = ws.cell(row, 4).value
                flags.append(f"{st}:{ws.cell(row, 2).value[:3]}" + (f"={v:.4g}" if isinstance(v, (int, float)) else ""))
                nfail += st == "FAIL"
    print(f"{tag:12} err={r.get('total_errors')} " + " ".join(f"{k}={f(v)}" for k, v in vals.items()), flags)
print("TOTAL FAIL:", nfail)
