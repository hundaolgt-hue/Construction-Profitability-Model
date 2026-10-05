import sys, json
from harness import run_many, values, nm
src = sys.argv[1]
which = sys.argv[2] if len(sys.argv) > 2 else "full"
specs = []
if which in ("full", "modes"):
    for d in (2, 3):
        for m in range(4):
            specs.append((f"d{d}_m{m}_s1", dict(mode=m, data=d, scen=1)))
if which in ("full", "scen"):
    for d in (2, 3):
        for s in (2, 3, 4):
            specs.append((f"d{d}_m{3 if d==3 else 0}_s{s}", dict(mode=3 if d == 3 else 0, data=d, scen=s)))
if which == "blank":
    for m in range(4):
        for s in (1, 3):
            specs.append((f"blank_m{m}_s{s}", dict(mode=m, data=1, scen=s)))
res = run_many(src, specs, workers=4, keep_dir=f"/tmp/claude-0/qa_{which}")
keys = ["TDC", "HOME_TOTAL", "CONSTR_COST", "FUND_GAP", "LN_INSTAL_M", "PEAK_DEBT", "EIRR", "PIRR", "POC", "MIN_DSCR", "VERDICT", "CHK_PASS", "CHK_FAIL", "CHK_WARN"]
for tag, (r, p) in sorted(res.items()):
    if "error" in r:
        print(tag, "RECALC ERROR", r["error"]); continue
    wb = values(p)
    vals = {k: nm(wb, k) for k in keys}
    fails = []
    ws = wb["60_Checks"]
    for row in range(9, 45):
        st = ws.cell(row, 6).value
        if st in ("FAIL", "WARN"):
            fails.append(f"{st}:{ws.cell(row,2).value[:3]}={ws.cell(row,4).value:.4g}" if isinstance(ws.cell(row,4).value,(int,float)) else f"{st}:{ws.cell(row,2).value[:3]}")
    def f(v):
        return f"{v:,.3f}" if isinstance(v, float) and abs(v) < 10 else (f"{v:,.0f}" if isinstance(v, (int, float)) else str(v))
    print(f"{tag:12} err={r.get('total_errors')} " + " ".join(f"{k}={f(v)}" for k, v in vals.items()), fails)
