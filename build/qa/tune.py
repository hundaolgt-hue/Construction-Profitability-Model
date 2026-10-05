import sys
from harness import run_many, values, nm
src = "../../output/test.xlsx"
specs = []
for sw in (0.25, 0.4):
    for ten in (10,):
        for pr in (125000, 130000):
            specs.append((f"sw{sw}_t{ten}_p{pr}", dict(mode=3, data=3, scen=1, inputs={"SWEEP": sw, "TENOR_Y": ten, "PRICE_RES": pr})))
res = run_many(src, specs, workers=4, keep_dir="/tmp/claude-0/qa_tune")
for tag, (r, p) in sorted(res.items()):
    wb = values(p)
    print(tag, {k: round(nm(wb, k), 3) if isinstance(nm(wb, k), float) else nm(wb, k) for k in ["EIRR", "PIRR", "POC", "MIN_DSCR", "AVG_DSCR", "N_BREACH", "LN_INSTAL", "PEAK_DEBT", "VERDICT", "N_DS_YEARS"]})
