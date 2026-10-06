# Build note — Rapha Development Profitability Model v2

## Deliverables
| File | What it is |
|---|---|
| `Rapha_Development_Profitability_Model_v2.xlsx` | Template with three samples. Opens on the Commercial track, Sample 1 (B+G+8 mixed-use), type 3 Mixed, Base scenario. The private track holds a sample G+1 house. |
| `Rapha_Model_Blank.xlsx` | The same workbook with every data column cleared (statutory rates kept). |
| `Rapha_Model_Tutorial_AM.pdf` | 34-page Amharic guide with a private path, a commercial path, Google Sheets use, two worked examples and a glossary. |

## What changed from v1
- **Track choice on Start.** Each choice greys out the other card and its "open" button. Every sheet of the track not chosen shows an amber header band saying it is not in use.
  - Private residence: separate sheets P1 Inputs, P2 Calc, P3 Estimate and P4 Checks, with their own cost library, monthly cost and funding, a vertical 240-month loan schedule and an affordability test.
  - Commercial project: type 1 Sell, 2 Hold or 3 Mixed. No personal fields (household income, savings and home loan moved to the private track).
- **13 Spaces rent roll.**
  - Up to 12 spaces, each with name, type (11 types, dropdown), net area, sold % (Mixed), price, rent and stabilised occupancy.
  - Hotel, owner-occupied and rent-free spaces are never sold; owner-held space is IAS 16.
- **New commercial inputs.**
  - Hotel operator lease: keys, ADR, occupancy, F&B uplift, turnover rent with a minimum guarantee.
  - Four named other-income streams and a lease-up ramp.
  - Specialist consultancy (capitalised) and pre-opening cost (expensed but funded).
  - Facility headroom for bank fees.
  - The VAT share on rent is computed from the rent roll.
  - A Confidence column (High/Medium/Low) and a Source column for every input, with a count of Low-confidence inputs.
- **Google Sheets.**
  - Static internal hyperlinks, which Google Sheets converts to sheet jumps; tested on a converted file.
  - Dropdowns read sheet ranges.
  - Conditional formats read only cells on their own sheet.
  - No Data Tables, and no names that look like cell references (`RATE` → `INT_RATE`, `DISC` → `DISC_RATE`, `OI1` → `OINC1`).
- **52 Sensitivity is fully live.**
  - What-if dials, with check C17 warning when a dial is not neutral.
  - First-order NPV tornado (sorted, charted on the Dashboard).
  - Eight closed-form break-evens with headroom: price ×2, rent, occupancy, cost overrun, rate, delay and exit cap.
- **Dashboard.** Seven auto-written key findings, plus a Low-confidence alert.

## Size
- 31 sheets (30 visible, plus a hidden `Lists` sheet).
- 598 named ranges, 24,041 formulas and 104 static links.
- Checks: commercial 25 (16 errors, 9 warnings); private 12 (8 errors, 4 warnings).

## Results (Base)
| Case | Formula errors | Checks | Key results |
|---|---|---|---|
| Sample 1, type 3 Mixed | 0 | 25/25 pass | TDC 550.8m, equity IRR 27.4%, min DSCR 1.77x, VIABLE |
| Sample 2, type 2 Hold (trade centre + hotel, land in kind) | 0 | 25/25 pass | TDC 1,449.0m, equity IRR 24.4%, min DSCR 1.39x, NCI 30%, VIABLE |
| Sample house (private track) | 0 | 12/12 pass | Total cost 12.91m, instalment 72,062/month, 24% of income |
| All 3 types × both samples, plus Upside, Downside and Lender scenarios | 0 | 0 FAIL in every case | Warnings only where the project genuinely breaches (e.g. downside DSCR) |
| Blank, all types and scenarios, both tracks | 0 | 0 FAIL | One warning: "enter at least one space" |

## Verified
- **Independent recomputation** (`qa/verify.py`, both samples):
  - Loan PMT and the IDC recursion.
  - Loan roll-forward for O3.
  - Tax for O5: building allowance, CIT and the UPT closed form.
  - Balance sheet for O5 rebuilt from its components, and BS cash = CF cash.
  - XIRR, IRR and NPV against Python.
- **Cached values:** all 24,042 cached formula values equal a fresh LibreOffice recalculation (`qa/inject_check.py`).
- **Google Sheets audit** (`qa/gsheets_audit.py`): no HYPERLINK formulas, Data Tables, cross-sheet or named conditional formats, name-sourced dropdowns, or cell-like names.
- **Hard-code scan:** the only literals are structural (month offsets, library row indices, unit constants) and the 50–95% net-to-gross plausibility band in check C22.

## Assumed / simplified
- **Allocation and impairment:** TDC is allocated to IAS 2, IAS 40 and IAS 16 by net area. IAS 36 is an indicator test only.
- **Hotel:** the operator bears the hotel's operating costs. Property opex (as % of rent) applies to space rent only.
- **Occupancy:** a rent-weighted stabilised occupancy is applied to every space, ramped up by RAMP_O1 and RAMP_O2.
- **Tornado:** first-order, pre-tax estimate. The exact effect of any driver comes from the what-if dials.
- **Private track:** the retention is set aside in the handover month. Capitalised interest during construction is added to the loan above the bank limit.
- **Sample figures:** all sample prices are illustrative 2026–27 Addis Ababa values — "Sample value — not market advice".

## Tax rates you must confirm (sheet 12 – all flagged "verify before use")
- **Income tax and allowances:** CIT 30%; building tax depreciation 5%; plant pool 20%; loss carry-forward 5 years.
- **VAT:** 15%.
- **Profit distribution:** undistributed profits tax 10%; dividend WHT 10%.
- **Legal reserve:** 5% of profit, up to 10% of capital.
- **Property tax:** 0.10%.
- **Land lease:** 10% down payment, 30 years, 7% interest.

## Rebuild
```
cd build
python finalize.py --out X.xlsx                    # Sample 1 commercial (default)
python finalize.py --blank --out Blank.xlsx
python finalize.py --data 3 --mode 2 --out s2.xlsx  # Sample 2 view
python finalize.py --track 1 --out priv.xlsx        # private-track view
python qa/matrix.py X.xlsx full ; python qa/verify.py <calc copy> ; python qa/gsheets_audit.py X.xlsx
cd ../tutorial && python shots.py <file> img2 <prefix> && RAPHA_SRC=<dir> python build_tutorial.py
```
