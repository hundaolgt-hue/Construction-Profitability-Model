# Build note — Rapha Development Profitability Model v1

## Deliverables
| File | What it is |
|---|---|
| `Rapha_Development_Profitability_Model_v1.xlsx` | Template pre-loaded with **Sample A** (G+1 house, Mode 0) and **Sample B** (B+G+8 mixed-use, Mode 3), switchable on the Start sheet. Opens on Sample B / Mode 3 / Base. |
| `Rapha_Model_Blank.xlsx` | Same workbook: every input column holds neutral defaults (statutory rates kept), and the Start sheet is set to My project / Mode 0. |
| `Rapha_Model_Tutorial_AM.pdf` | 36-page Amharic user guide (A4) with real, annotated screenshots of the model. |

## Size
- 27 sheets (26 visible, plus a hidden `Lists` sheet for dropdowns).
- 469 named ranges.
- 20,527 formulas, plus 14 live Excel Data Tables (297 table cells).
- 354 "Sample value — not market advice" comments.
- 23 integrity checks: 17 errors and 6 warnings.
- Timeline: construction is monthly (up to 48 months), rolled up to C1–C4. Operations run annually for O1–O10. The loan is modelled monthly for 288 months, so tenors of up to 20 years are supported.

## Check results
| Case | Formula errors | Checks |
|---|---|---|
| Sample B, Mode 3, Base (as delivered) | 0 | 23/23 pass |
| Sample A, Mode 0, Base | 0 | 23/23 pass |
| All 4 modes × both samples; all 4 scenarios | 0 | 0 FAIL in every case (warnings only where the project genuinely breaches, e.g. downside DSCR) |
| Blank, all 4 modes × Base/Downside | 0 | 23/23 pass |
| LibreOffice `recalc.py` on both shipped files | 0 of 20,770 | — |

## Verified
- **PMT:** Sample B's first quarterly repayment equals `numpy_financial.pmt` less interest (738,416 ETB).
- **Loan year O3:** opening + draws + capitalised interest − principal − sweep = closing.
- **IDC:** the monthly recursion matches (opening + draw) × r / 12.
- **Tax year O5:** taxable income, the 5% building allowance, CIT after loss usage, and the UPT closed form all reproduce by hand.
- **Balance sheet O5:** rebuilt from its components, assets = liabilities + equity, and BS cash = CF cash.
- **Returns:**
  - Equity IRR and project IRR (XIRR) match an independent Python XIRR to within 0.0001.
  - The per-period IRR matches `numpy_financial.irr`.
  - NPV matches independent discounting.
- **Data Tables:** every table cell was computed by a full model recalculation, one per input value. Sampled cells were re-checked against fresh recalculations and matched exactly. The tables remain live `TABLE()` Data Tables in Excel.
- **Home Estimate total equals the commercial TDC** for the same inputs (check C16), in all modes.
- **Hard-code scan:** no literal numbers remain in formulas except structural lookups and unit constants (12 months, 365 days).

## Assumed / simplified (documented in the workbook)
- **Reference model:** the TDOMA/Liiban reference workbook was not attached. The architecture follows the build prompt's description of it, and no reference numbers or names are used.
- **Cost allocation:** TDC is split across IAS 2, IAS 40 and IAS 16 by net area. IAS 36 is a test with an indicator only; no impairment is booked.
- **Project cash flow:** uses the actual (levered) tax paid.
- **Undistributed-profits tax:** charged in P&L using the closed form r·k·PAT/(1+r·k), with k = (1 − reserve %)(1 − payout).
- **Distributions:** dividends, return of capital and special dividends are subject to a DSCR lock-up test.
- **Land:** land is not depreciated. Lease instalments are annual, with interest on the unpaid balance.
- **VAT in Mode 0:** cost rates are treated as VAT-inclusive, because households cannot reclaim VAT.
- **Break-even:** the closed-form values are first-order estimates. The grid-refined values come from 21-row live Data Tables. The delay test is limited by the 48-month construction maximum.
- **Sample prices:** all sample costs and prices are illustrative 2026–27 Addis Ababa figures, not market advice.
- **LibreOffice:** its Data Table engine is unreliable, so use Excel for the sensitivity sheet (press F9 if calculation is set to "Automatic except tables").

## Tax rates you must confirm (sheet 12 — all flagged "verify before use")
| Item | Rate used | Source |
|---|---|---|
| CIT | 30% | Proc. 979/2016 Art. 19 |
| Tax depreciation – buildings | 5% straight line | Proc. 979/2016 Art. 25 / Reg. 410/2017 |
| Tax depreciation – plant pool | 20% diminishing | Proc. 979/2016 Art. 25 |
| Loss carry-forward | 5 years, 100% cap | Proc. 979/2016 Art. 26 |
| VAT | 15% | VAT Proc. 1341/2024 |
| VAT on building sales | 0% share (configurable) | VAT Proc. 1341/2024 |
| VAT on rent | 100% commercial (configurable) | VAT Proc. 1341/2024 |
| Undistributed profits tax | 10% | Proc. 979/2016 Art. 55 |
| Dividend WHT | 10% | Proc. 979/2016 Art. 55 |
| Legal reserve | 5% up to 10% of capital | Commercial Code Proc. 1243/2021 Art. 434 |
| Property tax | 0.10% (rate and base) | City proclamation |
| Land lease | 10% down payment, 30 years, 7% interest | Proc. 721/2011, lease contract |

## Rebuild
```
cd build
python finalize.py --out ../output/Rapha_Development_Profitability_Model_v1.xlsx
python finalize.py --blank --out ../output/Rapha_Model_Blank.xlsx
```
Data Table results are cached in `build/.dtcache`. A cold run takes about 25 minutes.

Then build the tutorial (`tutorial/build_tutorial.py`, after `tutorial/shots.py`), and run the QA scripts in `build/qa/` (`matrix.py`, `verify.py`, `dt_spot.py`, `inject_check.py`, `hardcode_scan.py`).
