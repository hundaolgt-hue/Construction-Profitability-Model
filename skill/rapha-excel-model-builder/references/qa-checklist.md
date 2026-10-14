# QA checklist (run before every delivery)

Do not report a model as finished until every line below holds and you have the evidence.

1. **0 formula errors.** Recalculate a copy with LibreOffice (`scripts/recalc_check.py model.xlsx NAME1,NAME2`). If
   only libreoffice-core is installed, install `libreoffice-calc` and `libreoffice-writer` first, or the conversion hangs.
2. **In-model checks sheet.** Build 20–25 checks into the model itself. Errors must PASS; warnings are advisory.
   - **Errors:** BS balances every year; CF closing cash = BS cash; sources = uses; TDC = accrual; allocation sums to
     100%; loan repaid in tenor; Σ principal = draws + capitalised interest; inventory and contract balances unwind;
     land not depreciated; tax losses ≥ 0; switches valid; an ISERROR scan of every sheet.
   - **Warnings:** DSCR covenant, minimum cash, funding gap, FAR, net-to-gross plausibility, dials not neutral.
3. **Matrix run.** Run every project type × sample × scenario × track, plus the blank file (`qa/matrix.py`):
   0 recalc errors and 0 FAIL everywhere. Warnings are allowed only where the case genuinely breaches (e.g. downside DSCR).
4. **Independent verification** (`qa/verify.py`): recompute in Python and compare —
   - one loan repayment (`numpy_financial.pmt`) and the IDC recursion;
   - one tax year;
   - the balance sheet rebuilt from its components;
   - XIRR, IRR and NPV.
5. **Cached values.** Recalculate with LibreOffice, then inject every value into the openpyxl-built file
   (`postprocess.inject`) so previews and Google Drive show numbers. Then `inject_check.py` must report 0 mismatches.
6. **Hard-code scan** (`hardcode_scan.py`): only structural literals remain (12, 365, offsets and indices).
7. **Google Sheets audit** (`gsheets_audit.py`): `NO GOOGLE SHEETS ISSUES`. For any new UI technique, run the real
   Drive round-trip test described in google-sheets-rules.md.
8. **Visual check.** Render the key sheets to PNG (`tutorial/shots.py`: LibreOffice → PDF → pdftoppm → crop) and
   look at Start, Dashboard, the report pages and the inputs. Fix overflowing labels and cramped rows.
9. **Blank template.** Generate it from the same code with the data columns cleared (statutory rates kept). It must pass too.

Report results honestly: say what was tested and how. In particular, say when something was **not** tested in the
buyer's real environment (e.g. the full file is too big to upload through the Drive connector).
