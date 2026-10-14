# Google Sheets compatibility rules (learned the hard way)

Buyers open the .xlsx in Google Sheets at least as often as in Excel. LibreOffice recalculation (the usual QA
step) is *more forgiving* than Google Sheets, so a file can pass every local check and still be broken for the
buyer. Every rule below comes from a real failure.

## Rules

| # | Rule | Why |
|---|------|-----|
| 1 | **Dropdown list sources must not start with `=`.** Write `formula1='"A,B,C"'` (inline list) or `formula1="Lists!$A$2:$A$5"`. Never `"=Lists!$A$2:$A$5"`, never a defined name. | Google Sheets silently **drops** list validations written with a leading `=`; only the numeric validations survive. LibreOffice accepts it, so local QA does not catch it. This broke the Private/Commercial selector in the field. |
| 2 | **Prefer inline lists** (`'"opt1,opt2,opt3"'`) whenever the joined text ≤ 250 chars and no item contains a comma or a double quote. Fall back to a sheet range otherwise. | Inline lists are stored inside the rule, so nothing can break on import. Rewrite labels to avoid commas (e.g. "sell part + hold part"). |
| 3 | **Key selectors: single unmerged cell, `showErrorMessage=False`, and accept a typed number.** Labels start with their code ("1 – Private residence …", "2 – Commercial …"); the helper formula is `IFERROR(MATCH(cell,LIST,0),IFERROR(MAX(1,MIN(n,VALUE(LEFT(TRIM(cell&""),1)))),default))`. | Gives users a fallback that works even if the dropdown UI fails: they can just type 1 or 2. Merged cells and strict validation are both fragile. |
| 4 | **Navigation = static hyperlinks**, never `HYPERLINK("#'Sheet'!A1",…)` formulas. In openpyxl: `cell.hyperlink = Hyperlink(ref=cell.coordinate, location="'Sheet'!A1", display=text)`. | Google Sheets converts static internal links to `#gid=` jumps (verified). HYPERLINK formulas with `#Sheet!A1` do nothing in Sheets. |
| 5 | **Conditional-format formulas may only reference cells on their own sheet.** No defined names, no other sheets. Put a small "helpers" column on the sheet (e.g. `P9 = TRACK`) and point the rule at `$P$9`. | Sheets rejects cross-sheet and named references in custom CF formulas. |
| 6 | **No Excel Data Tables (`TABLE()`).** Build sensitivity with "what-if dial" input cells and closed-form tornado/break-even formulas. | Sheets doesn't support them, and LibreOffice computes them wrongly. |
| 7 | **No defined names that look like cell references or clash with functions**: `OI1`, `FL1`, `R1C1`-like, `RATE`, `DISC`. Use `OINC1`, `INT_RATE`, `DISC_RATE`. | Ambiguous in Sheets and Excel. Run the name audit. |
| 8 | **Avoid** `INDIRECT`, `CELL`, `INFO`, `GETPIVOTDATA`, macros, and number-format toggles driven by conditional formatting. | Unsupported or volatile in Sheets. A units toggle via CF number formats is ignored by Sheets. |
| 9 | XIRR, EDATE, SUMPRODUCT(--ISERROR(...)), TEXT with Unicode, `#,##0,` thousands formats, CHOOSE/INDEX/MATCH, PMT and PV all work (verified). | Safe toolbox. |

## How to verify for real (don't trust LibreOffice alone)

1. Run `scripts/gsheets_audit.py model.xlsx` (it checks rules 1, 4–8). It must print `NO GOOGLE SHEETS ISSUES`.
2. For a new technique, make a **tiny test .xlsx** (two sheets, the feature in question) and upload it with the Google Drive
   connector `create_file` (base64, xlsx mime; it converts to a Google Sheet). Then:
   - `download_file_content` with `text/csv` shows the computed values of the first sheet.
   - `download_file_content` with the xlsx mime shows what survived the conversion (validations, links, formulas).
     Pad the test file with a sheet containing `=SEQUENCE(15000,20)` so the export is big enough to be saved to a file
     automatically. Then decode it with Python instead of pasting base64 by hand, which corrupts easily.
   - `application/zip` export gives HTML, where internal links show as `href="#gid=…"`.
   - Trash the test file afterwards.
3. When a user says "it doesn't work in Google Sheets", **search their Drive for the converted copy and read it** before
   changing anything. In this project the user was testing an old download. The Start sheet text showed it immediately.
   Give each release a visible marker (e.g. the selector label text) so versions can be told apart.
