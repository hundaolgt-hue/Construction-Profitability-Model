# Rapha Development Profitability Model v2 (የራፋ የግንባታ ፕሮጀክት ትርፋማነት ሞዴል)

One workbook with two separate tracks, chosen on the Start sheet:

- **A · Private residence.** Sheets P1–P4 cover a house cost estimate, the cash needed each month, a savings → family → home-loan plan, a 240-month loan schedule, an affordability test and an optional rent-out of part of the house.
- **B · Commercial project** (sell, hold or mixed). It covers a space schedule and rent roll, a hotel operator lease, other income streams, 10-year IFRS statements, the loan and lender ratios, returns, live sensitivity and break-even analysis, an auto-written key-findings section and 25 integrity checks.

The workbook works in Excel, LibreOffice and Google Sheets. Navigation uses static links, there are no Data Tables or macros, and conditional formats read only cells on their own sheet.

- `output/Rapha_Development_Profitability_Model_v2.xlsx` — the template, with a sample house and two commercial samples
- `output/Rapha_Model_Blank.xlsx` — a blank copy
- `output/Rapha_Model_Tutorial_AM.pdf` — the Amharic user guide (34 pages)
- `output/BUILD_NOTE.md` — counts, verification results, assumptions and the tax rates to confirm
- `output/dist/` — buyer copies with neutral metadata
- `build/` — the generator (openpyxl) and the QA harness; `tutorial/` — the guide generator (HTML → Chromium PDF)

This is a planning tool only. It is not investment, tax or legal advice.
