# Model architecture (real-estate / construction profitability)

The reference generator is in `assets/reference-generator/build/`. Read the file that matches the part you are building
rather than re-deriving it.

## Generator framework (`core.py`)
- `Builder` holds the workbook, the `names` dict (defined names) and `sheets`. A `Sheet` reserves rows lazily:
  - `line(key, label, f(i), fmt, unit, tot, style, n)` adds a time-series row; `f(i)` returns a formula string (no `=`),
    a number, or `"§text"` for literal text.
  - `scalar(key, label, f, name=…)` puts a single value in column D and can register a defined name.
  - `sec`, `note`, `gap` and `raw(fn, nrows)` handle layout and custom blocks.
- Formulas are rendered **after** every sheet has reserved its rows, so any sheet can reference rows on sheets declared
  later. Use lambdas when a scalar needs another sheet's row numbers.
- `header()` draws the navy band, the "◂ Start" and "Dashboard ▸" static links, period headers, freeze panes and print
  setup. Row 3 is a formula that shows a ⚠ "not in use" banner when the sheet belongs to the track not chosen on Start.
- Colours follow the FAST convention: blue text on amber for inputs, black for formulas, green for pure links, yellow
  for key outputs.

## Timeline
Construction is monthly (up to 48 months), rolled up to annual C1–C4; operations run annually O1–O10; the loan runs
monthly for 288 months. Annual periods sit in columns E…R (14 periods). Roll-ups use SUMIFS on a period-index row.

## Engine conventions that avoid circularity
- **IDC:** each month's interest = (opening balance + this month's draw) × rate / 12, capitalised during construction.
- **Arrangement fee:** charged on a facility base that excludes IDC.
- **Funding waterfall** (monthly): presale cash → equity commitment → debt (capped) → sponsor top-up. Add a
  "facility headroom %" so the debt also funds the bank fees; otherwise hold projects always show a small funding gap.
- **Undistributed-profits tax:** closed form r·k·PAT/(1+r·k).
- **Pre-opening costs:** included in the monthly funding need but expensed in P&L, not capitalised.

## IFRS treatment
- TDC is split across IAS 2 (sold), IAS 40 (let) and IAS 16 (owner-held) by net area from the space schedule.
- IFRS 15 sales: presales create a contract liability; revenue is recognised at handover or over time.
- IFRS 16 rent, with key money deferred.
- IAS 12: losses carried forward by vintage (FIFO, 5 years); deferred tax = CIT × (cumulative tax allowances −
  cumulative book depreciation).
- IAS 36 is an indicator only. IAS 7 cash flow uses the indirect method.
- NCI: land contributed in kind by a partner.

## Commercial features worth reusing
- **Space schedule / rent roll** (12 rows × type dropdown):
  - Non-saleable types: hotel, owner-occupied, rent-free.
  - Sold share = `IF(non-saleable,0,CHOOSE(mode,1,0,sold%))`.
- **Lease-up:** O1 and O2 occupancy as a % of stabilised occupancy.
- **Hotel operator lease:** keys × 365 × occupancy × ADR × (1 + F&B share); owner rent = MAX(turnover %, minimum guarantee).
- **Other income:** four named streams.
- **Inputs:** each input row has Confidence (High/Medium/Low) and Source columns; the count of Low inputs feeds a
  Dashboard alert.
- **Dashboard key findings:** sentences built with `TEXT()` from live names, so the commentary updates itself.
- **Live sensitivity:**
  - what-if dials (`FLX_*` multipliers) feed every driver, with a check that warns when a dial is not neutral;
  - a first-order NPV tornado from present-value helpers (sales, rent, hotel, capex, opex, interest, 6-month delay);
  - closed-form break-evens: price, rent, occupancy (DSCR), cost overrun (IRR at hurdle), rate, delay, exit cap.
- **Verdict:** VIABLE / MARGINAL / NOT VIABLE from equity IRR vs hurdle, profit on cost and the DSCR covenant.

## Two-track workbooks (private vs commercial)
When one product serves both a home builder and a developer, keep **separate sections**:
- **Private sheets** (P1 inputs, P2 calculations with a vertical 240-row loan schedule, P3 one-page estimate, P4 checks)
  hold the personal fields: savings, family money, home loan, household income, affordability.
- **Commercial sheets** must not contain personal fields.
- **Start sheet:** step 1 = track selector, then one card per track. The inactive card is greyed out by a same-sheet CF
  rule on a helper cell, with an "Open …" button per track.
- Both tracks always calculate; the choice only steers the user.

## Samples
- Keep 2–3 data sets in input columns (E = user, F/G = samples). Column H (active) = `CHOOSE(DATA_SET,E,F,G)`, text
  dropdowns become codes via MATCH, and Yes/No becomes 1/0.
- Tune samples until they are realistic **and** pass covenants. For example, a hold project needs capitalised grace in
  O1 plus a longer tenor, or DSCR fails in the lease-up year.
- Label every sample value with a cell comment: "Sample value — not market advice".
