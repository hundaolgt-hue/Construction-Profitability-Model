---
name: rapha-excel-model-builder
description: Build, rebuild or fix premium Excel / Google Sheets financial-model templates for Rapha (Hundaol) — construction cost estimators, development profitability models (sell / rent / mixed), home-loan and affordability calculators, real-estate feasibility, IFRS statements, loan schedules, sensitivity and break-even — generated with Python/openpyxl, fully QA'd, Google-Sheets-safe, with an Amharic PDF tutorial and buyer-ready files. Use this whenever the user asks to build, extend, "make premium", tune samples, restructure (e.g. private vs commercial track), fix dropdowns, links or switches, or make a model work in Google Sheets for products like ProfitCheck, SiteCalc Pro or ProjectControl, or uploads a build prompt or a reference workbook (e.g. a TDOMA-style model) to copy. Also use it when they report "the selector / dropdown / link is not working in Google Sheets", or want the Amharic tutorial, blank template, thank-you PDF or Drive buyer folder for such a model.
---

# Rapha Excel model builder

This skill captures how the Rapha Development Profitability Model (v1 → v2) was built:
- a ~24,000-formula workbook generated from Python;
- a private-residence track and a commercial track;
- IFRS statements, a monthly loan schedule, live sensitivity and 25 + 12 integrity checks;
- a 35-page Amharic tutorial and neutral buyer copies.

Reuse the method and the code. Don't start from a blank page.

## Resources in this skill
- `assets/reference-generator/build/` — the complete working generator: `core.py` framework, `inputs_def.py` /
  `private_def.py` input registers, `engine.py`, `fin.py`, `front.py`, `private.py`, `main.py`, `finalize.py` and `qa/`.
  Copy it into the project and adapt it; it is the fastest correct starting point.
- `assets/reference-generator/tutorial/` — the Amharic tutorial generator (`build_tutorial.py`, `shots.py`, `facts.py`).
- `assets/thank-you/` — the one-page bilingual thank-you / review-request PDF template.
- `scripts/` — standalone tools:
  - `gsheets_audit.py` — Google Sheets compatibility audit;
  - `recalc_check.py` — LibreOffice recalculation plus named outputs;
  - `inject_check.py` and `hardcode_scan.py` — cached-value and hard-code checks;
  - `postprocess.py` — inject cached values into the built file;
  - `neutralize.py` — buyer-copy metadata;
  - `html_to_pdf.js` — Chromium PDF with header and footer;
  - `model_core.py` — copy of the framework.
- `references/` — read the one you need:
  - `model-architecture.md` — framework, timeline, no-circularity tricks, IFRS, commercial features, the two-track pattern.
  - `google-sheets-rules.md` — **read before writing any dropdown, link, conditional format or name.** Every rule
    there came from a real failure.
  - `qa-checklist.md` — what must pass before delivery.
  - `tutorial-and-distribution.md` — the Amharic tutorial, dist files, Drive folder and thank-you PDF.

## Workflow

1. **Understand the ask.** Work out which product it is, which tracks or modes it needs, the samples, what to borrow
   from any reference workbook, and the output files. If the user changes direction mid-way (as happened with
   "keep the private estimator, but as its own section"), restate the new design in one line and follow it.
2. **Set up the generator.** Copy `assets/reference-generator/build` into the project and edit the input registers
   first: one tuple per input (key, English and Amharic labels, unit, format, user / sample values, note, validation).
   Every choice is a dropdown.
3. **Build in layers:** inputs → engine (timeline, area/spaces, capex, monthly construction and funding, revenue, opex,
   assets, loan, tax, VAT, working capital) → statements → returns, sensitivity and checks → front sheets (Start,
   Dashboard, report pages) → `main.py` ordering. Keep the formulas free of circular references (see model-architecture.md).
4. **Make it Google-Sheets-safe as you go:**
   - static hyperlinks;
   - inline dropdown lists with no leading `=`;
   - conditional formats that read only same-sheet helper cells;
   - no Data Tables;
   - selectors that also accept a typed number.

   Run `gsheets_audit.py`.
5. **Finalize.** Run `finalize.py`: build → LibreOffice recalc → inject cached values. Do this for each delivered view
   (main, blank, and other sample/track views for screenshots).
6. **QA and tune.** Work through `references/qa-checklist.md` completely, including the matrix and independent
   verification. Tune the samples until they are realistic **and** pass their covenants.
7. **Tutorial.** Render screenshots from the finished files, pull the numbers from cached values, then build the Amharic
   PDF (two-pass table of contents). Look at the rendered pages before delivering.
8. **Deliver.**
   - Make neutral-metadata copies in `output/dist/`, update BUILD_NOTE.md and the README, commit and try the push.
   - Send the files with `SendUserFile`, and tell the user how to recognise this version (e.g. the selector label text).

## When the user says "it doesn't work" (especially in Google Sheets)
Don't guess, and don't re-ship blindly:
1. Search their Google Drive for the converted copy and read it (`download_file_content` as CSV for values, or as xlsx to
   see which validations and links survived). The cause may simply be an old download.
2. Reproduce the suspected cause with a tiny test file uploaded to Drive and read it back (method in
   google-sheets-rules.md). Trash the test file afterwards.
3. Fix it in the generator, add the defect to `gsheets_audit.py` so it can't come back, re-run the full QA, then
   re-deliver with a clear "how to use and how to recognise the new version" note.
4. Always give the user an immediate fallback that needs no fix, such as typing 1 or 2 into the selector.

## Style of the deliverable
- **FAST colours:** blue-on-amber inputs, black formulas, green links, yellow key outputs.
- **Labels:** Amharic next to English.
- **Headers:** navy band with "◂ Start" and "Dashboard ▸"/"Estimate ▸".
- **Samples:** marked "Sample value — not market advice". A tax-law register flags every rate "verify before use".
- **Disclaimer:** planning tool only — not investment, tax or legal advice.
- **Buyer files:** no internal notes; author "Rapha".
