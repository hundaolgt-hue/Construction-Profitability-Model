# Amharic tutorial, buyer files and thank-you note

## Tutorial (PDF, 30–40 pages, Amharic with English terms)
- Generator: `assets/reference-generator/tutorial/build_tutorial.py`, which writes HTML and converts it to PDF with
  Playwright Chromium (`scripts/html_to_pdf.js` adds the running header and page numbers).
- Install the Noto Sans Ethiopic and Noto Serif Ethiopic fonts in `~/.fonts`. On cloud machines Chromium is at
  `/opt/pw-browsers/chromium-*/chrome-linux/chrome`.
- **Screenshots:** take them from the finished workbooks (`shots.py`, LibreOffice render). Crop by fractions and place
  numbered orange callouts with % coordinates, each explained in a legend under the image.
  - Re-check callout positions every time a layout changes.
  - Use `figure{display:inline-block}` so images don't split across pages.
- **Numbers in the text** come from the workbook's cached values (`facts.py`), never typed by hand.
- **Table of contents:** use a two-pass render. Put hidden markers (CH01…) in the chapter headings, read their page
  numbers with pdfplumber, then render again.
- **Structure for a two-track product:**
  1. Intro: what is new.
  2. The Start sheet and track choice, including the "type 1/2" fallback.
  3. Excel and Google Sheets setup: Open with › Google Sheets; leave iterative calculation off.
  4. Private chapters.
  5. Commercial chapters: inputs, spaces, engine, statements, returns, sensitivity, dashboard and checks.
  6. Worked samples.
  7. Own-project checklist and FAQ.
  8. Glossary and disclaimer.

## Buyer distribution
- Folder `output/dist/` holds the model, the blank template, the tutorial and the thank-you PDF.
- Run `scripts/neutralize.py output/dist <files>` so the author is "Rapha" and the files carry no internal names or
  tool fingerprints. Buyer files must have no internal notes.
- **Google Drive folder for buyers:** create it with the connector. The connector needs inline base64, so files over a
  few hundred KB can't be uploaded that way. Give the user the folder link, ask them to upload the dist files, and
  have them set "Anyone with the link – Viewer".
- **Thank-you PDF:** `assets/thank-you/ty.html` rendered with `p.js` (Playwright). It is one bilingual page that thanks
  the buyer and asks for a review.

## Delivery hygiene
- Commit on the designated branch. If `git push` returns 403, say so plainly and give the GitHub reconnect steps; don't
  hide it.
- Use `SendUserFile` with absolute paths to send the dist files on every release. In the same message, tell the user
  how to recognise the new version.
