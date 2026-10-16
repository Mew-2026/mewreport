# mewreport

Checks a handwritten truck loading sheet against the Luminous scrap tax invoice for the same load
and produces an Excel reconciliation report.

1. **Loading sheet**, either way:
   - **Photo**: read with the [renamed.to](https://www.renamed.to) API. renamed.to only returns a
     filename, so the photo is cut into small overlapping strips and each strip is asked to write its
     rows into the filename as `ROW.MODEL-QTY_...`; the strips are stitched back together
     (`loading_sheet_ocr.py`). One sheet costs roughly 11-15 credits.
   - **Excel / CSV** typed as `Sr No | Model | Qty`: read directly, no credits (`loading_sheet_excel.py`).
     The page offers the format as a download (`public/loading_sheet_template.xlsx`, pre-filled with
     the 13/10/26 sheet as an example; regenerate with `python loading_sheet_excel.py --template ...`).
   Either way, the rows read are shown on the page in an editable grid (next to the photo, when it is
   one) before anything is matched: fix misread models or quantities, add or delete rows, then
   *Check against invoices*. *Save as Excel* keeps the corrected sheet for reuse without credits.
2. **Invoices** (system-generated PDFs) are read from their text layer by position: the table's
   column lines and header labels locate each cell, so the result does not depend on how the
   PDF stores its text. First copy only; pcs read are checked against the printed total
   (`invoice_parser.py`).
   On the Step 1 result, incorrect rows can be **matched by hand**: tick any rows (they need not be
   next to each other), add an optional note, and group them. A group whose differences add up to 0
   becomes MATCHED; otherwise PARTLY MATCHED, showing the difference left. Groups can be undone; the
   Excel report (and Step 2's combined report) lists them with their net difference and note.
3. **Matching** pairs each sheet model with an invoice model by fuzzy match, ignoring version
   tags (V21, R2, C…) and keeping Pro separate (`reconcile.py`).

4. **Rate check** (step 2): upload the mastersheet (*Item code*, *Material description*, *weight (kg)*;
   `.xlsx`/`.xlsm`/`.csv`, columns found by their contents even if shifted under the headers) and type
   today's Rate/Kg. Expected rate = Rate/Kg x mastersheet kg; CORRECT when the invoice Rate(Rs) is within
   the tolerance % (default 0.5). Lines are found by item code (O/0 and I/1 treated alike), else by
   description with the same battery model and numbers (flagged blue). When Step 1 has been run on the
   page, the output is one combined report (the Step 1 Excel with a *Rate check* tab, and a PDF of both);
   otherwise the invoice PDFs chosen in Step 1 are checked on their own - Step 1 need not be run
   (`rate_check.py`, `mastersheet.py`, `step1_report.py`).

## Run on your computer (Windows)

1. Install Python 3.10+ from python.org (tick "Add python.exe to PATH").
2. Double-click `run_app.bat`. The first time it creates `.env` and opens it in Notepad: paste
   your renamed.to API key after `RENAMED_API_KEY=`, save, and double-click `run_app.bat` again.
3. The app opens at http://127.0.0.1:5000. Choose the loading sheet photo and the invoice PDFs.

Try it with `samples/loading_sheet_13-10-26.jpg` and the PDFs in `demo/`.

## Setup (any OS)

```bash
python -m pip install -r requirements.txt
```

Create a `.env` file next to the scripts (it is git-ignored):

```
RENAMED_API_KEY=rt_...
```

## Run

```bash
python reconcile.py --sheet loading_sheet.jpg --invoice 5260118338.PDF 5260118339.PDF --out reports/report.xlsx
```

Rate check on a Step 1 report, giving the combined `.xlsx` and `.pdf`:

```bash
python rate_check.py --step1 reports/report.xlsx --master mastersheet.xlsx --rate 119.54 --tolerance 0.5 --out reports/combined
```

(`--invoice a.pdf b.pdf` in place of `--step1` gives the rate check alone.)

The OCR result is saved next to the photo as `<name>.ocr.json`. To rebuild the report without
spending credits again, pass `--sheet-json <name>.ocr.json` instead of `--sheet`.

## Web app (Vercel)

`public/index.html` is an upload page; `api/index.py` (Flask) runs the same check and returns the
summary plus the Excel file. Set `RENAMED_API_KEY` (the renamed.to token) as an environment
variable in Vercel. The page has no password: anyone with the link can run checks and spend
renamed.to credits.

Run it locally with `python -m flask --app api/index.py run`.

A load can have several invoices: pass them all after `--invoice`. Their lines are added up
before matching; the same invoice given twice is counted once.

## Demo invoices

`demo/` holds DEMO invoices in the exact layout of the real invoice (grid, columns, fonts and
positions measured from it, drawn by `demo/invoice_layout.py`), with made-up supplier, recipient,
GSTINs, IRN and invoice numbers and a DEMO watermark. They are built from the 13/10/26 loading sheet:

- `match_1_*.pdf` + `match_2_*.pdf`: together match the sheet exactly (440 pcs) -> all CORRECT
- `mismatch_1_*.pdf` + `mismatch_2_*.pdf`: three planted errors (SC16060 billed 38 not 40,
  RC18000P not billed, crossed-out PC20042TT billed) -> 3 INCORRECT

Regenerate with `python demo/make_demo_invoices.py` (needs `reportlab`, see
`requirements-dev.txt`). Invoices are read from the PDF text layer (`pdfplumber`), never through renamed.to, so they cost no credits; only the handwritten sheet does.

## Report

Every row is marked **CORRECT** (green) or **INCORRECT** (red), incorrect rows first, with the reason:

| Status | Meaning |
|---|---|
| OK | same quantity on sheet and invoice |
| QTY MISMATCH | model on both, quantities differ |
| NOT IN INVOICE | model on the sheet but not billed |
| NOT IN SHEET | billed but not on the sheet |

Blue model cells mark fuzzy pairings below 90% — check those by eye. The *Loading sheet (OCR)*
tab shows the raw renamed.to replies so a misread can be traced.

## Limits

- Strip positions are tuned to the current sheet layout (one long left column, short right
  column). A different layout needs the `columns` values in `loading_sheet_ocr.py` adjusted.
- renamed.to's replies vary between calls; the script retries incomplete replies, but always
  compare the OCR total with the total written on the sheet.
- Crossed-out rows are usually skipped, not always.
