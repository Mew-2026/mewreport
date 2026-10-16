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
   becomes MATCHED; otherwise PARTLY MATCHED, showing the difference left. Mistakes are cheap to fix:
   a *Manual matches* panel lists every group with *Ungroup* (and *Ungroup all*), a single row can be
   taken out of its group, and every change offers *Undo*. Nothing is re-read, so no credits are spent.
   The Excel report (and Step 2's combined report) lists the groups with their net difference and note.
3. **Matching** is strict: a sheet model matches an invoice model only when it is the same battery after
   correcting known reading slips (ILTO -> ILTJ, ILT1 -> ILTT, UCTI -> UCTT, 2042ST -> 2042STJ, a stray
   L after LPTT12150H). Version tags (V21, R2, C...) are ignored; a trailing **P means Pro** and Pro is
   never paired with non-Pro. Rows with the same model on the sheet are added up. Anything else stays
   NOT IN INVOICE / NOT IN SHEET with a "possible match" hint, to be matched by hand (`reconcile.py`).

4. **Weight & rate check**: the page has **one input card** for the whole load - loading sheet (photo or
   Excel), invoice PDFs, mastersheet (*Item code*, *Material description*, *weight (kg)* per piece) and
   today's Rate/Kg. *Read loading sheet* shows the sheet for checking; *Run check* fills two result tabs,
   **Loading sheet vs invoices** and **Weight & rate check** (`weight_recon.py`):
   - **Loaded kg** = sheet nos x master kg/pc; **billed kg** = invoice nos x invoice kg/pc, where
     invoice kg/pc = rate / Rate/Kg rounded to 2 decimals. **Variance** = billed - loaded, split into
   - **Check 1, unit price** (first): an invoice line matches when its kg/pc (2 decimals) equals its item
     code's master kg/pc - no tolerance; differences shown in kg,
   - **Check 2, quantity**: (invoice nos - sheet nos) x master kg/pc per model,
   - plus billed kg of models with no weight at all.
   A model **not in the master** is flagged NOT IN MASTER and weighed at invoice rate / Rate/Kg (not rounded; the
   line with the most pcs when it has several rates), so its loaded pcs still
   count in the loaded kg (otherwise the variance swings by their whole weight). When it was not billed, the
   invoice kg/pc of its manual-match partner, or else the closest billed model (Pro only with Pro), is used.
   Its unit price cannot be checked: its quantity difference falls in Check 2; Check 1 keeps only rounding and
   lines billed at a second rate.
   kg and rupee figures are whole numbers (kg/pc and Rs/kg keep 2 decimals); the quantity part absorbs
   any 1 kg rounding so the parts always add up. One model can have several item codes with different
   weights: each invoice line uses its own. Manual matches re-run the weight check automatically.
   Every results table lists similar models together (by model number, then letters), so look-alikes
   such as PC18054TJP / RC18054TJP or BC18048ST / RC18048ST sit side by side.
   The Excel adds *Weight reconciliation* and *Master sheet (loaded)* tabs (column D = loaded qty;
   sheet models not in the master listed below); the PDF opens with the weight answer
   (`rate_check.py`, `mastersheet.py`, `step1_report.py`).

**Company report (PDF and Excel)**: what is sent to the company - the load details (invoices, recipient,
Rate/Kg), loaded vs invoiced in pieces, kg and Rs with the difference, and the model-wise annexure (loaded
and invoiced pcs, kg/pc, loaded kg, invoiced kg, difference, remark). The combined PDF opens with it.

Models are read from descriptions by family (PC, RC, EC, SC, LM, TC, TZ, BC, NEO, ILTT/ILTJ, LPTT, ASTT, LPT,
UCTT, AP, AB, 4-digit Amaze TT/STJ ...), so solar models such as LPTT12150L, LPTT12135H or LPT1240H find
their master weight and match the sheet. A model not in the master and not billed borrows the invoice rate
of a billed model with the same number (Pro with Pro first: TZ15060STJ -> TC15060, RC26000 -> RC26000P).

Every table on the page - the loading sheet, Loading sheet vs invoices, Check 1 (unit price, line by
line) and Check 2 (quantity in kg) - has **Excel** and **PDF** buttons. They export exactly what the
table shows (filter and search applied, noted in the file), with the load details, whole-number kg and
rupees (kg/pc 2 decimals) and coloured status cells; columns that are blank in every row are left out
(`export_table.py`, `/api/export`). The loading sheet's Excel uses the format the app reads back.

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
python rate_check.py --step1 reports/report.xlsx --master mastersheet.xlsx --rate 119.54 --out reports/combined
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

Blue model cells mark sheet text that was corrected to reach the model (e.g. ILTO18148 read as ILTJ18148). The *Loading sheet (OCR)*
tab shows the raw renamed.to replies so a misread can be traced.

## Limits

- Strip positions are tuned to the current sheet layout (one long left column, short right
  column). A different layout needs the `columns` values in `loading_sheet_ocr.py` adjusted.
- renamed.to's replies vary between calls; the script retries incomplete replies, but always
  compare the OCR total with the total written on the sheet.
- Crossed-out rows are usually skipped, not always.
