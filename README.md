# mewreport

Checks a handwritten truck loading sheet against the Luminous scrap tax invoice for the same load
and produces an Excel reconciliation report.

1. **Loading sheet** (photo) is read with the [renamed.to](https://www.renamed.to) API. renamed.to
   only returns a filename, so the photo is cut into small overlapping strips and each strip is
   asked to write its rows into the filename as `ROW.MODEL-QTY_...`. The strips are then stitched
   back together (`loading_sheet_ocr.py`). One sheet costs roughly 11–15 credits.
2. **Invoice** (PDF) is read from its text layer, first copy only (`invoice_parser.py`).
3. **Matching** pairs each sheet model with an invoice model by fuzzy match, ignoring version
   tags (V21, R2, C…) and keeping Pro separate (`reconcile.py`).

## Setup

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

`demo/` holds DEMO invoices (made-up supplier, recipient and invoice numbers, watermarked) in the
same layout as the real invoice, built from the 13/10/26 loading sheet:

- `match_1_*.pdf` + `match_2_*.pdf`: together match the sheet exactly (440 pcs) -> all CORRECT
- `mismatch_1_*.pdf` + `mismatch_2_*.pdf`: three planted errors (SC16060 billed 38 not 40,
  RC18000P not billed, crossed-out PC20042TT billed) -> 3 INCORRECT

Regenerate with `python demo/make_demo_invoices.py`. Invoices are read from the PDF text layer
(`pypdf`), never through renamed.to, so they cost no credits; only the handwritten sheet does.

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
