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
python reconcile.py --sheet loading_sheet.jpg --invoice 5260118338.PDF --out reports/report.xlsx
```

The OCR result is saved next to the photo as `<name>.ocr.json`. To rebuild the report without
spending credits again, pass `--sheet-json <name>.ocr.json` instead of `--sheet`.

## Report

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
