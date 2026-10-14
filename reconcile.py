"""Match a handwritten loading sheet against its tax invoice and write an Excel report.

Usage:
  python reconcile.py --sheet loading_sheet.jpg --invoice invoice.pdf [--out report.xlsx]
  python reconcile.py --sheet-json sheet.json --invoice invoice.pdf   # reuse an earlier OCR run (no credits)
"""
import argparse
import json
import os
import re
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from rapidfuzz import fuzz

from invoice_parser import read_invoice

MIN_SCORE = 75     # below this a sheet model is treated as "not on invoice"
CHECK_SCORE = 90   # below this the pairing is flagged for a human to check


def normalize(model):
    """Smooth over common handwriting/OCR slips before fuzzy matching."""
    m = model.upper()
    m = re.sub(r"^ILT[O0]", "ILTJ", m)   # ILTJ read as ILTO / ILT0
    m = re.sub(r"^ILT[1I]", "ILTT", m)   # ILTT read as ILT1
    m = re.sub(r"^UCT[I1]", "UCTT", m)
    m = re.sub(r"(\d{4})ST$", r"\1STJ", m)  # 2042ST -> 2042STJ (J often lost)
    m = re.sub(r"(\d{4})STP$", r"\1STJP", m)
    m = re.sub(r"(12150H)L$", r"\1", m)     # LPTT12150HL: trailing L is a stray mark
    return m


def match(sheet_rows, inv_models):
    """Pair each sheet row with an invoice model, one-to-one, best scores first."""
    pairs = sorted(
        ((fuzz.ratio(normalize(r["model"]), k), i, k) for i, r in enumerate(sheet_rows) for k in inv_models),
        reverse=True,
    )
    by_row, used = {}, set()
    for score, i, k in pairs:
        if score < MIN_SCORE:
            break
        if i not in by_row and k not in used:
            by_row[i] = (k, score)
            used.add(k)
    return by_row


def reconcile(sheet, inv):
    agg = defaultdict(lambda: {"qty": 0, "lines": [], "codes": set(), "descs": set()})
    for l in inv["lines"]:
        a = agg[l["model"]]
        a["qty"] += l["qty"]
        a["lines"].append(l["sr"])
        a["codes"].add(l["code"])
        a["descs"].add(l["desc"])
    by_row = match(sheet["rows"], list(agg))

    out = []
    for i, r in enumerate(sheet["rows"]):
        k, score = by_row.get(i, (None, 0))
        a = agg.get(k)
        iq = a["qty"] if a else 0
        if not a:
            status = "NOT IN INVOICE"
        elif iq == r["qty"]:
            status = "OK"
        else:
            status = "QTY MISMATCH"
        out.append({
            "line": r["line"], "sheet_model": r["model"], "invoice_model": k or "",
            "score": round(score), "sheet_qty": r["qty"], "invoice_qty": iq, "diff": iq - r["qty"],
            "status": status, "check": bool(a) and score < CHECK_SCORE,
            "invoice_lines": ", ".join(map(str, a["lines"])) if a else "",
            "item_codes": ", ".join(sorted(a["codes"])) if a else "",
            "description": " / ".join(sorted(a["descs"])) if a else "",
        })
    matched = {o["invoice_model"] for o in out}
    for k, a in agg.items():
        if k not in matched:
            out.append({
                "line": "", "sheet_model": "", "invoice_model": k, "score": "", "sheet_qty": 0,
                "invoice_qty": a["qty"], "diff": a["qty"], "status": "NOT IN SHEET", "check": False,
                "invoice_lines": ", ".join(map(str, a["lines"])), "item_codes": ", ".join(sorted(a["codes"])),
                "description": " / ".join(sorted(a["descs"])),
            })
    return out


FILL = {
    "OK": PatternFill("solid", fgColor="C6EFCE"),
    "QTY MISMATCH": PatternFill("solid", fgColor="FFC7CE"),
    "NOT IN INVOICE": PatternFill("solid", fgColor="FFEB9C"),
    "NOT IN SHEET": PatternFill("solid", fgColor="FFEB9C"),
}
CHECK_FILL = PatternFill("solid", fgColor="BDD7EE")
BOLD = Font(bold=True)


def write_report(path, rows, sheet, inv):
    wb = Workbook()
    ws = wb.active
    ws.title = "Reconciliation"
    sheet_total = sum(r["qty"] for r in sheet["rows"])
    inv_total = sum(l["qty"] for l in inv["lines"])
    ws.append(["Invoice", inv["invoice_no"], "Date", inv["date"], "Recipient", inv["recipient"]])
    ws.append(["Loading sheet qty", sheet_total, "Invoice qty", inv_total, "Difference", inv_total - sheet_total])
    counts = defaultdict(int)
    for r in rows:
        counts[r["status"]] += 1
    ws.append(["Rows OK", counts["OK"], "Qty mismatch", counts["QTY MISMATCH"],
               "Not in invoice / sheet", counts["NOT IN INVOICE"] + counts["NOT IN SHEET"]])
    for row in ws.iter_rows(min_row=1, max_row=3):
        for c in row[::2]:
            c.font = BOLD
    ws.append([])
    cols = [("line", "Sheet line"), ("sheet_model", "Sheet model (as read)"), ("invoice_model", "Invoice model"),
            ("score", "Match %"), ("sheet_qty", "Sheet qty"), ("invoice_qty", "Invoice qty"),
            ("diff", "Diff (inv - sheet)"), ("status", "Status"), ("invoice_lines", "Invoice Sr. Nos"),
            ("item_codes", "Item codes"), ("description", "Invoice description")]
    ws.append([h for _, h in cols])
    hdr = ws.max_row
    for c in ws[hdr]:
        c.font = BOLD
    for r in rows:
        ws.append([r[k] for k, _ in cols])
        cells = ws[ws.max_row]
        cells[7].fill = FILL[r["status"]]
        if r["check"]:
            for c in cells[1:4]:
                c.fill = CHECK_FILL
    ws.append(["", "TOTAL", "", "", f"=SUM(E{hdr + 1}:E{ws.max_row})", f"=SUM(F{hdr + 1}:F{ws.max_row})",
               f"=SUM(G{hdr + 1}:G{ws.max_row})"])
    for c in ws[ws.max_row]:
        c.font = BOLD
    ws.append([])
    ws.append(["Blue model cells = fuzzy match below "
               f"{CHECK_SCORE}%: check the sheet model was read and paired correctly."])
    ws.freeze_panes = ws.cell(row=hdr + 1, column=1)
    for col, w in zip("ABCDEFGHIJK", [10, 22, 16, 9, 10, 11, 15, 16, 16, 34, 60]):
        ws.column_dimensions[col].width = w

    il = wb.create_sheet("Invoice lines")
    il.append(["Sr", "Item code", "Description", "Model", "Qty", "Rate", "Taxable value"])
    for l in inv["lines"]:
        il.append([l["sr"], l["code"], l["desc"], l["model"], l["qty"], l["rate"], l["value"]])
    for c in il[1]:
        c.font = BOLD
    for col, w in zip("ABCDEFG", [6, 18, 48, 14, 7, 10, 14]):
        il.column_dimensions[col].width = w

    ls = wb.create_sheet("Loading sheet (OCR)")
    ls.append(["Line", "Row no. on sheet", "Model", "Qty"])
    for r in sheet["rows"]:
        ls.append([r["line"], r["sheet_row"], r["model"], r["qty"]])
    ls.append([])
    ls.append(["renamed.to replies per strip:"])
    for s in sheet.get("strips", []):
        ls.append([s["filename"]])
    for c in ls[1]:
        c.font = BOLD
    ls.column_dimensions["C"].width = 18
    for row in ws.iter_rows(min_row=hdr + 1):
        for c in row:
            c.alignment = Alignment(vertical="top")
    wb.save(path)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    src = ap.add_mutually_exclusive_group(required=True)
    src.add_argument("--sheet", help="loading sheet photo (read with renamed.to, uses credits)")
    src.add_argument("--sheet-json", help="saved result of an earlier loading_sheet_ocr run")
    ap.add_argument("--invoice", required=True, help="tax invoice PDF")
    ap.add_argument("--out", help="Excel report path (default: reconciliation_<invoice no>.xlsx)")
    args = ap.parse_args()

    inv = read_invoice(args.invoice)
    if args.sheet:
        from loading_sheet_ocr import read_sheet
        sheet = read_sheet(args.sheet)
        cache = os.path.splitext(args.sheet)[0] + ".ocr.json"
        with open(cache, "w", encoding="utf-8") as f:
            json.dump(sheet, f, indent=1)
        print(f"OCR saved to {cache} (rerun with --sheet-json to skip renamed.to)")
    else:
        with open(args.sheet_json, encoding="utf-8") as f:
            sheet = json.load(f)

    rows = reconcile(sheet, inv)
    out = args.out or f"reconciliation_{inv['invoice_no']}.xlsx"
    write_report(out, rows, sheet, inv)
    for r in rows:
        if r["status"] != "OK" or r["check"]:
            flag = " (check model)" if r["check"] else ""
            print(f"{r['status']:15} {r['sheet_model'] or '-':14} -> {r['invoice_model'] or '-':12} "
                  f"sheet {r['sheet_qty']:>3} invoice {r['invoice_qty']:>3}{flag}")
    print(f"sheet {sum(r['qty'] for r in sheet['rows'])} pcs, invoice {sum(l['qty'] for l in inv['lines'])} pcs "
          f"-> {out}")


if __name__ == "__main__":
    main()
