"""Match a handwritten loading sheet against the tax invoice(s) for the load and write an Excel report.

Usage:
  python reconcile.py --sheet loading_sheet.jpg --invoice inv1.pdf [inv2.pdf ...] [--out report.xlsx]
  python reconcile.py --sheet-json sheet.json --invoice inv1.pdf ...   # reuse an earlier OCR run (no credits)
"""
import argparse
import json
import os
import re
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from rapidfuzz import fuzz

from invoice_parser import combine_invoices, read_invoice

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


def refs(lines, multi):
    """Invoice Sr. Nos, grouped by invoice when the load has several invoices."""
    if not multi:
        return ", ".join(str(sr) for _, sr in lines)
    by_inv = defaultdict(list)
    for no, sr in lines:
        by_inv[no].append(str(sr))
    return "; ".join(f"{no}: {', '.join(srs)}" for no, srs in by_inv.items())


def reconcile(sheet, inv):
    """inv: result of invoice_parser.combine_invoices (one or more invoices for the load)."""
    multi = len(inv["invoices"]) > 1
    agg = defaultdict(lambda: {"qty": 0, "lines": [], "codes": set(), "descs": set()})
    for l in inv["lines"]:
        a = agg[l["model"]]
        a["qty"] += l["qty"]
        a["lines"].append((l["invoice_no"], l["sr"]))
        a["codes"].add(l["code"])
        a["descs"].add(l["desc"])
    by_row = match(sheet["rows"], list(agg))

    def row(line, sheet_model, k, score, sq, check):
        a = agg.get(k)
        iq = a["qty"] if a else 0
        if not sheet_model:
            status = "NOT IN SHEET"
        elif not a:
            status = "NOT IN INVOICE"
        elif iq == sq:
            status = "OK"
        else:
            status = "QTY MISMATCH"
        return {
            "line": line, "sheet_model": sheet_model, "invoice_model": k or "", "score": score,
            "sheet_qty": sq, "invoice_qty": iq, "diff": iq - sq, "status": status,
            "result": "CORRECT" if status == "OK" else "INCORRECT", "check": check,
            "invoices": ", ".join(dict.fromkeys(no for no, _ in a["lines"])) if a else "",
            "invoice_lines": refs(a["lines"], multi) if a else "",
            "item_codes": ", ".join(sorted(a["codes"])) if a else "",
            "description": " / ".join(sorted(a["descs"])) if a else "",
        }

    out = []
    for i, r in enumerate(sheet["rows"]):
        k, score = by_row.get(i, (None, 0))
        out.append(row(r["line"], r["model"], k, round(score), r["qty"], k is not None and score < CHECK_SCORE))
    matched = {o["invoice_model"] for o in out}
    out += [row("", "", k, "", 0, False) for k in agg if k not in matched]
    return out


def apply_groups(rows, groups):
    """Manual matches made on the page. groups: [{"rows": [row indexes], "note": str}].

    Rows grouped together are judged on their combined difference: net 0 -> MATCHED, otherwise
    PARTLY MATCHED with the remaining difference. Only rows that were INCORRECT can be grouped."""
    used = set()
    for n, g in enumerate(groups or [], 1):
        idx = sorted({int(i) for i in g.get("rows", [])})
        if len(idx) < 2:
            raise ValueError(f"Group {n}: choose at least 2 rows")
        for i in idx:
            if not 0 <= i < len(rows):
                raise ValueError(f"Group {n}: unknown row")
            if i in used:
                raise ValueError(f"Group {n}: a row can only be in one group")
            if rows[i]["result"] != "INCORRECT":
                raise ValueError(f"Group {n}: only incorrect rows can be matched by hand")
        used.update(idx)
        net = sum(rows[i]["diff"] for i in idx)
        note = " ".join(str(g.get("note") or "").split())[:200]
        for i in idx:
            rows[i].update(group=f"G{n}", group_net=net, note=note,
                           result="MATCHED" if net == 0 else "PARTLY MATCHED")
    return rows


RESULT_ORDER = {"INCORRECT": 0, "PARTLY MATCHED": 1, "MATCHED": 2, "CORRECT": 3}
RESULT_FILL = {"CORRECT": PatternFill("solid", fgColor="C6EFCE"), "INCORRECT": PatternFill("solid", fgColor="FFC7CE"),
               "MATCHED": PatternFill("solid", fgColor="DDEBF7"), "PARTLY MATCHED": PatternFill("solid", fgColor="FFF2CC")}
RESULT_FONT = {"CORRECT": Font(bold=True, color="006100"), "INCORRECT": Font(bold=True, color="9C0006"),
               "MATCHED": Font(bold=True, color="1F4E79"), "PARTLY MATCHED": Font(bold=True, color="7F6000")}
CHECK_FILL = PatternFill("solid", fgColor="BDD7EE")
BOLD = Font(bold=True)


def write_report(path, rows, sheet, inv):
    wb = Workbook()
    ws = wb.active
    ws.title = "Reconciliation"
    sheet_total = sum(r["qty"] for r in sheet["rows"])
    inv_total = sum(l["qty"] for l in inv["lines"])
    count = {k: sum(r["result"] == k for r in rows) for k in RESULT_ORDER}
    open_ = count["INCORRECT"] + count["PARTLY MATCHED"]

    ws.append(["LOADING SHEET vs INVOICE CHECK"])
    ws["A1"].font = Font(bold=True, size=14)
    ws.append(["Overall", "ALL CORRECT" if not open_ else f"{open_} INCORRECT of {len(rows)}"])
    ws.append(["Loading sheet pcs", sheet_total, "Invoice pcs", inv_total, "Difference", inv_total - sheet_total])
    ws.append(["Correct rows", count["CORRECT"], "Incorrect rows", count["INCORRECT"],
               "Matched by hand", count["MATCHED"], "Partly matched", count["PARTLY MATCHED"]])
    for row in ws.iter_rows(min_row=2, max_row=4):
        for c in row[::2]:
            c.font = BOLD
    ws["B2"].font = RESULT_FONT["INCORRECT" if open_ else "CORRECT"]
    ws.append([])
    ws.append(["Invoice no", "Date", "Recipient", "Lines", "Pcs read", "Printed total", "Read check", "File"])
    for c in ws[ws.max_row]:
        c.font = BOLD
    for i in inv["invoices"]:
        read = sum(l["qty"] for l in i["lines"])
        printed = i.get("printed_total_qty")
        ok = printed is not None and printed == read
        ws.append([i["invoice_no"], i["date"], i["recipient"], len(i["lines"]), read, printed,
                   "OK" if ok else "CHECK: pcs read differ from printed total", i.get("file") or ""])
        ws.cell(ws.max_row, 7).font = RESULT_FONT["CORRECT" if ok else "INCORRECT"]
    for d in inv.get("duplicates", []):
        ws.append([f"Skipped duplicate upload: {d}"])
    ws.append([])

    cols = [("result", "Result"), ("status", "Reason"), ("line", "Sheet line"),
            ("sheet_model", "Sheet model (as read)"), ("invoice_model", "Invoice model"), ("score", "Match %"),
            ("sheet_qty", "Sheet qty"), ("invoice_qty", "Invoice qty"), ("diff", "Diff (inv - sheet)"),
            ("group", "Group"), ("group_net", "Group net"), ("note", "Note"), ("invoices", "Invoice no"), ("invoice_lines", "Invoice Sr. Nos"), ("item_codes", "Item codes"),
            ("description", "Invoice description")]
    ws.append([h for _, h in cols])
    hdr = ws.max_row
    for c in ws[hdr]:
        c.font = BOLD
    # incorrect rows first, then partly / fully matched by hand, then correct; sheet order inside each
    for r in sorted(rows, key=lambda r: (RESULT_ORDER[r["result"]], r.get("group", ""), r["line"] == "", r["line"] or 0)):
        ws.append([r.get(k, "") for k, _ in cols])
        cells = ws[ws.max_row]
        for c in cells[:9]:
            c.fill = RESULT_FILL[r["result"]]
        cells[0].font = RESULT_FONT[r["result"]]
        if r["check"]:
            for c in cells[3:6]:
                c.fill = CHECK_FILL
        for c in cells:
            c.alignment = Alignment(vertical="top")
    last = ws.max_row
    ws.append(["", "", "", "TOTAL", "", "", f"=SUM(G{hdr + 1}:G{last})", f"=SUM(H{hdr + 1}:H{last})",
               f"=SUM(I{hdr + 1}:I{last})"])
    for c in ws[ws.max_row]:
        c.font = BOLD
    ws.append([])
    ws.append([f"Blue model cells: fuzzy match below {CHECK_SCORE}% - check the sheet model was read and paired correctly."])
    ws.append(["Compare 'Loading sheet pcs' with the total written on the paper: a misread row changes it."])
    groups = {}
    for r in rows:
        if r.get("group"):
            groups.setdefault(r["group"], []).append(r)
    if groups:
        ws.append([])
        ws.append(["Manual matches (made on the page)"])
        ws.cell(ws.max_row, 1).font = BOLD
        ws.append(["Group", "Result", "Net diff", "Rows (sheet model / invoice model : diff)", "", "", "", "", "", "Note"])
        for c in ws[ws.max_row]:
            c.font = BOLD
        for g, rs in sorted(groups.items(), key=lambda kv: int(kv[0][1:])):
            parts = ", ".join(f"{r['sheet_model'] or '-'} / {r['invoice_model'] or '-'} : {r['diff']:+d}" for r in rs)
            ws.append([g, rs[0]["result"], rs[0]["group_net"], parts, "", "", "", "", "", rs[0].get("note", "")])
            ws.cell(ws.max_row, 2).font = RESULT_FONT[rs[0]["result"]]
    ws.auto_filter.ref = f"A{hdr}:P{last}"
    ws.freeze_panes = ws.cell(row=hdr + 1, column=1)
    for col, w in zip("ABCDEFGHIJKLMNOP", [16, 16, 10, 22, 16, 9, 10, 11, 15, 8, 10, 30, 15, 22, 34, 60]):
        ws.column_dimensions[col].width = w

    il = wb.create_sheet("Invoice lines")
    il.append(["Invoice no", "Sr", "Item code", "Description", "Model", "Qty", "Rate", "Taxable value"])
    for l in inv["lines"]:
        il.append([l["invoice_no"], l["sr"], l["code"], l["desc"], l["model"], l["qty"], l["rate"], l["value"]])
    for c in il[1]:
        c.font = BOLD
    for col, w in zip("ABCDEFGH", [15, 6, 18, 48, 14, 7, 10, 14]):
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
    wb.save(path)


def report_name(inv):
    nos = [i["invoice_no"] for i in inv["invoices"] if i["invoice_no"]]
    if not nos:
        return "reconciliation.xlsx"
    return "reconciliation_" + "_".join(nos[:3]) + ("_etc" if len(nos) > 3 else "") + ".xlsx"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    src = ap.add_mutually_exclusive_group(required=True)
    src.add_argument("--sheet", help="loading sheet: photo (read with renamed.to, uses credits) or .xlsx/.csv (Sr No | Model | Qty)")
    src.add_argument("--sheet-json", help="saved result of an earlier loading_sheet_ocr run")
    ap.add_argument("--invoice", required=True, nargs="+", help="one or more tax invoice PDFs for the load")
    ap.add_argument("--out", help="Excel report path (default: reconciliation_<invoice nos>.xlsx)")
    args = ap.parse_args()

    inv = combine_invoices([read_invoice(p) for p in args.invoice], [os.path.basename(p) for p in args.invoice])
    for d in inv["duplicates"]:
        print(f"skipped duplicate invoice: {d}")
    if args.sheet:
        if os.path.splitext(args.sheet)[1].lower() in (".xlsx", ".xlsm", ".csv"):
            from loading_sheet_excel import read_sheet_excel   # typed sheet: no OCR, no credits
            sheet = read_sheet_excel(args.sheet)
        else:
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
    out = args.out or report_name(inv)
    write_report(out, rows, sheet, inv)
    for r in rows:
        if r["result"] != "CORRECT" or r["check"]:
            flag = " (check model)" if r["check"] else ""
            print(f"{r['status']:15} {r['sheet_model'] or '-':14} -> {r['invoice_model'] or '-':12} "
                  f"sheet {r['sheet_qty']:>3} invoice {r['invoice_qty']:>3}{flag}")
    correct = sum(r["result"] == "CORRECT" for r in rows)
    print(f"{correct} correct, {len(rows) - correct} incorrect; sheet {sum(r['qty'] for r in sheet['rows'])} pcs, "
          f"invoices {sum(l['qty'] for l in inv['lines'])} pcs -> {out}")


if __name__ == "__main__":
    main()
