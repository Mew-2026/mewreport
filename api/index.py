"""Vercel entry point: loading sheet (photo or Excel) + one or more invoice PDFs in, reconciliation Excel out.

Vercel caps a request body at ~4.5 MB, so the page uploads each invoice PDF to /api/invoice on its
own (parsed to JSON, no credits used). The loading sheet is read by /api/sheet and shown on the page
for checking and editing; /api/reconcile then gets the edited rows plus the parsed invoices.
"""
import base64
import io
import json
import os
import sys

from flask import Flask, jsonify, request, send_from_directory

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from invoice_parser import combine_invoices, read_invoice  # noqa: E402
from loading_sheet_excel import read_sheet_excel, write_template  # noqa: E402
from loading_sheet_ocr import read_sheet  # noqa: E402
from rate_check import ordered, run_on_step1, write_excel, write_pdf  # noqa: E402
from rate_check import run as rate_run  # noqa: E402
from rate_check import report_name as rate_report_name  # noqa: E402
from reconcile import apply_groups, reconcile, report_name, write_report  # noqa: E402

app = Flask(__name__, static_folder=os.path.join(ROOT, "public"), static_url_path="")  # local runs; Vercel serves public/ itself


@app.get("/")
def index():
    return send_from_directory(os.path.join(ROOT, "public"), "index.html")


@app.post("/api/invoice")
def parse_invoice():
    f = request.files.get("invoice")
    if not f:
        return jsonify(error="No invoice PDF uploaded"), 400
    try:
        inv = read_invoice(io.BytesIO(f.read()))
    except Exception as e:
        return jsonify(error=f"{f.filename}: could not read PDF ({type(e).__name__}: {e})"), 422
    if not inv["lines"]:
        return jsonify(error=f"{f.filename}: no line items found"), 422
    return jsonify(dict(inv, file=f.filename))


def _read_upload(f):
    """Loading sheet from an uploaded file: Excel / CSV read directly (no credits), a photo via renamed.to."""
    if os.path.splitext(f.filename or "")[1].lower() in (".xlsx", ".xlsm", ".csv"):
        return read_sheet_excel(f.read(), f.filename)
    return dict(read_sheet(f.read(), log=lambda m: None), source="photo")


def _sheet_from_rows(raw, base=None):
    """Loading sheet from the rows checked / edited on the page."""
    rows = []
    for r in raw:
        model = "".join(str(r.get("model") or "").split()).upper()
        try:
            qty = int(float(r.get("qty")))
        except (TypeError, ValueError):
            qty = 0
        if not model and not qty:
            continue  # empty row left on the page
        if not model or qty <= 0:
            raise ValueError(f"Loading sheet row {len(rows) + 1}: needs a model and a quantity above 0")
        rows.append({"line": len(rows) + 1, "sheet_row": r.get("sheet_row") or len(rows) + 1,
                     "model": model, "qty": qty})
    base = base or {}
    return {"rows": rows, "total": sum(r["qty"] for r in rows), "strips": base.get("strips", []),
            "source": base.get("source", "edited")}


@app.post("/api/sheet")
def read_loading_sheet():
    """Read the loading sheet only, so the page can show it for checking and editing."""
    f = request.files.get("sheet")
    if not f or not f.filename:
        return jsonify(error="Choose the loading sheet (photo or Excel)"), 400
    try:
        sheet = _read_upload(f)
    except ValueError as e:
        return jsonify(error=str(e)), 422
    except Exception as e:  # renamed.to errors
        return jsonify(error=f"{type(e).__name__}: {e}"), 500
    if not sheet["rows"]:
        return jsonify(error="Could not read any rows from the loading sheet - add them by hand"), 422
    return jsonify(sheet)


@app.post("/api/sheet-excel")
def sheet_excel():
    """The checked / edited loading sheet as Excel, to keep and reuse without reading the photo again."""
    try:
        sheet = _sheet_from_rows(json.loads(request.form.get("rows", "[]")))
    except ValueError as e:
        return jsonify(error=str(e)), 422
    buf = io.BytesIO()
    write_template(buf, [(r["model"], r["qty"]) for r in sheet["rows"]])
    return buf.getvalue(), 200, {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": 'attachment; filename="loading_sheet.xlsx"'}


@app.post("/api/reconcile")
def run():
    sheet_file = request.files.get("sheet")
    try:
        invoices = json.loads(request.form.get("invoices", "[]"))
        edited = json.loads(request.form["sheet_rows"]) if request.form.get("sheet_rows") else None
        meta = json.loads(request.form.get("sheet_meta", "{}"))
        groups = json.loads(request.form.get("groups", "[]"))   # manual matches made on the page
    except ValueError:
        return jsonify(error="Bad invoice or loading sheet data"), 400
    if (edited is None and not sheet_file) or not invoices:
        return jsonify(error="Add the loading sheet (photo or Excel) and at least one invoice PDF"), 400
    try:
        inv = combine_invoices(invoices, [i.get("file") for i in invoices])
        sheet = _sheet_from_rows(edited, meta) if edited is not None else _read_upload(sheet_file)
        if not sheet["rows"]:
            return jsonify(error="Could not read any rows from the loading sheet"), 422
        rows = apply_groups(reconcile(sheet, inv), groups)
        buf = io.BytesIO()
        write_report(buf, rows, sheet, inv)
    except ValueError as e:  # e.g. Excel without the Model / Qty headers
        return jsonify(error=str(e)), 422
    except Exception as e:  # surface renamed.to / parsing errors to the page
        return jsonify(error=f"{type(e).__name__}: {e}"), 500

    return jsonify(
        invoices=[{k: i.get(k) for k in ("invoice_no", "date", "recipient", "file", "printed_total_qty")}
                  for i in inv["invoices"]],
        duplicates=inv["duplicates"],
        sheet_total=sheet["total"],
        sheet_source=sheet.get("source", "photo"),
        invoice_total=sum(l["qty"] for l in inv["lines"]),
        rows=rows,
        filename=report_name(inv),
        xlsx=base64.b64encode(buf.getvalue()).decode(),
    )


@app.post("/api/ratecheck")
def rate_check():
    """Step 2: invoice rates vs mastersheet kg x today's Rate/Kg.

    With the Step 1 report (the Step 1 result on the page) the output is one combined report: the Step 1
    Excel plus a Rate check tab, and a PDF of both. Without it, the parsed invoices (from /api/invoice)
    are checked on their own - no loading sheet needed."""
    step1, master = request.files.get("step1"), request.files.get("master")
    if not master or not master.filename:
        return jsonify(error="Upload the mastersheet Excel"), 400
    try:
        invoices = json.loads(request.form.get("invoices", "[]"))
    except ValueError:
        return jsonify(error="Bad invoice data"), 400
    has_step1 = bool(step1 and step1.filename)
    if not has_step1 and not invoices:
        return jsonify(error="Choose the invoice PDFs in Step 1 (or run Step 1 first)"), 400
    try:
        rate = float(request.form.get("rate", ""))
        tol = float(request.form.get("tolerance", "0.5") or 0.5)
    except ValueError:
        return jsonify(error="Enter today's Rate/Kg and the tolerance as numbers"), 400
    try:
        if has_step1:
            rows, s, m, inv, wb, recon = run_on_step1(step1.read(), master.read(), master.filename, rate, tol)
            name, source = rate_report_name(inv, "report"), step1.filename
        else:
            rows, s, m, inv = rate_run(master.read(), master.filename, invoices, [i.get("file") for i in invoices],
                                  rate, tol)
            wb = recon = None
            name, source = rate_report_name(inv), "invoice PDFs (rate check only)"
        xlsx, pdf = io.BytesIO(), io.BytesIO()
        write_excel(xlsx, rows, s, m, wb)
        write_pdf(pdf, rows, s, recon)
    except ValueError as e:
        return jsonify(error=str(e)), 422
    except Exception as e:
        return jsonify(error=f"{type(e).__name__}: {e}"), 500
    return jsonify(summary=s, rows=ordered(rows), master_problems=m["problems"], step1_file=source,
                   combined=has_step1, filename=name,
                   xlsx=base64.b64encode(xlsx.getvalue()).decode(), pdf=base64.b64encode(pdf.getvalue()).decode())
