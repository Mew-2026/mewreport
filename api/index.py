"""Vercel entry point: loading sheet photo + one or more invoice PDFs in, reconciliation Excel out.

Vercel caps a request body at ~4.5 MB, so the page uploads each invoice PDF to /api/invoice on its
own (parsed to JSON, no credits used) and then sends the sheet photo plus those JSON results to
/api/reconcile.
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
from loading_sheet_excel import read_sheet_excel  # noqa: E402
from loading_sheet_ocr import read_sheet  # noqa: E402
from rate_check import ordered, run_on_step1, write_excel, write_pdf  # noqa: E402
from rate_check import run as rate_run  # noqa: E402
from rate_check import report_name as rate_report_name  # noqa: E402
from reconcile import reconcile, report_name, write_report  # noqa: E402

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


@app.post("/api/reconcile")
def run():
    sheet_file = request.files.get("sheet")
    try:
        invoices = json.loads(request.form.get("invoices", "[]"))
    except ValueError:
        return jsonify(error="Bad invoice data"), 400
    if not sheet_file or not invoices:
        return jsonify(error="Upload the loading sheet (photo or Excel) and at least one invoice PDF"), 400
    is_excel = os.path.splitext(sheet_file.filename or "")[1].lower() in (".xlsx", ".xlsm", ".csv")
    try:
        inv = combine_invoices(invoices, [i.get("file") for i in invoices])
        if is_excel:  # typed loading sheet: read directly, no renamed.to credits
            sheet = read_sheet_excel(sheet_file.read(), sheet_file.filename)
        else:
            sheet = read_sheet(sheet_file.read(), log=lambda m: None)
        if not sheet["rows"]:
            return jsonify(error="Could not read any rows from the loading sheet"), 422
        rows = reconcile(sheet, inv)
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
