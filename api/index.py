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
from loading_sheet_ocr import read_sheet  # noqa: E402
from reconcile import reconcile, report_name, write_report  # noqa: E402

app = Flask(__name__)


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
        return jsonify(error="Upload the loading sheet photo and at least one invoice PDF"), 400
    try:
        inv = combine_invoices(invoices, [i.get("file") for i in invoices])
        sheet = read_sheet(sheet_file.read(), log=lambda m: None)
        if not sheet["rows"]:
            return jsonify(error="Could not read any rows from the loading sheet"), 422
        rows = reconcile(sheet, inv)
        buf = io.BytesIO()
        write_report(buf, rows, sheet, inv)
    except Exception as e:  # surface renamed.to / parsing errors to the page
        return jsonify(error=f"{type(e).__name__}: {e}"), 500

    return jsonify(
        invoices=[{k: i.get(k) for k in ("invoice_no", "date", "recipient", "file", "printed_total_qty")}
                  for i in inv["invoices"]],
        duplicates=inv["duplicates"],
        sheet_total=sheet["total"],
        invoice_total=sum(l["qty"] for l in inv["lines"]),
        rows=rows,
        filename=report_name(inv),
        xlsx=base64.b64encode(buf.getvalue()).decode(),
    )
