"""Vercel entry point: upload a loading sheet photo + invoice PDF, get the reconciliation back."""
import base64
import io
import os
import sys

from flask import Flask, jsonify, request, send_from_directory

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from invoice_parser import read_invoice  # noqa: E402
from loading_sheet_ocr import read_sheet  # noqa: E402
from reconcile import reconcile, write_report  # noqa: E402

app = Flask(__name__)


@app.get("/")
def index():
    return send_from_directory(os.path.join(ROOT, "public"), "index.html")


@app.post("/api/reconcile")
def run():
    sheet_file, invoice_file = request.files.get("sheet"), request.files.get("invoice")
    if not sheet_file or not invoice_file:
        return jsonify(error="Upload both the loading sheet photo and the invoice PDF"), 400
    try:
        inv = read_invoice(io.BytesIO(invoice_file.read()))
        if not inv["lines"]:
            return jsonify(error="No line items found in the invoice PDF"), 422
        sheet = read_sheet(sheet_file.read(), log=lambda m: None)
        if not sheet["rows"]:
            return jsonify(error="Could not read any rows from the loading sheet"), 422
        rows = reconcile(sheet, inv)
        buf = io.BytesIO()
        write_report(buf, rows, sheet, inv)
    except Exception as e:  # surface renamed.to / parsing errors to the page
        return jsonify(error=f"{type(e).__name__}: {e}"), 500

    return jsonify(
        invoice_no=inv["invoice_no"],
        date=inv["date"],
        recipient=inv["recipient"],
        sheet_total=sheet["total"],
        invoice_total=sum(l["qty"] for l in inv["lines"]),
        rows=rows,
        filename=f"reconciliation_{inv['invoice_no']}.xlsx",
        xlsx=base64.b64encode(buf.getvalue()).decode(),
    )
