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
from export_table import to_pdf, to_xlsx  # noqa: E402
from loading_sheet_excel import read_sheet_excel, write_template  # noqa: E402
from loading_sheet_ocr import read_sheet  # noqa: E402
from rate_check import ordered, run_on_step1, write_excel, write_pdf  # noqa: E402
from rate_check import run as rate_run  # noqa: E402
from rate_check import report_name as rate_report_name  # noqa: E402
from reconcile import apply_groups, reconcile, report_name, write_report  # noqa: E402
from weight_recon import weight_recon, write_company_pdf, write_company_xlsx  # noqa: E402

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


@app.post("/api/export")
def export():
    """Any table on the page as Excel or PDF: body {"format": "xlsx" | "pdf", "spec": {...}} (see export_table)."""
    body = request.get_json(silent=True) or {}
    fmt, spec = body.get("format"), body.get("spec") or {}
    if fmt not in ("xlsx", "pdf"):
        return jsonify(error="format must be xlsx or pdf"), 400
    try:
        data = to_xlsx(spec) if fmt == "xlsx" else to_pdf(spec)
    except ValueError as e:
        return jsonify(error=str(e)), 422
    name = "".join(ch if ch.isalnum() or ch in "-_ " else "_" for ch in (spec.get("filename") or "table")).strip() or "table"
    mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if fmt == "xlsx" else "application/pdf"
    return data, 200, {"Content-Type": mime, "Content-Disposition": f'attachment; filename="{name}.{fmt}"'}


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
    """Step 2: weight & rate check.

    With the Step 1 result on the page (its sheet rows, invoices and manual groups) the report is the
    full one: weight reconciliation (loaded vs billed kg, split into Check 1 unit price and Check 2
    quantity), the master sheet with loaded qty in column D, the Step 1 table and the line-by-line rate
    check. Without Step 1, the parsed invoices get the rate check on their own."""
    master = request.files.get("master")
    if not master or not master.filename:
        return jsonify(error="Upload the mastersheet Excel"), 400
    try:
        invoices = json.loads(request.form.get("invoices", "[]"))
        sheet_rows = json.loads(request.form["sheet_rows"]) if request.form.get("sheet_rows") else None
        meta = json.loads(request.form.get("sheet_meta", "{}"))
        groups = json.loads(request.form.get("groups", "[]"))
    except ValueError:
        return jsonify(error="Bad invoice or loading sheet data"), 400
    if not invoices:
        return jsonify(error="Choose the invoice PDFs in Step 1"), 400
    try:
        rate = float(request.form.get("rate", ""))
    except ValueError:
        return jsonify(error="Enter today's Rate/Kg as a number"), 400
    try:
        master_bytes = master.read()
        weight = None
        if sheet_rows is not None:   # full check: rebuild Step 1 from the page's data (no credits)
            inv = combine_invoices(invoices, [i.get("file") for i in invoices])
            sheet = _sheet_from_rows(sheet_rows, meta)
            recon_rows = apply_groups(reconcile(sheet, inv), groups)
            step1 = io.BytesIO()
            write_report(step1, recon_rows, sheet, inv)
            rows, s, m, inv, wb, recon = run_on_step1(step1.getvalue(), master_bytes, master.filename, rate)
            weight = weight_recon(recon_rows, combine_invoices(invoices, [i.get("file") for i in invoices]), m, rate)
            name, source = rate_report_name(inv, "report"), "Step 1 result on the page"
        else:
            rows, s, m, inv = rate_run(master_bytes, master.filename, invoices, [i.get("file") for i in invoices],
                                       rate)
            wb = recon = None
            name, source = rate_report_name(inv), "invoice PDFs (rate check only)"
        xlsx, pdf = io.BytesIO(), io.BytesIO()
        write_excel(xlsx, rows, s, m, wb, weight)
        write_pdf(pdf, rows, s, recon, weight)
        report, report_x = io.BytesIO(), io.BytesIO()
        if weight:
            write_company_pdf(report, weight, s["invoices"])
            write_company_xlsx(report_x, weight, s["invoices"])
    except ValueError as e:
        return jsonify(error=str(e)), 422
    except Exception as e:
        return jsonify(error=f"{type(e).__name__}: {e}"), 500
    return jsonify(summary=s, rows=ordered(rows), master_problems=m["problems"], step1_file=source,
                   combined=weight is not None, filename=name,
                   weight=_json_weight(weight) if weight else None,
                   xlsx=base64.b64encode(xlsx.getvalue()).decode(), pdf=base64.b64encode(pdf.getvalue()).decode(),
                   report_pdf=base64.b64encode(report.getvalue()).decode() if weight else None,
                   report_xlsx=base64.b64encode(report_x.getvalue()).decode() if weight else None)


def _json_weight(w):
    r2 = lambda v: None if v is None else round(v, 2)
    keys = ("model", "sheet_model", "sheet_lines", "sheet_qty", "invoice_qty", "qty_diff", "group", "master_kg",
            "invoice_kg", "billed_kg", "loaded_kg", "price_kg", "qty_kg", "master_price", "invoice_price",
            "kgpc_gap", "price_status", "qty_status", "note", "kg_source")
    models = [{k: (r2(m.get(k)) if isinstance(m.get(k), float) else m.get(k)) for k in keys} for m in w["models"]]
    s = {k: v for k, v in w["summary"].items() if k != "sheet_not_in_master"}
    s["sheet_not_in_master"] = [{k: m[k] for k in ("model", "sheet_model", "sheet_lines", "sheet_qty", "master_kg", "kg_source")}
                                for m in w["summary"]["sheet_not_in_master"]]
    return {"summary": s, "models": models}
