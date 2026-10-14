"""Read line items from a system-generated tax-invoice PDF (text layer, no OCR needed).

The item table is read by position: the vertical ruling lines give the column edges, the header
labels (Sr, Item Code, Description, HSN, Qty, Rate, Taxable ...) say which column is which, and
each row's words are collected per column. That does not depend on the order text is stored in
the PDF, so it reads the supplier's invoices and generated look-alikes the same way.

The PDF usually holds the same invoice 2-3 times (original/duplicate/triplicate); only the first
copy is read.

Usage: python invoice_parser.py <invoice.pdf>
"""
import re
import sys

import pdfplumber

# header label (first word, lowercase) -> field name
HEADER_FIELDS = {
    "sr": "sr", "item": "code", "description": "desc", "hsn": "hsn", "qty": "qty", "uom": "uom",
    "rate(rs)": "rate", "discount": "discount", "taxable": "taxable", "igst": "igst", "net": "net",
}


def model_key(desc):
    """Battery model from a description, e.g. 'S-LUM 12V150AH RC18000ST PRO FNSD BTY' -> 'RC18000STP'.

    Version tags (V21, R2, C ...) are ignored; Pro is kept as a trailing P because Pro and
    non-Pro are separate lines on the loading sheet.
    """
    s = desc.upper().replace(" ", "")
    pro = "PRO" in s
    for pat in [
        r"(ASTT12150H)",
        r"(LPTT121[25]0H)",
        r"(UCTT24066)",
        r"(ILT[TJ]\d{5})",
        r"BOOSTCHARGE(BC\d{5}(?:ST)?)",
        r"((?:PC|RC|EC|SC|LM|TC|TZ|NEO)\d{5}(?:STJ|TJ|ST|TT)?)",
        r"(?<!\d)([2-5]\d{3}(?:STJ|TT))",
    ]:
        m = re.search(pat, s)
        if m:
            return m.group(1) + ("P" if pro else "")
    return s


def _num(s):
    s = s.replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return None


def _columns(page, header_top):
    """Column edges from the vertical lines that cross the table header."""
    xs = sorted({round(l["x0"], 0) for l in page.lines
                 if abs(l["x0"] - l["x1"]) < 1 and l["top"] - 1 <= header_top + 4 <= l["bottom"] + 1})
    edges = []
    for x in xs:
        if not edges or x - edges[-1] > 3:
            edges.append(x)
    return edges


def _table_layout(page, words):
    """Find the table header and map each column to a field. Returns (header_bottom, columns)."""
    sr = next((w for w in words if w["text"] == "Sr"), None)
    hsn = next((w for w in words if w["text"] == "HSN"), None)
    if not sr or not hsn:
        return None
    header_top = min(sr["top"], hsn["top"])
    edges = _columns(page, header_top)
    if len(edges) < 6:
        return None
    header_words = [w for w in words if header_top - 2 <= w["top"] <= header_top + 2]
    cols = []
    for x0, x1 in zip(edges, edges[1:]):
        inside = sorted((w for w in header_words if x0 <= (w["x0"] + w["x1"]) / 2 < x1), key=lambda w: w["x0"])
        label = inside[0]["text"].lower() if inside else ""
        cols.append((x0, x1, HEADER_FIELDS.get(label)))
    # header ends where the first body text starts: the lowest header-band word ("(Rs)" etc.)
    band = [w for w in words if header_top - 2 <= w["top"] <= header_top + 30
            and w["text"] in ("No", "(Rs)", "(%)", "Value", "Goods")]
    header_bottom = max((w["bottom"] for w in band), default=header_top + 10)
    return header_bottom, cols


def _cell_text(ws):
    ws = sorted(ws, key=lambda w: (round(w["top"]), w["x0"]))
    lines, cur, top = [], [], None
    for w in ws:
        if top is not None and abs(w["top"] - top) > 2:
            lines.append(" ".join(cur))
            cur = []
        cur.append(w["text"])
        top = w["top"]
    if cur:
        lines.append(" ".join(cur))
    return lines


def _page_rows(page):
    """Items on one page: list of dicts, plus description text that continues the previous page's last item."""
    words = page.extract_words(x_tolerance=1.5, y_tolerance=2, keep_blank_chars=False, extra_attrs=["size"])
    words = [w for w in words if w["size"] <= 9]  # table text is 6-7pt; skips titles and watermarks
    layout = _table_layout(page, words)
    if not layout:
        return [], [], None
    header_bottom, cols = layout
    stop = min((w["top"] for w in words if w["top"] > header_bottom and w["text"] in ("Total", "Remarks:")),
               default=page.height)
    body = [w for w in words if header_bottom < w["top"] < stop - 1]

    def field_of(w):
        cx = (w["x0"] + w["x1"]) / 2
        for x0, x1, f in cols:
            if x0 <= cx < x1:
                return f
        return None

    starts = sorted((w["top"], int(w["text"])) for w in body if field_of(w) == "sr" and w["text"].isdigit())
    # words above the first numbered row belong to an item continued from the previous page
    bands = [(header_bottom, starts[0][0] - 4 if starts else stop, None)]
    for i, (top, sr) in enumerate(starts):
        nxt = starts[i + 1][0] - 4 if i + 1 < len(starts) else stop
        bands.append((top - 4, nxt, sr))

    rows, carry = [], []
    for y0, y1, sr in bands:
        cells = {}
        for w in body:
            if y0 <= w["top"] < y1:
                cells.setdefault(field_of(w), []).append(w)
        if sr is None:
            carry = _cell_text(cells.get("desc", []))
            continue
        code_lines = _cell_text(cells.get("code", []))
        code = code_lines[0].rstrip(" /") if code_lines else ""
        qty = _num(" ".join(_cell_text(cells.get("qty", []))))
        rows.append({
            "sr": sr,
            "code": code.split()[0] if code else "",
            "desc": " ".join(_cell_text(cells.get("desc", []))),
            "qty": int(qty) if qty is not None else None,
            "rate": _num(" ".join(_cell_text(cells.get("rate", [])))),
            "value": _num(" ".join(_cell_text(cells.get("taxable", [])))),
        })

    total = None
    tw = next((w for w in words if w["text"] == "Total" and w["top"] > header_bottom), None)
    if tw:
        qty_col = next(((x0, x1) for x0, x1, f in cols if f == "qty"), None)
        for w in words:
            if (qty_col and abs(w["top"] - tw["top"]) < 3 and w["text"].isdigit()
                    and qty_col[0] <= (w["x0"] + w["x1"]) / 2 < qty_col[1]):
                total = int(w["text"])
    return rows, carry, total


def read_invoice(path):
    with pdfplumber.open(path) as pdf:
        lines, total, texts = {}, None, []
        last = None
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            m = re.search(r"Page No\.:\s*(\d+)\s*of\s*(\d+)", text)
            if i and m and m.group(1) == "1":
                break  # second copy starts
            texts.append(text)
            rows, carry, page_total = _page_rows(page)
            if carry and last is not None:
                lines[last]["desc"] = (lines[last]["desc"] + " " + " ".join(carry)).strip()
            for r in rows:
                lines[r["sr"]] = r
                last = r["sr"]
            total = page_total if page_total is not None else total
            if m and m.group(1) == m.group(2):
                break  # last page of the first copy

        header = pdf.pages[0]
        recipient = None
        label = header.search(r"Name, Address & GSTIN No of Recipient:")
        if label:
            top = label[0]["bottom"]
            left = header.crop((0, top + 0.5, label[0]["x1"] + 120, min(top + 14, header.height)))
            recipient = (left.extract_text() or "").split("\n")[0].strip() or None

    items = []
    for sr in sorted(lines):
        l = lines[sr]
        if l["qty"] is None:
            continue
        items.append(dict(l, desc=" ".join(l["desc"].split()), model=model_key(l["desc"])))
    all_text = "\n".join(texts)
    return {
        "invoice_no": _find(r"Inv\. No\.:\s*(\d+)", all_text),
        "date": _find(r"Issue Date:\s*([\d.]+)", all_text),
        "recipient": recipient,
        "lines": items,
        "printed_total_qty": total,
    }


def combine_invoices(invoices, names=None):
    """Merge several parsed invoices for one load. The same invoice uploaded twice is counted once."""
    names = names or [None] * len(invoices)
    seen, kept, skipped = set(), [], []
    for inv, name in zip(invoices, names):
        key = inv["invoice_no"] or name
        if key in seen:
            skipped.append(name or key)
            continue
        seen.add(key)
        kept.append(dict(inv, file=name))
    lines = [dict(l, invoice_no=inv["invoice_no"]) for inv in kept for l in inv["lines"]]
    return {"invoices": kept, "lines": lines, "duplicates": skipped}


def _find(pat, text):
    m = re.search(pat, text)
    return m.group(1).strip() if m else None


if __name__ == "__main__":
    inv = read_invoice(sys.argv[1])
    for l in inv["lines"]:
        print(l["sr"], l["code"], l["model"], l["qty"], l["desc"], sep="\t")
    got = sum(l["qty"] for l in inv["lines"])
    print(f"invoice {inv['invoice_no']} {inv['date']} {inv['recipient']}: "
          f"{len(inv['lines'])} lines, qty {got}, printed total {inv['printed_total_qty']}")
