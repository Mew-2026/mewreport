"""Read the mastersheet: Item code / Material description / weight (kg) per piece.

The header row is found by its labels, so it may sit below a title and the columns may be in any
order. Every sheet of the workbook is searched; the first one with the three headers is used.
Accepts .xlsx / .xlsm / .csv.
"""
import csv
import io
import os
import re

from openpyxl import load_workbook

from invoice_parser import model_key


def _norm_code(v):
    return re.sub(r"\s+", "", str(v or "")).upper()


def _num(v):
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "").strip())
    except ValueError:
        return None


def _header(row):
    """Column indexes of (code, desc, kg) if this row is the header, else None."""
    cols = {}
    for i, v in enumerate(row):
        t = str(v or "").strip().lower()
        if "code" in t and "code" not in cols:
            cols["code"] = i
        elif "desc" in t and "desc" not in cols:
            cols["desc"] = i
        elif ("weight" in t or t in ("kg", "wt", "wt (kg)")) and "kg" not in cols:
            cols["kg"] = i
    return (cols["code"], cols["desc"], cols["kg"]) if len(cols) == 3 else None


def _tables(src, name):
    """Yield (sheet name, list of row tuples) for every sheet in the file."""
    if isinstance(src, (bytes, bytearray)):
        src = io.BytesIO(src)
    if os.path.splitext(name or "")[1].lower() == ".csv":
        data = src.read() if hasattr(src, "read") else open(src, "rb").read()
        text = data.decode("utf-8-sig", errors="replace")
        yield "csv", list(csv.reader(io.StringIO(text)))
        return
    wb = load_workbook(src, read_only=True, data_only=True)
    for ws in wb.worksheets:
        yield ws.title, [tuple(r) for r in ws.iter_rows(values_only=True)]


def read_master(src, name=None):
    """src: path, bytes or file object. Returns rows, a code index and a list of problems found."""
    name = name or (src if isinstance(src, str) else "")
    found = []
    for sheet, table in _tables(src, name):
        for h, row in enumerate(table[:30]):
            cols = _header(row)
            if cols:
                cols, notes = _data_columns(table, h, cols)
                master = _build(table, h, cols, sheet, notes)
                if master["rows"]:
                    return master
                found.append(sheet)
                break
    if found:
        raise ValueError(f"Mastersheet: sheet '{found[0]}' has the headers but no rows with an item code and a "
                         f"weight under them")
    raise ValueError("Mastersheet: no sheet has the headers 'Item code', 'Material description' and 'weight (kg)'")


CODE_RE = re.compile(r"^[A-Z]{2}[0-9A-Z]{8,14}$")   # SF1147C572150, SF1102J3407D5


def code_key(code):
    """Item code for matching: O/0 and I/1 count as the same (sheets made from a picture mix them up)."""
    return _norm_code(code).replace("O", "0").replace("I", "1")


def _fits(kind, v):
    if v is None or str(v).strip() == "":
        return None
    if kind == "code":
        c = _norm_code(v)
        return bool(CODE_RE.match(c)) and any(ch.isdigit() for ch in c)
    if kind == "kg":
        n = _num(v)
        return n is not None and n > 0
    t = str(v).strip()
    return " " in t and sum(ch.isalpha() for ch in t) >= 4   # desc


def _data_columns(table, h, cols):
    """The header gives each column's position, but in sheets converted from a picture the values can sit
    one column off ('Item code' over column A, the codes in column B). Use the column whose values look
    right, nearest to its header."""
    body = table[h + 1:h + 201]
    width = max((len(r) for r in body), default=0)

    def score(kind, i):
        got = [_fits(kind, r[i] if i < len(r) else None) for r in body]
        got = [g for g in got if g is not None]
        return sum(got) / len(got) if got else 0.0

    names = {"code": "Item code", "desc": "Material description", "kg": "weight (kg)"}
    picked, taken, notes = {}, set(), []
    for kind, hint in (("code", cols[0]), ("kg", cols[2]), ("desc", cols[1])):
        options = [i for i in range(width) if i not in taken and score(kind, i) >= 0.6]
        i = hint if hint in options else min(options, key=lambda i: abs(i - hint), default=hint)
        if i != hint:
            notes.append(f"'{names[kind]}' values read from column {_letter(i)} "
                         f"(header is in column {_letter(hint)})")
        picked[kind] = i
        taken.add(i)
    return (picked["code"], picked["desc"], picked["kg"]), notes


def _letter(i):
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def _build(table, h, cols, sheet, notes=()):
    ci, di, ki = cols
    rows, by_code, problems = [], {}, list(notes)
    for n, r in enumerate(table[h + 1:], start=h + 2):
        get = lambda i: r[i] if i < len(r) else None
        code = _norm_code(get(ci))
        if not code:
            continue
        desc = " ".join(str(get(di) or "").split())
        kg = _num(get(ki))
        if kg is None or kg <= 0:
            problems.append(f"Row {n}: {code} has no valid weight ({get(ki)!r}) - skipped")
            continue
        key = code_key(code)
        if key in by_code:
            first = by_code[key]
            if abs(first["kg"] - kg) > 1e-9:
                problems.append(f"Row {n}: {code} listed again with {kg} kg (row {first['row']} has "
                                f"{first['kg']} kg) - row {first['row']} used")
            continue
        m = {"row": n, "code": code, "desc": desc, "kg": kg, "model": model_key(desc)}
        rows.append(m)
        by_code[key] = m
    return {"sheet": sheet, "rows": rows, "by_code": by_code, "problems": problems}
