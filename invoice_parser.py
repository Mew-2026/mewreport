"""Read line items from a Luminous scrap tax-invoice PDF (text layer, no OCR needed).

The PDF usually holds the same invoice 2-3 times (original/duplicate/triplicate);
only the first copy is read.

Usage: python invoice_parser.py <invoice.pdf>
"""
import re
import sys

from pypdf import PdfReader

ITEM_START = re.compile(r"(?:^|\s)(\d{1,3}) (S[A-Z]\w{10,12}) /", re.M)
QTY = re.compile(r"\d{8}\s+(\d+) NOS ([\d,]+\.\d+) ([\d,]+\.\d+)")


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


def read_invoice(path):
    pages = [p.extract_text() or "" for p in PdfReader(path).pages]
    header = pages[0]
    copy = []
    for i, text in enumerate(pages):
        m = re.search(r"Page No\.: (\d+) of (\d+)", text)
        if i and m and m.group(1) == "1":
            break  # second copy starts
        copy.append(text)

    body = ""
    for text in copy:
        m = re.search(r"Net Amount\n(.*?)(?:\nTotal |Certified that|\nPage No|\Z)", text, re.S)
        body += "\n" + (m.group(1) if m else "")

    # An item is "<sr> <code> /" ... "<hsn> <qty> NOS <rate> <value>"; descriptions can wrap and
    # an item can be split over a page break, so pair each item start with the next quantity.
    events = [(m.start(), "item", m) for m in ITEM_START.finditer(body)]
    events += [(m.start(), "qty", m) for m in QTY.finditer(body)]
    events.sort(key=lambda e: e[0])
    lines, cur = {}, None
    for pos, kind, m in events:
        if kind == "item":
            cur = m
        elif cur:
            desc = body[cur.end():pos]
            desc = " ".join(re.sub(r"S\.Loc / \S+", "", desc).split())
            sr = int(cur.group(1))
            lines[sr] = {
                "sr": sr,
                "code": cur.group(2),
                "desc": desc,
                "model": model_key(desc),
                "qty": int(m.group(1)),
                "rate": float(m.group(2).replace(",", "")),
                "value": float(m.group(3).replace(",", "")),
            }
            cur = None

    total = re.search(r"\nTotal\s+(\d+) ", "\n".join(copy))
    return {
        "invoice_no": _find(r"Inv\. No\.: (\d+)", "\n".join(copy)),
        "date": _find(r"Issue Date: ([\d.]+)", "\n".join(copy)),
        "recipient": _find(r"Name, Address & GSTIN No of Recipient:\n(.+)", header),
        "lines": [lines[k] for k in sorted(lines)],
        "printed_total_qty": int(total.group(1)) if total else None,
    }


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
