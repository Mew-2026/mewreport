"""Generate DEMO tax-invoice PDFs in the same layout as the Luminous scrap invoice.

The quantities come from the 13/10/26 loading sheet (Patiala -> Muzaffarpur, 440 pcs), so the
reconciliation can be tested end to end. Supplier/recipient names and invoice numbers are made up
and every page carries a DEMO watermark: these are test files, not tax documents.

Two sets are written to demo/:
  match_*.pdf     two invoices that together match the loading sheet exactly
  mismatch_*.pdf  the same load with three deliberate errors (see MISMATCH below)

Usage: python demo/make_demo_invoices.py
"""
import os
import textwrap

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

OUT = os.path.dirname(os.path.abspath(__file__))

# model key -> (item code, description, rate). Codes, descriptions and rates follow the real
# invoice; the last four models were not on it and are made up.
CATALOG = {
    "2042STJ": ("SF1147B570150", "S-AMAZE 12V150AH 2042STJ FNSD BTY", 3637.48),
    "2042TT": ("SF1147Q715150", "S-AMAZE 12V150AH 2042TT V22 FNSD BTY C", 3637.48),
    "2048STJ": ("SF1247B313150", "SCR-AMAZE 12V150AH 2048STJ FNSD BTY", 3637.48),
    "2048TT": ("SF1147Q942150", "S-AMAZE 12V150AH 2048TT FNSD BTY V24 C", 3637.48),
    "2060STJ": ("SF1147C831150", "S-AMAZE 12V150AH 2060STJ FNSD BTY", 3741.66),
    "2060TT": ("SF1147C698150", "S-AMAZE 12V150AH 2060TT V22 FNSD BTY", 3741.66),
    "2136TT": ("SF1147Q712150", "S-AMAZE 12V150AH 2136TT V22 FNSD BTY C", 3741.66),
    "2160TT": ("SF1147L711150", "S-AMAZE 12V150AH 2160TT FNSD BTY", 3741.66),
    "2260TT": ("SF1147Q815160", "S-AMAZE 12V 160AH 2260TT FNSD BTY C", 3993.50),
    "4060TT": ("SF1147C856180", "S-AMAZE 12V180AH 4060TT V23 FNSD BTY", 4489.13),
    "5048TT": ("SF1147C873200", "S-AMAZE 12V200AH 5048TT V24 FNSD BTY", 4987.92),
    "5060TT": ("SF1147C874200", "S-AMAZE 12V200AH 5060TT V24 FNSD BTY", 4987.92),
    "ASTT12150H": ("SF1147C645150", "S-AMAZE 12V150AH ASTT12150H V21 FNSD SLR", 4557.15),
    "BC16048ST": ("SF1156C834135", "S-LUM Boost Charge BC 16048ST FNSD BTY", 3081.04),
    "BC18048ST": ("SF1156C835150", "S-LUM Boost Charge BC 18048ST FNSD BTY", 3417.86),
    "BC18060": ("SF1156C870150", "S-LUM Boost Charge BC 18060 FNSD BTY", 3417.86),
    "EC18060": ("SF1139C828150", "S-LUM 12V 150AH EC18060 V23 FNSD BTY", 3681.70),
    "ILTJ18148": ("SF1103NXR0150", "SCR LUM 12V 150AH ILTJ18148 BTY", 3800.87),
    "ILTT18060": ("SF1103R720150", "S-LUM 12V150AH ILTT 18060 V22 C FNSD BTY", 3681.70),
    "ILTT18060P": ("SF1103C737150", "S-LUM 12V150AH ILTT 18060 Pro FNSD BTY", 3681.70),
    "ILTT20060": ("SF1103Q814160", "S-LUM 12V160AH ILTT 20060 FNSD BTY C", 3693.69),
    "LM18075": ("SF1145C577150", "S-LUM 12V150AH LM18075 V20 R3 FNSD BTY", 3681.70),
    "LPTT12120H": ("SF1113C623120", "S-LUM 12V120AH LPTT12120H FNSD V21 R1 SL", 3585.76),
    "LPTT12150H": ("SF1113C655150", "S-LUM 12V150AH LPTT12150H FNSD V21 R2 SL", 4557.15),
    "NEO18030": ("SF1154Q816150", "S-LUM 12V150AH NEO 18030 FNSD BTY C", 3417.86),
    "PC18054P": ("SF1135Q807150", "S-LUM 12V150AH PC18054 PRO FNSD BTY C", 3741.66),
    "PC18054TJP": ("SF1135C804150", "S-LUM 12V150AH PC18054TJ Pro FNSD BTY", 3741.66),
    "RC18000": ("SF1138Q657150", "S-LUM 12V150AH RC18000 FNSD BTY V21 R1 C", 3681.70),
    "RC18000ST": ("SF1138M204150", "SCR-LUM 12V150AH RC18000ST FNSD BTY", 3681.70),
    "RC18000STP": ("SF1138C800150", "S-LUM 12V150AH RC18000ST PRO FNSD BTY", 3681.70),
    "RC24000P": ("SF1138C801180", "S-LUM 12V180AH RC24000 PRO FNSD BTY", 4489.13),
    "RC25000": ("SF1138C648200", "S-LUM 12V200AH RC25000 V21 FNSD BTY", 4987.92),
    "RC25000P": ("SF1138C802200", "S-LUM 12V200AH RC25000 PRO FNSD BTY", 4987.92),
    "SC16060": ("SF1144L671135", "S-LUM 12V135AH SC 16060 FNSD BTY", 3081.04),
    "SC18060": ("SF1144L668150", "S-LUM 12V150AH SC 18060 FNSD BTY", 3417.86),
    "TC13548ST": ("SF1147C832135", "S-AMAZE TC13548ST FNSD BTY", 2698.31),
    "TC15048ST": ("SF1147C833150", "S-AMAZE TC15048ST FNSD BTY", 3417.86),
    "TC15060": ("SF1147C898150", "S-AMAZE TC15060 FNSD BTY", 3417.86),
    "TZ13548ST": ("SF1108C991135", "S-TEZ - TZ 13548ST FNSD BTY", 2698.31),
    "UCTT24066": ("SF1154C853180", "S-LUM 12V 180AH UCTT 24066 V23 FNSD BTY", 4489.13),
    "RC18048ST": ("SF1138C805150", "S-LUM 12V150AH RC18048ST FNSD BTY", 3681.70),
    "RC18000P": ("SF1138C806150", "S-LUM 12V150AH RC18000 PRO FNSD BTY", 3681.70),
    "ILTT24066": ("SF1103C899200", "S-LUM 12V200AH ILTT 24066 FNSD BTY", 4987.92),
    "PC20042TT": ("SF1135Q714160", "S-LUM 12V160AH PC20042TT V21 FNSD BTY C", 3993.50),
}

# The loading sheet of 13/10/26, in sheet order (crossed-out row left out).
SHEET = [
    ("2042STJ", 3), ("2042TT", 2), ("2048STJ", 9), ("2048TT", 1), ("2060STJ", 6), ("2060TT", 17),
    ("2136TT", 1), ("2160TT", 5), ("2260TT", 3), ("4060TT", 1), ("5048TT", 12), ("5060TT", 1),
    ("ASTT12150H", 2), ("BC16048ST", 48), ("BC18048ST", 24), ("BC18060", 5), ("EC18060", 14),
    ("ILTJ18148", 18), ("ILTT18060", 3), ("ILTT18060P", 7), ("ILTT20060", 2), ("LM18075", 3),
    ("LPTT12120H", 34), ("LPTT12150H", 103), ("NEO18030", 1), ("PC18054P", 7), ("PC18054TJP", 8),
    ("RC18000", 1), ("RC18000ST", 7), ("RC18000STP", 3), ("RC24000P", 1), ("RC25000", 2),
    ("RC25000P", 14), ("SC16060", 40), ("SC18060", 10), ("TC13548ST", 8), ("TC15048ST", 5),
    ("TC15060", 1), ("TZ13548ST", 2), ("UCTT24066", 2), ("RC18048ST", 1), ("RC18000P", 2),
    ("ILTT24066", 1),
]


def split_load(sheet):
    """Two invoices for one truck: first 23 sheet rows, then the rest. LPTT12150H (103 pcs) is billed
    60 + 43 across both invoices and BC16048ST as two lines, like the real invoice does."""
    inv1, inv2 = [], []
    for i, (model, qty) in enumerate(sheet):
        if model == "LPTT12150H":
            inv1.append((model, 60))
            inv2.append((model, qty - 60))
        elif model == "BC16048ST":
            inv1 += [(model, 39), (model, qty - 39)]
        else:
            (inv1 if i < 23 else inv2).append((model, qty))
    return inv1, inv2


# Deliberate errors for the mismatch set: SC16060 short-billed by 2, RC18000P not billed,
# PC20042TT (crossed out on the sheet) billed anyway.
def mismatch(inv1, inv2):
    inv2 = [(m, q - 2 if m == "SC16060" else q) for m, q in inv2 if m != "RC18000P"]
    return inv1, inv2 + [("PC20042TT", 1)]


SUPPLIER = [
    "DEMO BATTERY WORKS PVT. LTD. (SPECIMEN)",
    "Plot 1, Demo Industrial Estate",
    "Patiala-140417 INDIA",
    "GSTIN State with Code: 03 - Punjab",
    "Ph No.: 0000000000",
]
RECIPIENT = [
    "DEMO RECYCLERS (SPECIMEN)",
    "Plot 2, Demo Industrial Area,",
    "MUZAFFARPUR - 842005,",
    "GSTIN STATE WITH CODE: BIHAR - 10",
]
ITEMS_PER_PAGE = 8
COPIES = ["ORIGINAL FOR RECIPIENT", "DUPLICATE FOR TRANSPORTER"]


def money(x):
    return f"{x:,.2f}"


def lines_for(items):
    out = []
    for sr, (model, qty) in enumerate(items, 1):
        code, desc, rate = CATALOG[model]
        taxable = round(rate * qty, 2)
        igst = round(taxable * 0.18, 2)
        out.append({"sr": sr, "code": code, "desc": desc, "qty": qty, "rate": rate,
                    "taxable": taxable, "igst": igst, "net": round(taxable + igst, 2)})
    return out


def draw_page(c, inv, page_lines, page_no, pages, copy, total_line):
    w, h = A4
    # watermark
    c.saveState()
    c.setFillGray(0.88)
    c.setFont("Helvetica-Bold", 70)
    c.translate(w / 2, h / 2)
    c.rotate(35)
    c.drawCentredString(0, 0, "DEMO")
    c.setFont("Helvetica-Bold", 18)
    c.drawCentredString(0, -40, "NOT A VALID TAX INVOICE")
    c.restoreState()

    y = h - 36
    c.setFont("Helvetica-Bold", 13)
    c.drawCentredString(w / 2, y, " TAX INVOICE")
    c.setFont("Helvetica", 7)
    c.drawRightString(w - 30, y, copy)

    c.setFont("Helvetica", 7.5)
    y -= 18
    header = [
        f"Mode of Transport: Road Veh.No.: {inv['vehicle']}",
        "Name of Transporter:",
        "GR No.: GR Date: ",
        "Insurance Policy No.: ",
        "Way bill/Road permit No.:",
        f"Challan/other doc.ref.No.: {inv['challan']}",
        f"Order No.: {inv['order']} ",
        "Incoterms: EXW - Ex Works",
        "Payment Terms: Immediate",
        "PO No.: Scrap Order PO Date:Name, Address & GSTIN No of Recipient:",
    ]
    for t in header:
        c.drawString(30, y, t)
        y -= 9.5
    for t in RECIPIENT:
        c.drawString(30, y, t)
        y -= 9.5

    # table header; the last header cell must read "Net Amount" (the parser keys on it)
    y -= 6
    c.line(30, y + 8, w - 30, y + 8)
    c.setFont("Helvetica-Bold", 7)
    c.drawString(30, y, "Sr")
    c.drawString(30, y - 8, "No")
    c.drawString(52, y, "Item Code Description & Specification of")
    c.drawString(52, y - 8, "Goods & Replacement Note No")
    c.drawString(200, y - 16, "HSN Qty UOM Rate(Rs) Discount")
    c.drawString(330, y - 16, "(Rs)")
    c.drawString(360, y - 16, "Taxable")
    c.drawString(360, y - 24, "Value")
    c.drawString(360, y - 32, "(Rs)")
    c.drawString(400, y - 32, "Rate")
    c.drawString(400, y - 40, "(%)")
    c.drawString(430, y - 40, "CGST Amt SGST Amt IGST Amt Net Amount")
    y -= 50
    c.line(30, y + 6, w - 30, y + 6)

    c.setFont("Helvetica", 7.5)
    for l in page_lines:
        y -= 4
        c.drawString(30, y, f"{l['sr']} {l['code']} /")
        c.drawString(52, y - 9, "S.Loc / SC01")
        desc = textwrap.wrap(l["desc"], 26)
        for k, t in enumerate(desc):
            c.drawString(52, y - 18 - 9 * k, t)
        y -= 18 + 9 * len(desc)
        # one string so text extraction keeps the column order of the real invoice
        c.drawString(52, y, f"85491100                {l['qty']} NOS {money(l['rate'])} {money(l['taxable'])} "
                            f"18.00 0.00 0.00 {money(l['igst'])} {money(l['net'])} ")
        y -= 12

    if total_line:
        y -= 4
        c.line(30, y + 10, w - 30, y + 10)
        c.setFont("Helvetica-Bold", 7.5)
        c.drawString(30, y, total_line)
        y -= 14

    c.setFont("Helvetica", 6.8)
    footer = [
        "Certified that the particulars given above are true & correct And the",
        "amount indicated represents the price actually charged and that",
        "there is no additional flow directly or indirectly from the buyer",
        "Goods Dispatched From: DEMO BATTERY WORKS PVT. LTD. (SPECIMEN)",
        "Remarks: DEMO DOCUMENT FOR SOFTWARE TESTING ONLY",
        "Whether the tax is payable on reverse charge basis: NO",
        "Name, Address & GSTIN of Supplier:",
        *SUPPLIER,
        f"Inv. No.: {inv['no']}",
        f"Issue Date: {inv['date']}",
        f"Page No.: {page_no} of {pages}",
    ]
    fy = 150
    for t in footer:
        c.drawString(30, fy, t)
        fy -= 8.5
    c.showPage()


def write_invoice(path, inv):
    lines = lines_for(inv["items"])
    pages = [lines[i:i + ITEMS_PER_PAGE] for i in range(0, len(lines), ITEMS_PER_PAGE)]
    qty = sum(l["qty"] for l in lines)
    total = (f"Total             {qty} 0.00 {money(sum(l['taxable'] for l in lines))} 0.00 0.00 "
             f"{money(sum(l['igst'] for l in lines))} {money(sum(l['net'] for l in lines))} ")
    c = canvas.Canvas(path, pagesize=A4)
    c.setTitle(f"DEMO invoice {inv['no']}")
    for copy in COPIES:
        for i, page in enumerate(pages, 1):
            draw_page(c, inv, page, i, len(pages), copy, total if i == len(pages) else None)
    c.save()
    return qty


def main():
    base = {"date": "13.10.2026", "vehicle": "DEMO-0000", "order": "DEMO-ORDER"}
    match1, match2 = split_load(SHEET)
    sets = {
        "match": (match1, match2),
        "mismatch": mismatch(match1, match2),
    }
    n = 1
    for name, invoices in sets.items():
        for k, items in enumerate(invoices, 1):
            no = f"9000000000{n:02d}"  # all digits like real invoice numbers; 9000... is not a real series
            inv = dict(base, no=no, challan=f"DEMO/{no}", items=items)
            path = os.path.join(OUT, f"{name}_{k}_{no}.pdf")
            qty = write_invoice(path, inv)
            print(f"{os.path.basename(path)}: {len(items)} lines, {qty} pcs")
            n += 1


if __name__ == "__main__":
    main()
