"""Draw a tax invoice in the exact layout of the Luminous scrap invoice (A4, 8 items per page).

Coordinates, column edges, fonts and sizes were measured from the real PDF. Text is drawn in the
same order as the real invoice's content stream, so PDF text extraction reads a generated invoice
the same way as a real one (item code, S.Loc, description, then the numbers).
"""
from reportlab.graphics import renderPDF
from reportlab.graphics.barcode import qr
from reportlab.graphics.shapes import Drawing
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

ITEMS_PER_PAGE = 8   # the real invoice fits 8 full rows per page
COPIES = 3           # the real PDF repeats the whole invoice three times

# Layout measured from the real invoice (A4, points from the top edge).
H = 842
COLS = [5.7, 22.6, 76.5, 168.1, 205.8, 248.2, 277.8, 316.3, 358.8, 397.6, 418.4, 448.7, 479.9, 513.0, 558.7]
TABLE_TOP, HEAD_BOTTOM, ROW_H, TOTAL_H = 334.9, 372.4, 32.1, 13.5
FOOTER_TOP, FOOTER_BOTTOM = 668.5, 742.8
R, B = "Helvetica", "Helvetica-Bold"


def money(x):
    return f"{x:,.2f}"


class Page:
    """Thin wrapper so coordinates can be given from the top edge, as measured."""

    def __init__(self, c):
        self.c = c

    def text(self, x, top, s, font=R, size=6, align="left"):
        c = self.c
        c.setFont(font, size)
        y = H - top - size * 0.78
        if align == "right":
            c.drawRightString(x, y, s)
        elif align == "center":
            c.drawCentredString(x, y, s)
        else:
            c.drawString(x, y, s)

    def label_value(self, x, top, label, value, value_x=None, value_font=R):
        self.text(x, top, label, B)
        if value:
            vx = value_x if value_x else x + self.c.stringWidth(label, B, 6) + 1.5
            self.text(vx, top, value, value_font)

    def hline(self, x0, x1, top):
        self.c.line(x0, H - top, x1, H - top)

    def vline(self, x, top0, top1):
        self.c.line(x, H - top0, x, H - top1)


def wrap(c, text, width, font=B, size=6):
    words, lines, cur = text.split(), [], ""
    for w in words:
        t = (cur + " " + w).strip()
        if c.stringWidth(t, font, size) <= width or not cur:
            cur = t
        else:
            lines.append(cur)
            cur = w
    return lines + [cur]


def draw_watermark(c, text):
    if not text:
        return
    c.saveState()
    c.setFillColorRGB(0.86, 0.86, 0.86)
    c.translate(300, 420)
    c.rotate(35)
    c.setFont(B, 64)
    c.drawCentredString(0, 0, text[0])
    c.setFont(B, 14)
    c.drawCentredString(0, -30, text[1])
    c.restoreState()


def draw_qr(c, data):
    widget = qr.QrCodeWidget(data)
    x0, y0, x1, y1 = widget.getBounds()
    size = 126
    d = Drawing(size, size, transform=[size / (x1 - x0), 0, 0, size / (y1 - y0), 0, 0])
    d.add(widget)
    renderPDF.draw(d, c, 414, H - 152)


def draw_page(c, inv, page_lines, page_no, pages, total):
    s, r = inv["supplier"], inv["recipient"]
    pg = Page(c)
    draw_watermark(c, inv.get("watermark"))
    c.setLineWidth(0.75)
    c.setFillColorRGB(0, 0, 0)

    # frame and header boxes
    pg.hline(5.3, 559.1, 1.5)
    pg.vline(5.7, 1.5, FOOTER_BOTTOM)
    pg.vline(558.8, 1.5, FOOTER_BOTTOM)
    pg.vline(382.7, 1.5, 177.9)
    pg.hline(382.4, 559.1, 21.4)
    pg.hline(5.3, 559.1, 177.9)
    pg.vline(285.7, 177.9, TABLE_TOP)
    pg.hline(5.3, 285.9, 255.8)
    pg.hline(285.4, 559.1, 209.3)
    pg.hline(285.4, 559.1, 255.8)
    pg.text(248.2, 4.0, "TAX INVOICE", B, 12, "center")
    draw_qr(c, inv["qr"])

    # supplier block
    pg.text(6.2, 47.7, "Name, Address & GSTIN of Supplier:", B)
    pg.label_value(215.4, 50.1, "Inv. No.:", inv["no"], 288.9, B)
    pg.text(6.2, 59.7, s["name"])
    pg.label_value(215.4, 64.3, "Issue Date:", inv["date"], 288.9, B)
    pg.text(6.2, 71.7, s["address"][0])
    pg.label_value(215.4, 78.4, "Page No.:", f"{page_no} of {pages}", 288.9, B)
    pg.text(6.2, 83.7, s["address"][1])
    pg.text(6.2, 95.7, s["state"])
    pg.text(6.2, 107.7, f"Ph No.: {s['phone']}")
    pg.text(6.2, 119.7, f"PAN No : {s['pan']}")
    pg.label_value(6.2, 149.6, "GSTIN/UIN No.:", s["gstin"])
    pg.text(6.2, 159.8, f"IRN No: {inv['irn']}", B)
    pg.text(289.1, 159.8, f"IRN Dt:{inv['irn_date']}", B)

    # recipient / delivery (left), order & transport (right)
    pg.text(6.2, 185.1, "Name, Address & GSTIN No of Recipient:", B)
    pg.text(6.5, 195.3, r["name"], B)
    pg.text(6.5, 205.5, r["address"][0])
    pg.text(6.5, 215.7, r["address"][1])
    pg.text(6.5, 225.9, f"Ph No.:     {r['phone']}    ,PAN No: {r['pan']}")
    pg.label_value(6.2, 242.0, "GSTIN STATE WITH CODE:", r["state"])
    pg.label_value(6.2, 248.0, "GSTIN/UIN NO.:", r["gstin"])
    pg.text(6.2, 263.0, "Name & Address(To Whom Goods to be Delivered):", B)
    pg.text(6.5, 273.2, r["name"], B)
    pg.text(6.5, 283.4, r["address"][0])
    pg.text(6.5, 293.6, r["address"][1] + "Bihar")
    pg.text(6.5, 303.8, f"Ph No.: {r['phone']}  ,PAN No:   {r['pan']}")
    pg.label_value(6.2, 320.0, "GSTIN STATE WITH CODE:", r["state"])
    pg.label_value(6.2, 326.0, "GSTIN/UIN NO.:", r["gstin"])

    pg.label_value(286.3, 185.1, "PO No.:", "Scrap Order", 310.1)
    pg.text(447.9, 185.1, "PO Date:", B)
    pg.label_value(286.3, 201.4, "Payment Terms:", "Immediate", 369.9)
    pg.text(286.3, 216.5, "Mode of Transport:", B)
    pg.label_value(442.2, 216.5, "Veh.No.:", inv.get("vehicle", ""), 470.0)
    pg.text(286.3, 232.4, "Name of Transporter:", B)
    pg.text(286.3, 246.5, "GR No.:")
    pg.text(443.1, 246.5, "GR Date:")
    pg.text(286.3, 263.0, "Insurance Policy No.:", B)
    pg.text(286.3, 276.9, "Way bill/Road permit No.:", B)
    pg.label_value(286.3, 297.9, "Incoterms:", "EXW - Ex Works", 329.6)
    pg.label_value(286.3, 312.1, "Order No.:", inv["order"], 332.5)
    pg.label_value(286.3, 325.4, "Challan/other doc.ref.No.:", inv["challan"], 361.0)

    # item table grid
    body_bottom = HEAD_BOTTOM + ROW_H * len(page_lines)
    table_bottom = body_bottom + (TOTAL_H if total else 0)
    for top in [TABLE_TOP, HEAD_BOTTOM] + [HEAD_BOTTOM + ROW_H * (i + 1) for i in range(len(page_lines))]:
        pg.hline(5.3, 559.1, top)
    if total:
        pg.hline(5.3, 559.1, table_bottom)
    for x in COLS:
        pg.vline(x, TABLE_TOP, table_bottom)

    head = [(11.0, "Sr"), (35.0, "Item Code"), (79.5, "Description & Specification of"), (180.6, "HSN"),
            (222.1, "Qty"), (256.0, "UOM"), (284.7, "Rate(Rs)"), (324.8, "Discount"), (367.1, "Taxable"),
            (401.5, "Rate"), (418.6, "CGST Amt"), (450.0, "SGST Amt"), (482.9, "IGST Amt"), (519.0, "Net Amount")]
    for x, t in head:
        pg.text(x, 342.1, t, B)
    for x, t in [(10.2, "No"), (77.9, "Goods & Replacement Note No"), (331.8, "(Rs)"), (370.2, "Value"),
                 (403.4, "(%)")]:
        pg.text(x, 354.1, t, B)
    pg.text(372.4, 366.1, "(Rs)", B)

    for i, l in enumerate(page_lines):
        top = HEAD_BOTTOM + ROW_H * i + 7.2
        pg.text(16.7, top, f" {l['sr']}", align="right")
        pg.text(23.2, top, f"{l['code']} /")
        pg.text(23.2, top + 12.0, "S.Loc / SC01")
        for k, t in enumerate(wrap(c, l["desc"], 88)[:2]):
            pg.text(77.0, top - 1.8 + 10.2 * k, t, B)
        pg.text(173.5, top, "85491100")
        pg.text(246.0, top, f"{l['qty']}", align="right")
        pg.text(263.0, top, "NOS", align="center")
        pg.text(313.5, top, money(l["rate"]), align="right")
        pg.text(395.0, top, money(l["taxable"]), align="right")
        pg.text(399.6, top, "18.00")
        pg.text(418.4, top, "0.00")
        pg.text(468.3, top, "0.00", align="right")
        pg.text(508.5, top, money(l["igst"]), align="right")
        pg.text(547.4, top, money(l["net"]), align="right")

    if total:
        top = body_bottom + 3.8
        pg.text(79.8, top, "Total", B)
        pg.text(246.0, top, str(total["qty"]), B, align="right")
        pg.text(356.0, top, "0.00", B, align="right")
        pg.text(395.0, top, money(total["taxable"]), B, align="right")
        pg.text(446.0, top, "0.00", B, align="right")
        pg.text(477.0, top, "0.00", B, align="right")
        pg.text(511.0, top, money(total["igst"]), B, align="right")
        pg.text(556.0, top, money(total["net"]), B, align="right")

    # remarks and footer
    pg.hline(5.3, 559.1, FOOTER_TOP)
    pg.hline(5.3, 559.1, FOOTER_BOTTOM)
    pg.vline(223.9, FOOTER_TOP - 1.4, FOOTER_BOTTOM)
    pg.vline(397.6, FOOTER_TOP - 1.4, FOOTER_BOTTOM)
    pg.text(6.2, 640.3, "Remarks: Taxable value included AGT Amount RS:0.00", B)
    pg.text(202.1, 658.4, "Whether the tax is payable on reverse charge basis: NO", B)
    pg.text(6.5, 661.9, "For terms and Conditions please read Overleaf", B)
    certified = [
        (669.3, 8.5, "Certified that the particulars given above are true & correct And the"),
        (676.9, 8.5, "amount indicated represents the price actually charged and that"),
        (684.5, 8.5, "there is no additional flow directly or indirectly from the buyer"),
        (692.7, 69.0, "OR"),
        (700.5, 8.5, "Certified that the particulars given above are true & correct & the"),
        (708.3, 8.5, "amount indicated is provisional consideration"),
    ]
    for top, x, t in certified:
        pg.text(x, top, t, R, 7)
    pg.text(227.8, 674.3, "Goods Dispatched  From:", B)
    pg.text(398.2, 674.3, "Plant Address:")
    from_lines = [s["name"], s["address"][0], s["address"][1], f"Ph No.: {s['phone']}", f"GSTIN : {s['gstin']}"]
    for k, t in enumerate(from_lines):
        pg.text(229.6, 683.9 + 9.6 * k, t, B)
        pg.text(403.3, 683.9 + 9.6 * k, t)


def write_invoice(path, inv):
    """inv: no, date, irn, irn_date, order, challan, qr, supplier, recipient, watermark,
    lines (sr, code, desc, qty, rate, taxable, igst, net). Returns the total qty."""
    lines = inv["lines"]
    pages = [lines[i:i + ITEMS_PER_PAGE] for i in range(0, len(lines), ITEMS_PER_PAGE)]
    total = {"qty": sum(l["qty"] for l in lines), "taxable": sum(l["taxable"] for l in lines),
             "igst": sum(l["igst"] for l in lines), "net": sum(l["net"] for l in lines)}
    c = canvas.Canvas(path, pagesize=A4)
    c.setTitle(f"Invoice {inv['no']}")
    for _ in range(COPIES):
        for i, page in enumerate(pages, 1):
            draw_page(c, inv, page, i, len(pages), total if i == len(pages) else None)
            c.showPage()
    c.save()
    return total["qty"]
