"""Weight reconciliation: why does the invoice bill more (or fewer) kg than were loaded?

  Loaded weight = sum over models of  loading-sheet qty x master kg/pc
  Billed weight = sum over invoice lines of  qty x invoice kg/pc,  invoice kg/pc = rate / Rate per kg to 2 decimals
  Variance      = billed - loaded, split exactly into
    Check 1  unit price / weight :  invoice qty x (invoice kg/pc - master kg/pc)   per model
    Check 2  quantity            :  (invoice qty - sheet qty) x master kg/pc        per model
    + billed kg of models with no weight at all (not in master, and nothing billed to take a weight from)

A model missing from the master is weighed with its invoice kg/pc (rate / Rate per kg), so its loaded pcs
still count; when it was not billed, the invoice kg/pc of its manual-match partner or of the closest billed
model (same Pro / non-Pro) is used. Its unit price cannot be checked, so Check 1 is 0 for it.

Works on the Step 1 rows (reconcile.reconcile, one row per model) and the parsed invoice lines, so
nothing is read again.
"""
import re
from collections import Counter, defaultdict

from rapidfuzz import fuzz

from invoice_parser import model_sort_key
from mastersheet import code_key
from reconcile import SUGGEST_SCORE


def _model_kgs(master, model):
    """kg/pc of master rows whose description gives this model (most common first)."""
    return [k for k, _ in Counter(m["kg"] for m in master["rows"] if m["model"] == model).most_common()]


def _similar(key, candidates):
    """Billed model to weigh an unbilled model that is not in the master: one with the same model number
    (Pro with Pro preferred: RC18000P -> RC18000STP, RC26000 -> RC26000P, TZ15060STJ -> TC15060), else the
    closest name (Pro only with Pro)."""
    pro = key.endswith("P")
    num = re.search(r"\d{4,5}", key)
    same = [c for c in candidates if c != key and num and re.search(r"\d{4,5}", c)
            and re.search(r"\d{4,5}", c).group() == num.group()]
    if same:
        return max(same, key=lambda c: (c.endswith("P") == pro, fuzz.ratio(key, c)))
    best = max(((fuzz.ratio(key, c), c) for c in candidates if c != key and c.endswith("P") == pro), default=(0, None))
    return best[1] if best[0] >= SUGGEST_SCORE else None


def weight_recon(recon_rows, inv, master, rate_per_kg):
    """One model can sit under several item codes with different weights (and rates), so the unit price is
    checked line by line against each line's own reference kg/pc; the model's kg/pc for the loaded weight and
    the quantity part is the average of its billed lines' reference weights (weighted by qty), or the master's
    weight when not billed.

    Reference kg/pc of an invoice line: its item code's master weight; an item code missing from the master
    takes the master row of the same model with the closest description (rate_check.find_by_desc); a model
    missing from the master altogether takes rate / Rate per kg of its main line (flagged NOT IN MASTER)."""
    from rate_check import find_by_desc

    lines_by_model = defaultdict(list)
    for l in inv["lines"]:
        lines_by_model[l["model"]].append(l)
    # invoice kg/pc per billed model, for models not in the master: rate / Rate per kg (not rounded), from the
    # line with the most pcs (the first such line on a tie)
    inv_kg = {}
    for mod, ls in lines_by_model.items():
        priced = [l for l in ls if l.get("rate")]
        if priced:
            inv_kg[mod] = max(priced, key=lambda l: l["qty"])["rate"] / rate_per_kg
    group_models = defaultdict(list)   # manual match group -> invoice models in it
    for r in recon_rows:
        if r.get("group") and r["invoice_model"]:
            group_models[r["group"]].append(r["invoice_model"])

    def reference(l, fallback):
        """(kg/pc, master item code, how) the line's billed kg/pc is checked against."""
        m_ = master["by_code"].get(code_key(l["code"]))
        if m_:
            return m_["kg"], m_["code"], "item code"
        row, _ = find_by_desc(master, l["desc"])
        if row:
            return row["kg"], row["code"], "same model (item code not in master)"
        if fallback:
            return fallback[0], "", "same model (item code not in master)"
        return None, "", ""

    models, not_in_master = [], []
    for r in recon_rows:
        key = r.get("model") or r["invoice_model"] or r["sheet_model"]
        lines = lines_by_model.get(r["invoice_model"], []) if r["invoice_model"] else []
        fallback = _model_kgs(master, key)
        line_kg = [(l, reference(l, fallback)) for l in lines]
        known_lines = [(l, ref[0]) for l, ref in line_kg if ref[0] is not None]
        sq, iq = r["sheet_qty"], r["invoice_qty"]
        source = "master"
        if known_lines:
            kg = sum(l["qty"] * k for l, k in known_lines) / sum(l["qty"] for l, _ in known_lines)
        elif fallback:
            kg = fallback[0]
        else:   # not in the master: weigh with the invoice kg/pc instead of leaving the pcs out
            kg, source = None, None
            if r["invoice_model"] in inv_kg:
                kg, source = inv_kg[r["invoice_model"]], "invoice"
            else:
                partner = next((g for g in group_models.get(r.get("group"), []) if g in inv_kg), None) \
                    or _similar(key, inv_kg)
                if partner:
                    kg, source = inv_kg[partner], f"invoice {partner}"
        kgs = sorted({round(k, 4) for _, k in known_lines} | set(fallback if not known_lines else []))
        priced = [(l, ref[0] if ref[0] is not None else kg) for l, ref in line_kg if l.get("rate")]
        billed = sum(l["qty"] * round(l["rate"] / rate_per_kg, 2) for l, _ in priced)   # kg/pc to 2 dp, like the master
        pq = sum(l["qty"] for l, _ in priced)
        # each invoice line: billed kg/pc against its reference kg/pc
        detail = []
        for l, ref in line_kg:
            ref_kg, ref_code, how = ref if ref[0] is not None else (kg, "", "not in master: main invoice rate"
                                                                     if source == "invoice" else "")
            bkg = round(l["rate"] / rate_per_kg, 2) if l.get("rate") else None
            d = {"invoice_no": l["invoice_no"], "sr": l["sr"], "code": l["code"], "desc": l["desc"], "model": key,
                 "qty": l["qty"], "rate": l.get("rate"), "invoice_kg": bkg, "ref_kg": ref_kg, "ref_code": ref_code,
                 "ref_how": how, "gap": None, "effect_kg": None, "status": "NO RATE" if bkg is None else "NOT IN MASTER"}
            if bkg is not None and ref_kg is not None:
                d["gap"] = round(bkg - round(ref_kg, 2), 2) + 0.0
                d["effect_kg"] = l["qty"] * d["gap"]
                d["status"] = "OK" if abs(d["gap"]) < 0.005 else ("PRICE HIGH" if d["gap"] > 0 else "PRICE LOW")
            detail.append(d)
        m = {"model": key, "sheet_model": r["sheet_model"], "sheet_lines": r.get("sheet_lines", ""),
             "sheet_qty": sq, "invoice_qty": iq, "qty_diff": iq - sq, "group": r.get("group", ""),
             "master_kg": kg, "master_kgs": kgs, "invoice_kg": billed / pq if pq else None, "billed_kg": billed,
             "kg_source": source, "loaded_kg": None, "price_kg": None, "qty_kg": None,
             "master_price": kg * rate_per_kg if kg else None,
             "invoice_price": sum(l["qty"] * l["rate"] for l, _ in priced) / pq if pq else None,
             "lines": detail}
        if kg is None:
            m["price_status"] = "NOT IN MASTER"
            m["qty_status"] = "NOT IN MASTER"
            not_in_master.append(m)
        else:
            m["loaded_kg"] = sq * kg
            m["qty_kg"] = (iq - sq) * kg
            m["price_kg"] = billed - iq * kg   # = sum over lines of qty x (invoice kg - its code's kg); keeps the split exact
            # unit price: each line's billed kg/pc (2 dp) against its own item code's master kg/pc; worst line decides
            gaps = [d["gap"] for d in detail if d["gap"] is not None]
            if source != "master":
                m["price_status"] = "NOT IN MASTER"
                m["note"] = (f"NOT IN MASTER - weighed at {kg:.2f} kg/pc = invoice rate / {rate_per_kg:g}"
                             + ("" if source == "invoice" else f" of {source[8:]} (similar model, not billed itself)"))
                not_in_master.append(m)
            elif not gaps:
                m["price_status"] = "NOT BILLED"
            else:
                worst = max(gaps, key=abs)
                m["kgpc_gap"] = round(worst, 2)
                m["price_status"] = "OK" if abs(worst) < 0.005 else ("PRICE HIGH" if worst > 0 else "PRICE LOW")
            m["qty_status"] = "OK" if iq == sq else ("MORE BILLED" if iq > sq else "LESS BILLED")
            if source == "master" and len(kgs) > 1:
                m["note"] = (f"item codes weigh {', '.join(f'{x:g}' for x in kgs)} kg - each invoice line uses its own; "
                             f"{kg:.2f} kg/pc average used for the sheet qty")
        models.append(m)

    known = [m for m in models if m["master_kg"] is not None]
    loaded = sum(m["loaded_kg"] for m in known)
    billed = sum(m["billed_kg"] for m in models)
    price = sum(m["price_kg"] for m in known)
    qty = sum(m["qty_kg"] for m in known)
    unknown_billed = sum(m["billed_kg"] for m in models if m["master_kg"] is None)
    # whole kg everywhere; the quantity part takes up any 1 kg rounding so the parts always add up on paper
    loaded_r, billed_r, price_r, nim_r = round(loaded), round(billed), round(price), round(unknown_billed)
    summary = {
        "rate_per_kg": rate_per_kg,
        "loaded_kg": loaded_r, "billed_kg": billed_r, "variance_kg": billed_r - loaded_r,
        "price_kg": price_r, "qty_kg": billed_r - loaded_r - price_r - nim_r, "not_in_master_kg": nim_r,
        "loaded_pcs": sum(m["sheet_qty"] for m in models), "billed_pcs": sum(m["invoice_qty"] for m in models),
        "loaded_pcs_weighed": sum(m["sheet_qty"] for m in known),   # sheet pcs with a master weight
        "price_issues": sum(m["price_status"] in ("PRICE HIGH", "PRICE LOW") for m in models),
        "qty_issues": sum(m["qty_status"] in ("MORE BILLED", "LESS BILLED") for m in models),
        "not_in_master": len(not_in_master),
        "sheet_not_in_master": [m for m in not_in_master if m["sheet_qty"]],   # weighed from the invoice, or not at all
    }
    # the parts must add up to the variance; anything else is a bug
    assert abs((price + qty + unknown_billed) - (billed - loaded)) < 0.01, "weight split does not add up"
    return {"summary": summary, "models": models}


# ---------------------------------------------------------------- Excel
def add_weight_sheets(wb, w, master):
    """Adds 'Weight reconciliation' (summary, Check 1, Check 2) and 'Master sheet (loaded)' tabs."""
    from openpyxl.styles import Font, PatternFill

    bold, title = Font(bold=True), Font(bold=True, size=14)
    red, green, amber = (PatternFill("solid", fgColor=c) for c in ("FFC7CE", "C6EFCE", "FFF2CC"))
    kgf = "#,##0"   # whole kg; kg/pc keeps 2 decimals
    s = w["summary"]

    ws = wb.create_sheet("Weight reconciliation", 0 if "Reconciliation" not in wb.sheetnames else 1)
    ws.append(["WEIGHT RECONCILIATION - loaded vs billed"])
    ws["A1"].font = title
    rows = [
        ("Today's Rate/Kg", s["rate_per_kg"]), ("Unit price match", "invoice kg/pc = master kg/pc (2 decimals)"),
        ("Loaded weight (sheet qty x master kg)", s["loaded_kg"]), ("Billed weight (invoice rate / Rate/Kg x qty)", s["billed_kg"]),
        ("VARIANCE (billed - loaded)", s["variance_kg"]),
        ("  of which Check 1 - unit price / weight", s["price_kg"]), ("  of which Check 2 - quantity", s["qty_kg"]),
        ("  of which invoice models not in master", s["not_in_master_kg"]),
    ]
    for label, v in rows:
        ws.append([label, v])
        ws.cell(ws.max_row, 1).font = bold
        if isinstance(v, (int, float)) and label != "Today's Rate/Kg":
            ws.cell(ws.max_row, 2).number_format = kgf
    ws.cell(6, 2).font = Font(bold=True, color="9C0006" if abs(s["variance_kg"]) >= 0.01 else "006100")
    ws.append(["Loaded pcs", s["loaded_pcs"], "Billed pcs", s["billed_pcs"]])
    ws.append([])

    def table(head, data, fmts, status_col=None):
        ws.append(head)
        for c in ws[ws.max_row]:
            c.font = bold
        for d in data:
            ws.append(d)
            for c, f in zip(ws[ws.max_row], fmts):
                if f:
                    c.number_format = f
            if status_col is not None:
                st = d[status_col]
                fill = green if st == "OK" else amber if st in ("NOT IN MASTER", "NOT BILLED") else red
                ws.cell(ws.max_row, status_col + 1).fill = fill

    models = sorted(w["models"], key=lambda m: model_sort_key(m["model"]))
    ws.append(["CHECK 1 - Unit price, line by line: invoice rate / Rate/Kg vs standard kg/pc (master)"])
    ws.cell(ws.max_row, 1).font = Font(bold=True, size=12)
    rate = s["rate_per_kg"]
    table(["Model", "Status", "Invoice / Sr", "Item code", "Rate (Rs)", "Billed kg/pc", "Standard kg/pc",
           "Diff kg/pc", "Pcs", "Effect (kg)", "Effect (Rs)", "Standard from"],
          [[d["model"], d["status"], f"{d['invoice_no']} / {d['sr']}", d["code"], d["rate"], d["invoice_kg"], d["ref_kg"],
            d["gap"], d["qty"], d["effect_kg"], None if d["effect_kg"] is None else d["effect_kg"] * rate,
            " ".join(x for x in (d["ref_code"], d["ref_how"]) if x)] for m in models for d in m["lines"]],
          [None, None, None, None, "#,##0.00", "0.00", "0.00", "+0.00;-0.00;0.00", "0", "+0.0;-0.0;0", kgf, None],
          status_col=1)
    ws.append(["Total (with rounding of kg/pc)", "", "", "", "", "", "", "", "", s["price_kg"], s["price_kg"] * rate])
    for c in ws[ws.max_row]:
        c.font = bold
    ws.append([])
    ws.append(["CHECK 2 - Quantity: loading sheet vs invoice, converted to kg with master kg/pc"])
    ws.cell(ws.max_row, 1).font = Font(bold=True, size=12)
    table(["Model", "Status", "Sheet model", "Sheet line", "Sheet qty", "Invoice qty", "Diff (inv - sheet)", "Master kg/pc", "Effect (kg)", "Group / note"],
          [[m["model"], m["qty_status"], m["sheet_model"], m["sheet_lines"], m["sheet_qty"], m["invoice_qty"], m["qty_diff"],
            m["master_kg"], m["qty_kg"], " · ".join(x for x in (m["group"], m.get("note", "")) if x)] for m in models],
          [None, None, None, None, "0", "0", "+0;-0;0", "0.00", kgf, None], status_col=1)
    for col, wd in zip("ABCDEFGHIJKL", [16, 16, 20, 18, 12, 12, 14, 12, 8, 12, 12, 50]):
        ws.column_dimensions[col].width = wd

    # master sheet with column D = loaded qty from the loading sheet
    ms = wb.create_sheet("Master sheet (loaded)")
    ms.append(["Item code", "Material description", "weight (kg)", "Loaded qty (nos)", "Loaded kg", "Model"])
    for c in ms[1]:
        c.font = bold
    loaded = {m["model"]: m for m in w["models"] if m["kg_source"] == "master" and m["sheet_qty"]}
    done = set()
    for r in master["rows"]:
        m = loaded.get(r["model"])
        first = m is not None and r["model"] not in done   # a model's qty goes on its first master row
        if first:
            done.add(r["model"])
        ms.append([r["code"], r["desc"], r["kg"], m["sheet_qty"] if first else None,
                   round(m["loaded_kg"], 2) if first else None, r["model"]])   # same kg/pc as the reconciliation
        if first:
            ms.cell(ms.max_row, 4).fill = green
    # sheet models matched through an invoice item code but not by master description
    for key, m in loaded.items():
        if key not in done:
            ms.append(["(via invoice code)", key, m["master_kg"], m["sheet_qty"], round(m["loaded_kg"], 2), key])
            ms.cell(ms.max_row, 4).fill = green
    ms.append([])
    ms.append(["Loading sheet models NOT found in the master sheet (review)"])
    ms.cell(ms.max_row, 1).font = Font(bold=True, size=12)
    ms.append(["Model", "Written on sheet", "Sheet line", "Qty (nos)", "kg/pc used", "Loaded kg", "Weight from"])
    for c in ms[ms.max_row]:
        c.font = bold
    for m in sorted(w["summary"]["sheet_not_in_master"], key=lambda m: model_sort_key(m["model"])):
        ms.append([m["model"], m["sheet_model"], m["sheet_lines"], m["sheet_qty"], m["master_kg"],
                   round(m["loaded_kg"], 2) if m["loaded_kg"] is not None else None, m["kg_source"] or "no weight"])
        ms.cell(ms.max_row, 4).fill = amber
    for col, wd in zip("ABCDEF", [18, 46, 12, 16, 12, 14]):
        ms.column_dimensions[col].width = wd
    ms.freeze_panes = "A2"


def _rs(kg, rate):
    return None if kg is None else kg * rate


def _annex_rows(w):
    """Model-wise weight rows for the company report: (model, loaded pcs, invoiced pcs, kg/pc, loaded kg,
    invoiced kg, difference kg, remark)."""
    out = []
    for m in sorted(w["models"], key=lambda m: model_sort_key(m["model"])):
        diff = None if m["loaded_kg"] is None else m["billed_kg"] - m["loaded_kg"]
        remark = [x for x in (
            m["qty_status"].lower() if m["qty_status"] not in ("OK", "NOT IN MASTER") else "",
            m["price_status"].lower() if m["price_status"] in ("PRICE HIGH", "PRICE LOW") else "",
            ("not in master, no weight" if m["master_kg"] is None else "not in master")
            if m["kg_source"] != "master" else "") if x]
        out.append((m["model"], m["sheet_qty"], m["invoice_qty"], m["master_kg"], m["loaded_kg"], m["billed_kg"],
                    diff, ", ".join(remark)))
    return out


def _report_info(w, invoices):
    s, rate = w["summary"], w["summary"]["rate_per_kg"]
    inv_txt = ", ".join(f"{i.get('invoice_no') or '?'}" + (f" dated {i['date']}" if i.get("date") else "") for i in invoices)
    recipient = next((i.get("recipient") for i in invoices if i.get("recipient")), "")
    line = (f"Loading sheet vs tax invoice{'s' if len(invoices) > 1 else ''} {inv_txt}"
            + (f" · {recipient}" if recipient else "") + f" · Rate Rs {rate:,.2f} per kg")
    var = s["variance_kg"]
    pcs = s["billed_pcs"] - s["loaded_pcs"]
    head = [["", "Pieces", "Weight (kg)", "Value (Rs)"],
            ["Loaded, as per loading sheet", s["loaded_pcs"], s["loaded_kg"], s["loaded_kg"] * rate],
            ["Invoiced", s["billed_pcs"], s["billed_kg"], s["billed_kg"] * rate],
            ["Difference (invoiced - loaded)", pcs, var, var * rate]]
    verdict = (f"The invoice bills {abs(var):,.0f} kg (Rs {abs(var) * rate:,.0f}) {'MORE' if var > 0 else 'LESS'} "
               "than the weight loaded." if var else "The invoiced weight equals the weight loaded.")
    return line, head, verdict


def company_story(w, invoices, ss, colors, mm, Paragraph, Spacer, Table, TableStyle, KeepTogether=None, PageBreak=None):
    """The weight report for the company: load details, loaded vs invoiced (pcs, kg, Rs) and the model-wise
    annexure. invoices: [{"invoice_no", "date", "recipient"}]."""
    from reportlab.lib.styles import ParagraphStyle

    s = w["summary"]
    body = ParagraphStyle("body", parent=ss["Normal"], fontSize=9, leading=12)
    big = ParagraphStyle("big", parent=ss["Normal"], fontSize=13, leading=17, spaceBefore=4, spaceAfter=2)
    h2 = ParagraphStyle("h2", parent=ss["Heading2"], fontSize=12, spaceBefore=8, spaceAfter=4)
    kg = lambda v: "" if v is None else f"{v:,.0f}"
    skg = lambda v: "" if v is None else ("+" if round(v) > 0 else "") + f"{v:,.0f}"
    k2 = lambda v: "" if v is None else f"{v:,.2f}"
    head_bg, grid = colors.HexColor("#E7E9EE"), colors.HexColor("#9AA1AD")
    red, green, amber = colors.HexColor("#FFC7CE"), colors.HexColor("#C6EFCE"), colors.HexColor("#FFF2CC")

    def table(data, widths, extra=()):
        t = Table(data, repeatRows=1, colWidths=[x * mm for x in widths], hAlign="LEFT")
        t.setStyle(TableStyle([("FONTSIZE", (0, 0), (-1, -1), 8), ("GRID", (0, 0), (-1, -1), 0.3, grid),
                               ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("BACKGROUND", (0, 0), (-1, 0), head_bg),
                               ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
                               ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
                               ("BACKGROUND", (0, -1), (-1, -1), head_bg)] + list(extra)))
        return t

    line, head, verdict = _report_info(w, invoices)
    var = s["variance_kg"]
    story = [Paragraph("Weight Reconciliation Report", ss["Title"]), Paragraph(line.replace("·", "&middot;"), body),
             Spacer(0, 4 * mm)]
    data = [head[0]] + [[r[0], f"{r[1]:+d}" if n == 3 and r[1] else r[1], skg(r[2]) if n == 3 else kg(r[2]),
                         skg(r[3]) if n == 3 else kg(r[3])] for n, r in enumerate(head[1:], 1)]
    story += [table(data, [70, 25, 35, 35], extra=[("BACKGROUND", (2, -1), (3, -1), red if var else green),
                                                   ("FONTSIZE", (0, 0), (-1, -1), 9.5)]),
              Spacer(0, 3 * mm), Paragraph(f"<b>{verdict}</b>", big)]

    story.append(Paragraph("Annexure: model-wise weight", h2))
    data = [["Model", "Loaded pcs", "Invoiced pcs", "kg / pc", "Loaded kg", "Invoiced kg", "Difference kg", "Remark"]]
    style = []
    for n, r in enumerate(_annex_rows(w), 1):
        data.append([r[0], r[1], r[2], k2(r[3]), kg(r[4]), kg(r[5]), skg(r[6]), r[7]])
        if r[7]:
            style.append(("BACKGROUND", (7, n), (7, n), amber if r[7].startswith("not in master") else red))
    data.append(["Total", s["loaded_pcs"], s["billed_pcs"], "", kg(s["loaded_kg"]), kg(s["billed_kg"]), skg(var), ""])
    story.append(table(data, [27, 18, 21, 15, 19, 19, 21, 46],
                       extra=style + [("ALIGN", (7, 1), (7, -1), "LEFT"), ("FONTSIZE", (7, 1), (7, -1), 7)]))
    return story


def write_company_xlsx(path, w, invoices):
    """The company report as Excel: same info block and model-wise annexure as the PDF."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    s = w["summary"]
    line, head, verdict = _report_info(w, invoices)
    wb = Workbook()
    ws = wb.active
    ws.title = "Weight report"
    bold = Font(bold=True)
    thin = Side(style="thin", color="9AA1AD")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr = PatternFill("solid", fgColor="E7E9EE")
    red, green, amber = (PatternFill("solid", fgColor=c) for c in ("FFC7CE", "C6EFCE", "FFF2CC"))
    kgf, signed = "#,##0", "+#,##0;-#,##0;0"

    ws.append(["Weight Reconciliation Report"])
    ws["A1"].font = Font(bold=True, size=16)
    ws.append([line])
    ws.append([])

    def row(values, fmts=(), fill_cols=(), fill=None, header=False):
        ws.append(values)
        r = ws.max_row
        for c, v in enumerate(values, 1):
            cell = ws.cell(r, c)
            cell.border = box
            if header:
                cell.font, cell.fill = bold, hdr
            if c - 1 < len(fmts) and fmts[c - 1]:
                cell.number_format = fmts[c - 1]
        for c in fill_cols:
            ws.cell(r, c).fill = fill
        return r

    row(head[0], header=True)
    row(head[1], (None, "0", kgf, kgf))
    row(head[2], (None, "0", kgf, kgf))
    r = row([head[3][0], head[3][1], round(head[3][2]), round(head[3][3])], (None, "+0;-0;0", signed, signed),
            fill_cols=(3, 4), fill=red if s["variance_kg"] else green)
    for c in range(1, 5):
        ws.cell(r, c).font = bold
    ws.append([])
    ws.append([verdict])
    ws.cell(ws.max_row, 1).font = Font(bold=True, size=13)
    ws.append([])
    ws.append(["Annexure: model-wise weight"])
    ws.cell(ws.max_row, 1).font = Font(bold=True, size=12)
    row(["Model", "Loaded pcs", "Invoiced pcs", "kg / pc", "Loaded kg", "Invoiced kg", "Difference kg", "Remark"],
        header=True)
    first = ws.max_row + 1
    fmts = (None, "0", "0", "0.00", kgf, kgf, signed, None)
    for a in _annex_rows(w):
        r = row([a[0], a[1], a[2], None if a[3] is None else round(a[3], 2),
                 None if a[4] is None else round(a[4]), round(a[5]), None if a[6] is None else round(a[6]), a[7]], fmts)
        if a[7]:
            ws.cell(r, 8).fill = amber if a[7].startswith("not in master") else red
    r = row(["Total", s["loaded_pcs"], s["billed_pcs"], None, s["loaded_kg"], s["billed_kg"], s["variance_kg"], None],
            fmts)
    for c in range(1, 9):
        ws.cell(r, c).font, ws.cell(r, c).fill = bold, hdr
    for col, wd in zip("ABCDEFGH", [34, 12, 13, 10, 12, 12, 14, 40]):
        ws.column_dimensions[col].width = wd
    ws.freeze_panes = ws.cell(first, 1)
    for rr in ws.iter_rows(min_row=first, max_row=ws.max_row):
        rr[0].alignment = Alignment(horizontal="left")
    wb.save(path)


def write_company_pdf(path, w, invoices):
    """Stand-alone weight gap report to send to the company (A4 portrait)."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    doc = SimpleDocTemplate(path, pagesize=A4, leftMargin=12 * mm, rightMargin=12 * mm, topMargin=12 * mm,
                            bottomMargin=12 * mm, title="Weight Reconciliation Report", author="mewreport")

    def footer(canvas, doc_):
        canvas.setFont("Helvetica", 7)
        canvas.drawRightString(A4[0] - 12 * mm, 6 * mm, f"Page {doc_.page}")
    doc.build(company_story(w, invoices, getSampleStyleSheet(), colors, mm, Paragraph, Spacer, Table, TableStyle,
                            KeepTogether, PageBreak), onFirstPage=footer, onLaterPages=footer)
