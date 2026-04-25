"""
Reportlab-based PDF generator for ENCON quote/offer documents.

Builds a clean, self-contained PDF directly from the same `quote_data` that
feeds the Word template. Bypasses LibreOffice entirely — no Java, no font
discovery, no headless conversion crashes.

Public entry point:
    generate_quote_pdf(quote_data: dict, output_path: str) -> None

Sections produced:
    1. Header  (title, quote no, date, validity)
    2. Customer details
    3. Technical data (key-value table)
    4. Scope of Supply (gas / air / pilot / temperature-control bullets,
       sourced from quote_data["customer"]["bom_items"])
    5. Make list (Item / Make table)
    6. Commercial summary (subtotal, GST, grand total, total in words)
    7. Standard terms (delivery, payment, validity, GST/HSN)
    8. Footer with "Page X / Y"

The layout deliberately keeps the visual model simple — text flows on
A4 pages with 2 cm margins. No floating drawings, no anchored shapes.
This trades the marketing flair of the Word template for 100% reliable
rendering everywhere.
"""
from __future__ import annotations
from typing import Iterable

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm, mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, KeepTogether,
)
from reportlab.platypus.flowables import HRFlowable

NAVY    = colors.HexColor("#1F3A5F")
BAND_BG = colors.HexColor("#F2F4F8")
ROW_ALT = colors.HexColor("#FAFBFD")
BORDER  = colors.HexColor("#CBD5E1")


# ─────────────── styles ───────────────
def _styles():
    base = getSampleStyleSheet()
    s = {
        "Title":   ParagraphStyle("Title", parent=base["Title"],
                                  fontName="Helvetica-Bold", fontSize=18,
                                  textColor=NAVY, spaceAfter=4, alignment=1),
        "Sub":     ParagraphStyle("Sub", parent=base["Normal"],
                                  fontName="Helvetica", fontSize=10,
                                  textColor=colors.grey, alignment=1, spaceAfter=12),
        "H1":      ParagraphStyle("H1", parent=base["Heading1"],
                                  fontName="Helvetica-Bold", fontSize=13,
                                  textColor=NAVY, spaceBefore=14, spaceAfter=6),
        "H2":      ParagraphStyle("H2", parent=base["Heading2"],
                                  fontName="Helvetica-Bold", fontSize=11,
                                  textColor=NAVY, spaceBefore=8, spaceAfter=4),
        "H3":      ParagraphStyle("H3", parent=base["Heading3"],
                                  fontName="Helvetica-Bold", fontSize=10,
                                  textColor=colors.black, spaceBefore=4, spaceAfter=2),
        "Body":    ParagraphStyle("Body", parent=base["BodyText"],
                                  fontName="Helvetica", fontSize=10,
                                  leading=13),
        "Bullet":  ParagraphStyle("Bullet", parent=base["BodyText"],
                                  fontName="Helvetica", fontSize=10,
                                  leading=13, leftIndent=14, bulletIndent=4),
        "Small":   ParagraphStyle("Small", parent=base["BodyText"],
                                  fontName="Helvetica", fontSize=9,
                                  leading=12, textColor=colors.grey),
    }
    return s


# ─────────────── BOM categorisation ───────────────
def _split_bom_into_scope(bom_items: Iterable[dict]) -> dict:
    """Same rule as quote_writer._control_system_sections, plus an explicit
    'pilot' bucket. Returns {gas, air, pilot, temp}: list[str] each."""
    import re
    gas, air, pilot, temp = [], [], [], []
    seen = {"gas": set(), "air": set(), "pilot": set(), "temp": set()}

    def _clean(name: str) -> str:
        n = name.strip()
        n = re.sub(r"^(GAS\s+TRAIN)\s+.*$", r"\1", n, flags=re.IGNORECASE)
        return n.strip()

    def _add(bucket, key, entry):
        if entry in seen[key]:
            return
        seen[key].add(entry)
        bucket.append(entry)

    for x in bom_items or []:
        item = (x.get("item") or "").strip()
        if not item or item in ("BOUGHT OUT ITEMS", "ENCON ITEMS", "GRAND TOTAL"):
            continue
        media = (x.get("media") or "").strip().upper()
        entry = _clean(item)
        upper_item = item.upper()
        is_pilot_named = "(PILOT BURNER)" in upper_item or "(UV LINE)" in upper_item
        is_pilot_media = media.endswith(" PILOT LINE")
        if is_pilot_named or is_pilot_media:
            _add(pilot, "pilot", entry)
        elif media == "COMB AIR":
            _add(air, "air", entry)
        elif media == "MISC ITEMS":
            _add(temp, "temp", entry)
        elif media.endswith(" LINE") and media != "PURGING LINE":
            _add(gas, "gas", entry)
    return {"gas": gas, "air": air, "pilot": pilot, "temp": temp}


# ─────────────── flowable builders ───────────────
def _kv_table(rows, st, label_w_cm=5.5, total_w_cm=17.0):
    """Build a 2-column Item|Value table from a list of (label, value) tuples."""
    data = []
    for label, value in rows:
        if value is None or value == "":
            continue
        data.append([
            Paragraph(f"<b>{label}</b>", st["Body"]),
            Paragraph(str(value), st["Body"]),
        ])
    if not data:
        return None
    val_w = total_w_cm - label_w_cm
    t = Table(data, colWidths=[label_w_cm * cm, val_w * cm])
    t.setStyle(TableStyle([
        ("VALIGN",     (0, 0), (-1, -1), "TOP"),
        ("INNERGRID",  (0, 0), (-1, -1), 0.4, BORDER),
        ("BOX",        (0, 0), (-1, -1), 0.6, BORDER),
        ("BACKGROUND", (0, 0), (0, -1), BAND_BG),
        ("LEFTPADDING",  (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING",   (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 4),
    ]))
    return t


def _bullets(items, st):
    return [Paragraph(f"• {x}", st["Bullet"]) for x in items]


def _make_list_table(make_list, st, total_w_cm=17.0):
    if not make_list:
        return None
    data = [[Paragraph("<b>ITEM</b>", st["Body"]),
             Paragraph("<b>MAKE</b>", st["Body"])]]
    for x in make_list:
        data.append([
            Paragraph(x.get("item", ""), st["Body"]),
            Paragraph(x.get("make", ""), st["Body"]),
        ])
    item_w = total_w_cm * 0.7
    make_w = total_w_cm * 0.3
    t = Table(data, colWidths=[item_w * cm, make_w * cm], repeatRows=1)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("TEXTCOLOR",  (0, 0), (-1, 0), colors.white),
        ("FONTNAME",   (0, 0), (-1, 0), "Helvetica-Bold"),
        ("INNERGRID",  (0, 0), (-1, -1), 0.4, BORDER),
        ("BOX",        (0, 0), (-1, -1), 0.6, BORDER),
        ("VALIGN",     (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING",  (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING",   (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 3),
    ]
    # Alternate row backgrounds (skipping the header)
    for i in range(2, len(data), 2):
        style.append(("BACKGROUND", (0, i), (-1, i), ROW_ALT))
    t.setStyle(TableStyle(style))
    return t


# ─────────────── page footer (Page X / Y) ───────────────
def _draw_footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica-Bold", 9)
    canvas.setFillColor(NAVY)
    text = f"Page {doc.page} / {{TOTAL_PAGES}}"
    canvas.drawRightString(A4[0] - 2 * cm, 1.2 * cm, text)
    canvas.restoreState()


def _post_process_total_pages(pdf_path: str):
    """The footer hooks emit 'Page N / {TOTAL_PAGES}' as a literal text string.
    After the doc is built we know the real page count, so rewrite the file
    to substitute the placeholder. Cheap, single-pass byte replace."""
    with open(pdf_path, "rb") as f:
        data = f.read()
    placeholder = b"{TOTAL_PAGES}"
    if placeholder not in data:
        return
    # Count actual pages by counting "/Type /Page" objects
    page_count = data.count(b"/Type /Page\n") or data.count(b"/Type /Page ") or 1
    # Pad replacement to keep byte length stable so xref offsets stay valid.
    repl = str(page_count).encode("ascii")
    # Pad with leading spaces to match placeholder length
    repl_padded = repl.rjust(len(placeholder))
    new_data = data.replace(placeholder, repl_padded)
    with open(pdf_path, "wb") as f:
        f.write(new_data)


# ─────────────── main entry ───────────────
def generate_quote_pdf(quote_data: dict, output_path: str) -> None:
    """Build the PDF and save it to output_path."""
    customer = quote_data.get("customer", {}) or {}
    items    = quote_data.get("items", []) or []
    grand    = float(quote_data.get("grand_total") or quote_data.get("subtotal") or 0)

    st = _styles()
    doc = SimpleDocTemplate(
        output_path, pagesize=A4,
        leftMargin=2 * cm, rightMargin=2 * cm,
        topMargin=2 * cm, bottomMargin=2 * cm,
        title=f"Quote {quote_data.get('quote_no', '')}",
        author="ENCON Combustion Pvt Ltd",
    )

    flow = []

    # ── 1. Header ────────────────────────────────────────────────────
    flow.append(Paragraph("ENCON COMBUSTION PVT LTD", st["Title"]))
    flow.append(Paragraph(
        f"Offer for {customer.get('subject') or customer.get('project_name') or 'Preheating System'}",
        st["Sub"]))
    flow.append(HRFlowable(width="100%", thickness=1, color=NAVY, spaceAfter=10))

    # ── 2. Quote / customer summary ──────────────────────────────────
    summary_rows = [
        ("Quote No.",   quote_data.get("quote_no", "")),
        ("Date",        quote_data.get("date", "")),
        ("Validity",    f"{quote_data.get('valid_days', 30)} days"),
        ("Company",     customer.get("company_name", "")),
        ("City / State", ", ".join([x for x in (customer.get("company_city", ""),
                                                  customer.get("company_state", "")) if x])),
        ("GSTIN",       customer.get("gstin", "")),
        ("Contact",     " — ".join([x for x in (customer.get("poc_name", ""),
                                                  customer.get("poc_designation", "")) if x])),
        ("Mobile",      customer.get("mobile_no", "")),
        ("Email",       customer.get("email", "")),
    ]
    t = _kv_table(summary_rows, st)
    if t: flow.append(t)
    flow.append(Spacer(1, 0.4 * cm))

    # ── 3. Technical data ────────────────────────────────────────────
    flow.append(Paragraph("Technical Data", st["H1"]))
    tech_rows = [
        ("Ladle Type",         _ladle_type(items, customer)),
        ("Ladle Capacity",     _fmt_with_unit(customer.get("ladle_tons"), "Ton")),
        ("Ladle Dimensions",   customer.get("ladle_dim", "")),
        ("Reference TS",       customer.get("ladle_drawing_no", "")),
        ("Refractory Weight",  _fmt_with_unit(customer.get("refractory_weight_kg"), "Kg")),
        ("Heating Schedule",   customer.get("heating_schedule", "")),
        ("Heating Time",       customer.get("heating_time", "")),
        ("Fuel",               customer.get("fuel_name", "")),
        ("Calorific Value",    customer.get("fuel_cv", "")),
        ("Firing Rate",        customer.get("fuel_consumption", "")),
        ("Burner",             customer.get("burner_model", "")),
        ("Burner Capacity",    customer.get("burner_capacity_range", "")),
        ("Combustion Blower",  customer.get("blower_model", "")),
        ("Blower Capacity",    customer.get("blower_capacity", "")),
        ("Pumping Unit",       customer.get("pumping_unit", "")),
        ("Hood Movement",      _hood_label(customer.get("hood_type"), customer.get("hood_movement"))),
        ("Pilot Burner Fuel",  customer.get("pilot_gas_type", "")),
        ("Ignition Method",    customer.get("ignition_method", "")),
        ("Power Pack",         customer.get("hydraulic_motor_hp", "")),
        ("Max Electrical Load", customer.get("max_electrical_load", "")),
    ]
    t = _kv_table(tech_rows, st)
    if t: flow.append(t)

    # ── 4. Scope of Supply ───────────────────────────────────────────
    flow.append(Paragraph("Scope of Supply", st["H1"]))
    scope = _split_bom_into_scope(customer.get("bom_items"))
    flow.append(Paragraph("Burner Control System", st["H2"]))

    flow.append(Paragraph("On the main gas pipeline:", st["H3"]))
    flow.extend(_bullets(scope["gas"] or ["(none)"], st))

    flow.append(Paragraph("On the main air pipeline:", st["H3"]))
    flow.extend(_bullets(scope["air"] or ["(none)"], st))

    if scope["pilot"]:
        flow.append(Paragraph("Pilot:", st["H3"]))
        flow.extend(_bullets(scope["pilot"], st))

    flow.append(Paragraph("Temperature Control System", st["H2"]))
    flow.extend(_bullets(scope["temp"] or ["(none)"], st))

    # ── 5. Make List ─────────────────────────────────────────────────
    flow.append(Spacer(1, 0.4 * cm))
    flow.append(Paragraph("Make List", st["H1"]))
    mlt = _make_list_table(quote_data.get("make_list") or
                            customer.get("bom_items") or [], st)
    if mlt:
        flow.append(mlt)

    # ── 6. Commercial summary ────────────────────────────────────────
    flow.append(Spacer(1, 0.4 * cm))
    flow.append(Paragraph("Commercial Summary", st["H1"]))
    money_rows = [
        ("Subtotal",        f"Rs. {float(quote_data.get('subtotal', grand)):,.2f}"),
        ("GST (Extra)",     "18% on basic value"),
        ("Grand Total (excl. GST)", f"Rs. {grand:,.2f}"),
    ]
    if customer.get("total_in_words"):
        money_rows.append(("In Words", customer["total_in_words"]))
    t = _kv_table(money_rows, st)
    if t: flow.append(t)

    # ── 7. Standard terms ────────────────────────────────────────────
    flow.append(Spacer(1, 0.4 * cm))
    flow.append(Paragraph("Terms & Conditions", st["H1"]))
    terms = [
        ("Prices",     "Ex Works Bhagola, Dist: Palwal, Haryana — unpacked"),
        ("Delivery",   "10–12 weeks from receipt of advance & approved drawing, whichever is later"),
        ("GST",        "18 % Extra"),
        ("HSN Code",   "84541000"),
        ("PAN / GSTIN","AAACE0327M / 06AAACE0327M1ZV"),
        ("Payment",    "30 % advance with PO; 70 % against proforma invoice prior to dispatch"),
        ("Validity",   f"{quote_data.get('valid_days', 30)} days"),
    ]
    t = _kv_table(terms, st, label_w_cm=4.5)
    if t: flow.append(t)

    # Build with footer page numbers
    doc.build(flow, onFirstPage=_draw_footer, onLaterPages=_draw_footer)
    _post_process_total_pages(output_path)


# ─────────────── small helpers ───────────────
def _fmt_with_unit(val, unit):
    if val is None or val == "":
        return ""
    s = str(val).strip()
    return s if unit in s else f"{s} {unit}"


def _ladle_type(items, customer):
    pt = ((items[0].get("product_type") or "") if items else "").lower()
    tons = customer.get("ladle_tons")
    if "vertical" in pt:
        return f"Vertical Ladle Preheater – {tons} Ton" if tons else "Vertical Ladle Preheater"
    if "horizontal" in pt:
        return f"Horizontal Ladle Preheater – {tons} Ton" if tons else "Horizontal Ladle Preheater"
    if "tundish" in pt:
        return f"Tundish Preheater – {tons} Ton" if tons else "Tundish Preheater"
    return f"Ladle Preheater – {tons} Ton" if tons else "Ladle Preheater"


def _hood_label(hood_type, hood_movement):
    if hood_movement and hood_movement.strip():
        return hood_movement
    return {
        "up_down":       "Up and Down (hydraulic)",
        "swivel_manual": "Swivelling — Manual",
        "swivel_geared": "Swivelling — Geared",
        "swivel":        "Swivelling",
    }.get((hood_type or "").lower(), hood_type or "")
