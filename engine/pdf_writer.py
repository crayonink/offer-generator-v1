"""
Reportlab-based PDF generator for ENCON quote/offer documents.

Builds a clean, self-contained PDF directly from the same `quote_data` that
feeds the Word template. Bypasses LibreOffice entirely.

Public entry point:
    generate_quote_pdf(quote_data: dict, output_path: str) -> None

The Scope of Supply is structured to match ENCON's standard offer format:
descriptive prose for the major equipment areas, single-column component
tables for each pipeline (combustion air, main fuel gas train, pilot line,
purging line), a numbered Temperature Control System list, and closing
sections for Operational Sequence, Painting, Cabling and Pipeline scope.
"""
from __future__ import annotations
from typing import Iterable

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    KeepTogether,
)
from reportlab.platypus.flowables import HRFlowable

NAVY    = colors.HexColor("#1F3A5F")
BAND_BG = colors.HexColor("#F2F4F8")
ROW_ALT = colors.HexColor("#FAFBFD")
BORDER  = colors.HexColor("#CBD5E1")


# ─────────────── styles ───────────────
def _styles():
    base = getSampleStyleSheet()
    return {
        "Title":  ParagraphStyle("Title", parent=base["Title"],
                                 fontName="Helvetica-Bold", fontSize=18,
                                 textColor=NAVY, spaceAfter=4, alignment=1),
        "Sub":    ParagraphStyle("Sub", parent=base["Normal"],
                                 fontName="Helvetica", fontSize=10,
                                 textColor=colors.grey, alignment=1, spaceAfter=12),
        "H1":     ParagraphStyle("H1", parent=base["Heading1"],
                                 fontName="Helvetica-Bold", fontSize=13,
                                 textColor=NAVY, spaceBefore=14, spaceAfter=6),
        "H2":     ParagraphStyle("H2", parent=base["Heading2"],
                                 fontName="Helvetica-Bold", fontSize=11,
                                 textColor=NAVY, spaceBefore=10, spaceAfter=4,
                                 underlineWidth=0.5),
        "H3":     ParagraphStyle("H3", parent=base["Heading3"],
                                 fontName="Helvetica-Bold", fontSize=10,
                                 textColor=colors.black, spaceBefore=4, spaceAfter=2),
        "Body":   ParagraphStyle("Body", parent=base["BodyText"],
                                 fontName="Helvetica", fontSize=10, leading=13,
                                 spaceAfter=4, alignment=4),  # justified
        "Numb":   ParagraphStyle("Numb", parent=base["BodyText"],
                                 fontName="Helvetica", fontSize=10, leading=13,
                                 leftIndent=18, bulletIndent=4),
    }


# ─────────────── BOM categorisation ───────────────
def _split_bom(bom_items: Iterable[dict]) -> dict:
    """Group BOM rows by their offer-doc bucket. Returns lists keyed by:
        air            - combustion-air line + pilot/UV burner line accessories
        gas_main       - main-fuel gas-train rows (NG, BFG, COG, MIXED GAS, LPG)
        gas_main_label - label for the main-fuel section ("MIX GAS TRAIN", etc.)
        pilot          - dedicated pilot gas-line rows (LPG / NG PILOT LINE)
        pilot_label    - "LPG LINE FOR PILOT BURNER" (matches pilot fuel)
        purging        - nitrogen-purging-line rows
        temp           - MISC ITEMS (temperature/control instrumentation)
    """
    import re
    air, gas_main, pilot, purging, temp = [], [], [], [], []
    seen = {k: set() for k in ("air", "gas", "pilot", "purge", "temp")}
    main_media = None
    pilot_media = None

    def _clean(name: str) -> str:
        n = (name or "").strip()
        # 'GAS TRAIN 500 NM3/Hr ...' -> 'GAS TRAIN'
        n = re.sub(r"^(GAS\s+TRAIN)\b.*$", r"\1", n, flags=re.IGNORECASE)
        # Strip trailing ' - <size/spec>' (e.g., '- 250 NB F150#',
        # '- 80 NB', '- RANGE- 0-1600 mBAR'). Em/en dash too.
        n = re.sub(r"\s+[\-–—]\s+.*$", "", n)
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
        # Items whose name marks them as pilot/UV accessories belong with
        # the pilot equipment, never the main combustion air list -- even
        # when their MEDIA tag is "COMB AIR".
        is_pilot_named = "(PILOT BURNER)" in upper_item or "(UV LINE)" in upper_item
        if media.endswith(" PILOT LINE") or is_pilot_named:
            if media.endswith(" PILOT LINE"):
                pilot_media = pilot_media or media
            _add(pilot, "pilot", entry)
        elif media == "COMB AIR":
            _add(air, "air", entry)
        elif media == "PURGING LINE":
            _add(purging, "purge", entry)
        elif media == "MISC ITEMS":
            _add(temp, "temp", entry)
        elif media.endswith(" LINE"):
            main_media = main_media or media
            _add(gas_main, "gas", entry)

    # Friendly labels for the gas/oil-train heading. Both short codes
    # ("BG LINE", "MG LINE") and full names ("FURNACE OIL LINE") accepted.
    main_label_map = {
        # gas
        "MIXED GAS LINE":  "MIX GAS TRAIN FOR MAIN BURNER",
        "MIX GAS LINE":    "MIX GAS TRAIN FOR MAIN BURNER",
        "MG LINE":         "MIX GAS TRAIN FOR MAIN BURNER",
        "NG LINE":         "NATURAL GAS TRAIN FOR MAIN BURNER",
        "LPG LINE":        "LPG GAS TRAIN FOR MAIN BURNER",
        "RLNG LINE":       "RLNG GAS TRAIN FOR MAIN BURNER",
        "BFG LINE":        "BFG TRAIN FOR MAIN BURNER",
        "BG LINE":         "BFG TRAIN FOR MAIN BURNER",
        "COG LINE":        "COG TRAIN FOR MAIN BURNER",
        # oil
        "LDO LINE":        "LDO LINE FOR MAIN BURNER",
        "FO LINE":         "FURNACE OIL LINE FOR MAIN BURNER",
        "FURNACE OIL LINE":"FURNACE OIL LINE FOR MAIN BURNER",
        "LSHS LINE":       "LSHS LINE FOR MAIN BURNER",
        "HSD LINE":        "HSD LINE FOR MAIN BURNER",
        "SKO LINE":        "SKO LINE FOR MAIN BURNER",
        "HDO LINE":        "HDO LINE FOR MAIN BURNER",
        "CFO LINE":        "CFO LINE FOR MAIN BURNER",
        "OIL LINE":        "OIL LINE FOR MAIN BURNER",
    }
    main_label = main_label_map.get(main_media or "", "MAIN FUEL TRAIN")

    # Intro paragraph wording — three categories:
    #   NG / LPG / RLNG          -> packaged MADAS gas-train assembly
    #   COG / Mixed Gas / BFG    -> field-assembled discrete components
    #   LDO / FO / LSHS / HSD ...-> oil pipeline from pumping unit
    PACKAGED  = {"NG LINE", "LPG LINE", "RLNG LINE"}
    DISCRETE  = {"COG LINE", "MIXED GAS LINE", "MIX GAS LINE", "MG LINE",
                 "BFG LINE", "BG LINE"}
    OIL       = {"LDO LINE", "FO LINE", "FURNACE OIL LINE", "LSHS LINE",
                 "HSD LINE", "SKO LINE", "HDO LINE", "CFO LINE", "OIL LINE"}
    if main_media in PACKAGED:
        main_intro = ("We will supply a packaged MADAS gas train for firing of "
                      "the Burner, consisting of the following:")
    elif main_media in DISCRETE:
        main_intro = ("Gas train will be field-assembled from the following "
                      "discrete components:")
    elif main_media in OIL:
        main_intro = ("Oil pipeline from the Heating & Pumping Unit to the "
                      "Burner will consist of the following components:")
    else:
        main_intro = ("Gas train will be supplied for firing of Burner, "
                      "consisting of the following components:")

    pilot_fuel = (pilot_media or "").replace(" PILOT LINE", "").strip()
    pilot_label = f"{pilot_fuel} LINE FOR PILOT BURNER" if pilot_fuel else "PILOT LINE"

    return {
        "air": air,
        "gas_main": gas_main, "gas_main_label": main_label, "gas_main_intro": main_intro,
        "pilot": pilot,       "pilot_label": pilot_label,
        "purging": purging,
        "temp": temp,
    }


# ─────────────── small helpers ───────────────
def _kv_table(rows, st, label_w_cm=5.5, total_w_cm=17.0):
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
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("INNERGRID", (0, 0), (-1, -1), 0.4, BORDER),
        ("BOX", (0, 0), (-1, -1), 0.6, BORDER),
        ("BACKGROUND", (0, 0), (0, -1), BAND_BG),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    return t


def _component_table(items, total_w_cm=10.0):
    """Single-column bordered table — matches the reference offer's Combustion
    Air Line / Gas Train / Pilot / Purging tables."""
    if not items:
        return None
    data = [[x] for x in items]
    t = Table(data, colWidths=[total_w_cm * cm])
    t.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOX", (0, 0), (-1, -1), 0.6, colors.black),
        ("INNERGRID", (0, 0), (-1, -1), 0.4, colors.black),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    return t


def _make_list_table(make_list, st, total_w_cm=17.0):
    if not make_list:
        return None
    data = [[Paragraph("<b>S. No.</b>", st["Body"]),
             Paragraph("<b>ITEM NAME</b>", st["Body"]),
             Paragraph("<b>MAKE</b>", st["Body"])]]
    for i, x in enumerate(make_list, start=1):
        data.append([
            Paragraph(str(i), st["Body"]),
            Paragraph(x.get("item", ""), st["Body"]),
            Paragraph(x.get("make", ""), st["Body"]),
        ])
    sno_w  = total_w_cm * 0.10
    item_w = total_w_cm * 0.55
    make_w = total_w_cm * 0.35
    t = Table(data, colWidths=[sno_w * cm, item_w * cm, make_w * cm], repeatRows=1)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ALIGN", (0, 0), (0, -1), "CENTER"),     # S.No column centered
        ("ALIGN", (-1, 0), (-1, -1), "LEFT"),      # MAKE column left
        ("INNERGRID", (0, 0), (-1, -1), 0.4, BORDER),
        ("BOX", (0, 0), (-1, -1), 0.6, BORDER),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]
    for i in range(2, len(data), 2):
        style.append(("BACKGROUND", (0, i), (-1, i), ROW_ALT))
    t.setStyle(TableStyle(style))
    return t


# ─────────────── prose blocks (Scope of Supply) ───────────────
def _hood_is_swivel(hood_type: str) -> bool:
    return (hood_type or "").lower() in ("swivel", "swivel_manual", "swivel_geared")


def _product_kind(items):
    """Return 'vlph' | 'hlph' | 'tundish' from the first item's product_type."""
    pt = ((items[0].get("product_type") or "") if items else "").lower()
    if "horizontal" in pt: return "hlph"
    if "tundish"    in pt: return "tundish"
    return "vlph"


def _prose_blocks(customer: dict, scope: dict, control_mode: str | None,
                  has_purging: bool, product_kind: str = "vlph"):
    """Return ordered list of (heading, body_text) tuples. Body can be a single
    string or a list of strings (multiple paragraphs).
    """
    hood_type     = (customer.get("hood_type") or "").lower()
    is_swivel     = _hood_is_swivel(hood_type)
    is_swivel_geared = hood_type == "swivel_geared"
    is_hlph       = product_kind == "hlph"
    fuel_name     = (customer.get("fuel_name") or "").strip() or "Mixed Gas"
    pilot_fuel    = (customer.get("pilot_gas_type") or "LPG").upper()
    is_oil        = bool(customer.get("is_oil"))
    auto_ignition = bool(customer.get("special_auto_ignition"))

    # 1. Steel structure
    yield "STEEL STRUCTURE", (
        "Fabricated steel structure will consist of mild steel plates and rolled "
        "steel sections required to hold and support the ladle hood which will "
        "be attached to the steel fabricated, sturdy and rigid column."
    )

    # 2. Ladle hood / dish
    if is_swivel:
        yield "LADLE DISH", (
            "The Ladle Dish fabricated out of mild steel plates of suitable "
            "thickness will be held with suitable supports. Combustion Burner "
            "will be installed on the ladle dish."
        )
        yield "DISH LINING", (
            "Ceramic fiber modular lining of suitable thickness will be provided "
            "in the hood. The lining will be held with T-anchors made of SS-304 "
            "specially designed for increase the life of fiber lining & its "
            "safety. The thickness of fiber lining on preheater's hood shall "
            "be 200 mm."
        )
    else:
        yield "LADLE HOOD WITH ARM", (
            "The Ladle hood will be fabricated out of mild steel plates of "
            "suitable thickness and will be held through suitable support. "
            "Burner will be installed on the hood. We are providing burner "
            "block made of high-temperature-resistant metal casting for "
            "longer and trouble-free life."
        )
        yield "HOOD LINING", (
            "Ceramic fiber modular lining of suitable thickness will be "
            "provided in the hood. The lining will be held with T-anchors "
            "made of SS-304 specially designed for increase the life of "
            "fiber lining & its safety."
        )

    # 3. Hood movement mechanism — varies by product (HLPH vs VLPH)
    #    and by VLPH hood selection (up_down / swivel_manual / swivel_geared).
    if is_hlph:
        yield "HOOD MOVEMENT MECHANISM", (
            "Movement of the Ladle hood shall be done through trolley drive & "
            "geared motor mechanism. Our scope of supply shall include trolley "
            "carrying all the above equipment, and burner firing hood driven by "
            "Electro-mechanical drive (Geared Motor). The trolley will be "
            "supported on 4 nos. of wheel bearing assembly. The maximum forward "
            "and reverse of the trolley will be governed through Limit Switches. "
            "The rail for trolley is in CLIENT Scope."
        )
    elif is_swivel_geared:
        yield "HOOD MOVEMENT MECHANISM", (
            "The hood will be raised and swivelled away from the ladle through "
            "an electro-mechanical (geared motor) drive mounted on the structural "
            "column. The maximum swivel of the hood will be governed through "
            "Limit Switches. A mechanical locking arrangement will be provided "
            "to prevent the hood from accidental movement during preheating."
        )
    elif is_swivel:
        yield "HOOD MOVEMENT MECHANISM", (
            "The hood will be raised and swivelled away from the ladle through "
            "a manually operated swivelling mechanism mounted on the structural "
            "column. Mechanical locking arrangement will be provided to prevent "
            "the hood from accidental movement during preheating."
        )
    else:
        # VLPH up_down (hydraulic)
        yield "HOOD MOVEMENT MECHANISM", (
            "The Lifting–lowering mechanism of the hood will be done by Hydraulic "
            "Cylinder & Power-Pack. The maximum lifting and lowering of the Hood "
            "will be governed through Limit Switches. An additional mechanical "
            "locking arrangement will be provided to prevent the hood from "
            "dropping down, in case of failure of the hydraulic system. The "
            "hydraulic system will consist of an oil tank, oil pump with motor, "
            "suction strainer, pressure gauges with thumb operated needle valve, "
            "Direction valves, pilot operated check valves, pressure hoses & "
            "level switch, return line filter, all filters with clogging switch "
            "and indicator, pressure relieving valve, pressure switch and "
            "temperature gauge."
        )
        yield "HYDRAULIC SYSTEM", (
            "To take out the Ladle easily after drying/heating, the system will "
            "be provided with \"Up and Down\" movement of hood. The lifting/"
            "lowering of the Ladle hood will be provided by a single acting "
            "(in push) hydraulic cylinder fitted in column with dedicated power "
            "pack. Limit switches LS1 & LS2 will be provided to specify their "
            "location. LS1 will indicate its parking position & LS2 will "
            "indicate its working position. The maximum lift of the arm should "
            "be 85° to its parking position. One no. of dedicated power pack "
            "of suitable capacity shall be provided for operating the hydraulic "
            "cylinder."
        )

    # 4. Burner
    yield "BURNER", (
        f"One Number ENCON {fuel_name} burner would be provided with the "
        "system. The burner would require compressed air for atomization, and "
        "the flame will be short and intense so that the Ladle refractory can "
        "be heated properly without damaging the refractory. The burner will "
        "be ignited automatically through ignition electrode in the main "
        "burner. The flame will be sensed using an Ionization rod present "
        "in the burner."
    )

    # 5. Pilot burner — only when Auto Ignition is requested
    if auto_ignition:
        yield "PILOT BURNER", (
            f"We shall be supplying one no. automatically {pilot_fuel} fired pilot "
            "for ignition of main burner. The pilot burner will be equipped with "
            "electrode assembly, flame sensor, ignition transformer and burner "
            "control unit."
        )

    # 6. Combustion air blower
    yield "COMBUSTION AIR BLOWER", (
        "To supply combustion and atomizing air to the above burner, we will "
        "be supplying a centrifugal steel plate air blower equipped with "
        "squirrel-cage induction motor of reputed make such as ALSTOM, ABB, "
        "or Crompton, etc."
    )


def _temp_control_items(temp_bucket: list[str], control_mode: str | None,
                         auto_control_type: str | None):
    """Convert the MISC ITEMS bucket into a numbered list matching the
    reference offer's "Temperature Control System" section. We keep it close
    to the BOM but use friendlier display names."""
    items = list(temp_bucket)
    if not items:
        # fallback to a sane default per control mode
        items = [
            "PLC with HMI",
            "Thermocouple with temperature transmitter",
        ]
        if (control_mode or "automatic") == "automatic" and (auto_control_type or "plc") == "plc":
            items += [
                "Orifice plate fitted with mass flow transmitter on gas line",
                "Orifice plate fitted with differential pressure transmitter on air line",
            ]
    return items


# ─────────────── footer (Page X / Y) ───────────────
def _draw_footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica-Bold", 9)
    canvas.setFillColor(NAVY)
    text = f"Page {doc.page} / {{TOTAL_PAGES}}"
    canvas.drawRightString(A4[0] - 2 * cm, 1.2 * cm, text)
    canvas.restoreState()


def _post_process_total_pages(pdf_path: str):
    with open(pdf_path, "rb") as f:
        data = f.read()
    placeholder = b"{TOTAL_PAGES}"
    if placeholder not in data:
        return
    page_count = data.count(b"/Type /Page\n") or data.count(b"/Type /Page ") or 1
    repl = str(page_count).encode("ascii").rjust(len(placeholder))
    new_data = data.replace(placeholder, repl)
    with open(pdf_path, "wb") as f:
        f.write(new_data)


# ─────────────── main entry ───────────────
def generate_quote_pdf(quote_data: dict, output_path: str) -> None:
    customer = quote_data.get("customer", {}) or {}
    items    = quote_data.get("items", []) or []
    grand    = float(quote_data.get("grand_total") or quote_data.get("subtotal") or 0)
    control_mode      = customer.get("control_mode") or "automatic"
    auto_control_type = customer.get("auto_control_type") or "plc"

    st = _styles()
    doc = SimpleDocTemplate(
        output_path, pagesize=A4,
        leftMargin=2 * cm, rightMargin=2 * cm,
        topMargin=2 * cm, bottomMargin=2 * cm,
        title=f"Quote {quote_data.get('quote_no', '')}",
        author="ENCON Combustion Pvt Ltd",
    )

    flow = []
    scope = _split_bom(customer.get("bom_items"))

    # ── Header ────────────────────────────────────────────────────────
    flow.append(Paragraph("ENCON COMBUSTION PVT LTD", st["Title"]))
    flow.append(Paragraph(
        f"Offer for {customer.get('subject') or customer.get('project_name') or 'Preheating System'}",
        st["Sub"]))
    flow.append(HRFlowable(width="100%", thickness=1, color=NAVY, spaceAfter=10))

    # ── Quote summary ────────────────────────────────────────────────
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

    # ── Scope of Supply ──────────────────────────────────────────────
    flow.append(Paragraph("Scope of Supply", st["H1"]))
    has_purging = bool(scope["purging"])

    # 1. Static prose blocks
    product_kind = _product_kind(items)
    for heading, body in _prose_blocks(customer, scope, control_mode, has_purging, product_kind):
        flow.append(Paragraph(heading, st["H2"]))
        if isinstance(body, (list, tuple)):
            for para in body:
                flow.append(Paragraph(para, st["Body"]))
        else:
            flow.append(Paragraph(body, st["Body"]))

    # 2. Combustion Air Line table
    flow.append(Paragraph("COMBUSTION AIR LINE", st["H2"]))
    flow.append(Paragraph("The airline will consist of the following Items:", st["Body"]))
    t = _component_table(scope["air"])
    if t: flow.append(t)
    flow.append(Spacer(1, 0.2 * cm))

    # 3. Main fuel gas train
    if scope["gas_main"]:
        flow.append(Paragraph(scope["gas_main_label"], st["H2"]))
        flow.append(Paragraph(scope["gas_main_intro"], st["Body"]))
        t = _component_table(scope["gas_main"])
        if t: flow.append(t)
        flow.append(Spacer(1, 0.2 * cm))

    # 4. Pilot line — only when Auto Ignition is requested
    auto_ignition = bool(customer.get("special_auto_ignition"))
    if auto_ignition and scope["pilot"]:
        flow.append(Paragraph(scope["pilot_label"], st["H2"]))
        flow.append(Paragraph(
            "We shall be supplying a gas train for Pilot Burner, which will "
            "supply required gas flow and pressure to the pilot burner for "
            "ignition of the main burner. The pilot line will consist of the "
            "following main components:", st["Body"]))
        t = _component_table(scope["pilot"])
        if t: flow.append(t)
        flow.append(Spacer(1, 0.2 * cm))

    # 5. Nitrogen purging line (only when purging is enabled)
    if has_purging:
        flow.append(Paragraph("NITROGEN PURGING LINE", st["H2"]))
        flow.append(Paragraph(
            "Pre and Post Purging of mix gas line shall be done from Nitrogen "
            "gas. The following components shall be on the purging line:",
            st["Body"]))
        t = _component_table(scope["purging"])
        if t: flow.append(t)
        flow.append(Spacer(1, 0.2 * cm))

    # 6. Temperature Control System (numbered list)
    flow.append(Paragraph("TEMPERATURE CONTROL SYSTEM", st["H2"]))
    flow.append(Paragraph(
        "To control and maintain the temperature of the ladle accurately, the "
        "thermocouple will be fitted in the Ladle at suitable location. The "
        "scheme and sequence of burner operation is described as below. "
        "Temperature control system will consist of the following main "
        "components:", st["Body"]))
    for i, item in enumerate(_temp_control_items(scope["temp"], control_mode, auto_control_type), start=1):
        flow.append(Paragraph(f"{i}. {item}", st["Numb"]))

    # 7. Operational Sequence (varies by control mode)
    flow.append(Paragraph("OPERATIONAL SEQUENCE", st["H2"]))
    flow.append(Paragraph(_operational_sequence_text(control_mode, auto_control_type), st["Body"]))

    # 7b. Pumping Unit (oil / dual fuel only — heading flips PUMPING UNIT vs
    # HEATING & PUMPING UNIT by fuel type)
    if bool(customer.get("is_oil")) or bool(customer.get("is_dual")):
        from engine.quote_writer import _pumping_unit_block
        pu_heading, pu_intro, pu_bullets = _pumping_unit_block(
            customer.get("fuel_name"),
            bool(customer.get("is_oil")),
            bool(customer.get("is_dual")),
        )
        if pu_heading:
            flow.append(Paragraph(pu_heading, st["H2"]))
            flow.append(Paragraph(pu_intro, st["Body"]))
            for b in pu_bullets:
                flow.append(Paragraph(f"&bull; {b}", st["Bullet"]))

    # 8. Painting / Cabling / Pipeline
    flow.append(Paragraph("PAINTING", st["H2"]))
    flow.append(Paragraph(
        "All steel material will be painted and surface will be cleaned "
        "according to general specifications of painting provided by the "
        "client. Final painting will be done at site after erection. All "
        "bought-out components will remain in standard supplier's finish. "
        "Pipelines shall be painted as per color code.", st["Body"]))

    flow.append(Paragraph("CABLING", st["H2"]))
    flow.append(Paragraph(
        "All necessary cabling to run the system will be provided by Client.",
        st["Body"]))

    flow.append(Paragraph("PIPELINE", st["H2"]))
    flow.append(Paragraph(
        "The scope of supply of pipeline shall include combustion air pipeline "
        "from blower to the burner. Gas pipeline from the gas train to the "
        "header and burners will be in client's scope.", st["Body"]))

    # ── Make List ────────────────────────────────────────────────────
    flow.append(Spacer(1, 0.4 * cm))
    flow.append(Paragraph("Make List", st["H1"]))
    mlt = _make_list_table(quote_data.get("make_list") or [], st)
    if mlt: flow.append(mlt)

    # ── Commercial summary ───────────────────────────────────────────
    flow.append(Spacer(1, 0.4 * cm))
    flow.append(Paragraph("Commercial Summary", st["H1"]))
    money_rows = [
        ("Subtotal",                f"Rs. {float(quote_data.get('subtotal', grand)):,.2f}"),
        ("GST (Extra)",             "18% on basic value"),
        ("Grand Total (excl. GST)", f"Rs. {grand:,.2f}"),
    ]
    if customer.get("total_in_words"):
        money_rows.append(("In Words", customer["total_in_words"]))
    t = _kv_table(money_rows, st)
    if t: flow.append(t)

    # ── Standard terms ───────────────────────────────────────────────
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

    doc.build(flow, onFirstPage=_draw_footer, onLaterPages=_draw_footer)
    _post_process_total_pages(output_path)


# ─────────────── helpers ───────────────
def _operational_sequence_text(control_mode: str | None, auto_control_type: str | None):
    cm = (control_mode or "automatic").lower()
    act = (auto_control_type or "plc").lower()
    if cm == "manual":
        return (
            "The temperature of the ladle will be monitored manually through "
            "the temperature indicator fitted on the panel. The operator will "
            "open/close the air and gas control valves to maintain the desired "
            "temperature profile of the ladle as per the heating schedule."
        )
    if act == "pid":
        return (
            "The temperature of the ladle will be controlled automatically "
            "through a PID controller. The thermocouple fitted in the ladle "
            "will sense the temperature and feed it to the PID controller. "
            "The PID controller will modulate the air control valve via the "
            "Air-Gas Ratio regulator to maintain the air/gas ratio as the "
            "temperature rises/falls to the set values."
        )
    if act == "plc_agr":
        return (
            "The temperature of the ladle will be controlled automatically "
            "through P.L.C. The thermocouple fitted in the ladle will sense "
            "the temperature and signal the P.L.C. The P.L.C will modulate "
            "the air control valve through the Air-Gas Ratio regulator and "
            "accordingly the gas flow will be controlled to maintain the "
            "air/gas ratio as the temperature decreases/increases to the set "
            "values."
        )
    # default: PLC
    return (
        "The temperature of the ladle will be controlled automatically through "
        "P.L.C. The thermocouple fitted in the ladle will sense the temperature "
        "and will give signal to the P.L.C. The P.L.C will send a signal to "
        "the control valve fitted on the airline; the air control valve will "
        "be modulated and accordingly the gas flow will be controlled, "
        "maintaining the mass-flow air/gas ratio as the temperature "
        "decreases/increases to the set values."
    )


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
    if hood_movement and str(hood_movement).strip():
        return hood_movement
    return {
        "up_down":       "Up and Down (hydraulic)",
        "swivel_manual": "Swivelling — Manual",
        "swivel_geared": "Swivelling — Geared",
        "swivel":        "Swivelling",
    }.get((hood_type or "").lower(), hood_type or "")
