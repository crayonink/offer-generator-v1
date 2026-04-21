import os
from export.word_writer import generate_word_offer

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_PATH = os.path.join(BASE_DIR, "Offer_Template.docx")
TUNDISH_TEMPLATE_PATH = os.path.join(BASE_DIR, "Tundish_Offer_Template.docx")


# ── Indian-English number-to-words (lakh / crore system) ──────────────────
_ONES = ("", "ONE", "TWO", "THREE", "FOUR", "FIVE", "SIX", "SEVEN", "EIGHT",
         "NINE", "TEN", "ELEVEN", "TWELVE", "THIRTEEN", "FOURTEEN", "FIFTEEN",
         "SIXTEEN", "SEVENTEEN", "EIGHTEEN", "NINETEEN")
_TENS = ("", "", "TWENTY", "THIRTY", "FORTY", "FIFTY", "SIXTY", "SEVENTY",
         "EIGHTY", "NINETY")


def _two_digits(n: int) -> str:
    if n < 20:
        return _ONES[n]
    t, o = divmod(n, 10)
    return _TENS[t] + ("-" + _ONES[o] if o else "")


def _three_digits(n: int) -> str:
    h, r = divmod(n, 100)
    parts = []
    if h:
        parts.append(_ONES[h] + " HUNDRED")
    if r:
        parts.append(_two_digits(r))
    return " ".join(parts)


def amount_in_words_indian(amount) -> str:
    """Format an integer rupee amount in Indian English words (lakh / crore)."""
    try:
        n = int(round(float(amount)))
    except (TypeError, ValueError):
        return ""
    if n == 0:
        return "ZERO"
    crore, rem  = divmod(n, 10000000)
    lakh,  rem  = divmod(rem, 100000)
    thou,  rem  = divmod(rem, 1000)
    parts = []
    if crore:
        parts.append(_three_digits(crore) + " CRORE")
    if lakh:
        parts.append(_two_digits(lakh) + " LAKH")
    if thou:
        parts.append(_two_digits(thou) + " THOUSAND")
    if rem:
        parts.append(_three_digits(rem))
    return " ".join(parts).strip()


def _supervision_rates() -> tuple:
    """Look up supervision-charge rates from component_price_master.
    Returns (mech_rate_str, plc_rate_str) formatted as Indian rupee strings."""
    import sqlite3
    db_path = os.path.join(BASE_DIR, "vlph.db")
    rates = {"mech": 12500, "plc": 15000}  # safe defaults
    try:
        conn = sqlite3.connect(db_path)
        for key, item in (("mech", "SUPERVISION CHARGE - MECHANICAL"),
                          ("plc",  "SUPERVISION CHARGE - PLC & INSTRUMENTATION")):
            row = conn.execute(
                "SELECT price FROM component_price_master WHERE item=? LIMIT 1",
                (item,)
            ).fetchone()
            if row:
                rates[key] = float(row[0])
        conn.close()
    except Exception:
        pass
    return f"{rates['mech']:,.2f}", f"{rates['plc']:,.2f}"


def _combine_dual(val1, val2) -> str:
    """Combine two fuel values into one string for single-column display.
    If val2 is empty, returns val1 only."""
    v1 = (val1 or "").strip()
    v2 = (val2 or "").strip()
    if v2:
        return f"{v1}\n{v2}"
    return v1


def _fuel_label(fuel_name: str) -> str:
    """Extract the first fuel name from 'NG & MG' or 'Natural Gas & Furnace Oil'."""
    if "&" in fuel_name:
        return fuel_name.split("&")[0].strip()
    return fuel_name.strip()


def _fuel_label2(fuel_name: str) -> str:
    """Extract the second fuel name from 'NG & MG'."""
    if "&" in fuel_name:
        return fuel_name.split("&")[1].strip()
    return ""


import re as _re_pilot


def _rewrite_pilot_name(item: str, pilot_gas: str) -> str:
    """The BOM carries one fixed pilot-burner model. Rewrite its name so the
    MAKE LIST reflects the fuel the user actually picked for the pilot.
    Only rewrite rows whose model starts with 'ENCON-PB' / 'ENCON PB' — do NOT
    match items like 'BALL VALVE (Pilot Burner)' that merely mention the pilot."""
    if not item:
        return item
    s = item.strip()
    upper = s.upper()
    if not (upper.startswith("ENCON-PB") or upper.startswith("ENCON PB")):
        return s
    # Detect KW rating by a number directly followed by KW (word-boundary).
    m = _re_pilot.search(r"\b(\d+)\s*KW\b", upper)
    rating = f"{m.group(1)} KW" if m else "100 KW"
    return f"ENCON-PB {pilot_gas} {rating}"


def _build_equipment_name(customer, quote_data):
    """Derive equipment name from items product_type or customer fields."""
    # Check items for product type
    for item in quote_data.get("items", []):
        pt = (item.get("product_type") or "").lower()
        if "tundish" in pt:
            tons = customer.get("ladle_tons")
            return f"Tundish Preheater – {tons} Ton" if tons else "Tundish Preheater"
        elif "horizontal" in pt:
            tons = customer.get("ladle_tons")
            return f"Horizontal Ladle Preheater – {tons} Ton" if tons else "Horizontal Ladle Preheater"
    # Default: Vertical Ladle Preheater
    tons = customer.get("ladle_tons")
    if tons:
        return f"Vertical Ladle Preheater – {tons} Ton"
    return customer.get("project_name") or ""


def generate_quote_docx(quote_data: dict, output_path: str):
    customer = quote_data["customer"]
    sup_mech, sup_plc = _supervision_rates()

    # Determine vertical vs horizontal price from items
    total_vertical   = 0.0
    total_horizontal = 0.0
    for item in quote_data.get("items", []):
        pt = (item.get("product_type") or "").lower()
        if "vertical" in pt:
            total_vertical += item.get("total", 0)
        elif "horizontal" in pt:
            total_horizontal += item.get("total", 0)
        else:
            total_vertical += item.get("total", 0)   # default to vertical

    # Map to template variable names (must match {{...}} in Offer_Template.docx)
    context = {
        # Customer / contact
        "project_name":          customer.get("project_name") or "",
        "subject":               customer.get("subject") or customer.get("project_name") or "",
        "company_name":          customer.get("company_name") or "",
        "company_city":          customer.get("company_city") or "",
        "company_state":         customer.get("company_state") or "",
        "company_address":       customer.get("address") or "",
        "mobile_no":             customer.get("mobile_no") or "",
        "email":                 customer.get("email") or "",
        "gstin":                 customer.get("gstin") or "",
        "poc_name":              customer.get("poc_name") or "",
        "poc_designation":       customer.get("poc_designation") or "",
        # Reference / enquiry
        "quote_no":              quote_data.get("quote_no", ""),
        "date":                  quote_data.get("date", ""),
        "ref_no":                customer.get("ref_no", ""),
        "your_ref":              customer.get("your_ref") or customer.get("ref_no", ""),
        "enquiry_ref":           customer.get("enquiry_ref") or customer.get("ref_no", ""),
        "marketing_person":      customer.get("marketing_person", ""),
        "marketing_phone":       customer.get("marketing_phone", ""),
        "marketing_email":       customer.get("marketing_email", ""),
        "technical_person":      customer.get("technical_person", ""),
        "technical_phone":       customer.get("technical_phone", ""),
        "technical_email":       customer.get("technical_email", ""),
        # Pricing
        "total_price_vertical":  f"{total_vertical:,.2f}",
        "total_price_horizontal": f"{total_horizontal:,.2f}" if total_horizontal else "N/A",
        "total_in_words": (
            customer.get("total_in_words")
            or amount_in_words_indian(quote_data.get("grand_total", 0)) + " ONLY"
        ),
        "equipment_name": _build_equipment_name(customer, quote_data),
        # Supervision-charge rates (from component_price_master)
        "supervision_mech": sup_mech,
        "supervision_plc":  sup_plc,
        "subtotal":              f"₹ {quote_data.get('subtotal', 0):,.0f}",
        "grand_total":           f"₹ {quote_data.get('grand_total', 0):,.0f}",
        "valid_days":            quote_data.get("valid_days", 30),
        "items":                 quote_data.get("items", []),
        # Technical data (populates the Tech Data table in the template)
        "ladle_tons":            customer.get("ladle_tons") or "",
        "ladle_dim":             customer.get("ladle_dim") or "",
        "ladle_drawing_no":      customer.get("ladle_drawing_no") or "",
        "refractory_weight_kg":  customer.get("refractory_weight_kg") or "",
        "heating_schedule":      customer.get("heating_schedule") or "",
        "fuel_cv":               _combine_dual(customer.get("fuel_cv"), customer.get("fuel2_cv")),
        "fuel_consumption":      _combine_dual(customer.get("fuel_consumption"), customer.get("fuel2_consumption")),
        "burner_model":          customer.get("burner_model") or "",
        # Display name shown in the 'ENCON GAS/OIL/DUAL BURNER' body paragraphs.
        # Prefixed per the ENCON pricelist: oil -> 'ENCON 7A', gas -> 'ENCON -G 7A',
        # dual -> 'ENCON DUAL- 7A'. Falls back to burner_model if empty.
        "burner_display_model":  (
            customer.get("burner_display_model")
            or (
                customer.get("burner_model", "").replace("ENCON ", "ENCON -G ", 1)
                if (not bool(customer.get("is_oil")) and not bool(customer.get("is_dual")))
                else customer.get("burner_model", "").replace("ENCON ", "ENCON DUAL- ", 1)
                     if bool(customer.get("is_dual"))
                else customer.get("burner_model", "")
            )
        ),
        "burner_firing_rate":    (
            customer.get("burner_firing_rate")
            or customer.get("fuel_consumption", "")
        ),
        "blower_model":          customer.get("blower_model") or "",
        "blower_size":           customer.get("blower_size") or "",
        "blower_capacity":       customer.get("blower_capacity") or "",
        "hydraulic_motor_hp":    customer.get("hydraulic_motor_hp") or "",
        "max_electrical_load":   customer.get("max_electrical_load") or "",
        # Extra tech-data fields for PLC template
        "heating_time":          customer.get("heating_time") or "",
        "fuel_name":             customer.get("fuel_name") or "",
        "burner_capacity_range": customer.get("burner_capacity_range") or "",
        "pumping_unit":          customer.get("pumping_unit") or "",
        "hood_movement":         customer.get("hood_movement") or "Vertical Swiveling through bearing mechanism.",
        "hood_type":             customer.get("hood_type") or "up_down",
        "pilot_gas_type":        customer.get("pilot_gas_type") or "LPG",
        "ignition_method":       customer.get("ignition_method") or "Automatic Through LPG Fired Pilot Burner",
        # Tundish-specific tech-data placeholders (pre-heating station)
        "fuel1_label":           _fuel_label(customer.get("fuel_name", "")),
        "fuel2_label":           _fuel_label2(customer.get("fuel_name", "")),
        "num_burners":           customer.get("num_burners") or "",
        "fuel1_cv":              customer.get("fuel_cv") or "",
        "fuel2_cv":              customer.get("fuel2_cv") or "",
        "max_temperature":       customer.get("heating_schedule") or "",
        "cycle_time":            customer.get("heating_time") or "",
        "lifting_lowering":      customer.get("hood_movement") or "With Hydraulic cylinder & Power pack",
        "firing_rate_fuel1":     customer.get("fuel_consumption") or "",
        "max_consumption_fuel1": customer.get("max_fuel_consumption1") or "",
        "firing_rate_fuel2":     customer.get("fuel2_consumption") or "",
        "max_consumption_fuel2": customer.get("max_fuel_consumption2") or "",
        "combustion_blower":     "Centrifugal type",
        "motor_power_pack":      customer.get("hydraulic_motor_hp") or "10 HP",
        # Tundish drying station — defaults to same as pre-heating (user edits in Word)
        "max_temperature_dry":       customer.get("max_temperature_dry") or customer.get("heating_schedule") or "",
        "cycle_time_dry":            customer.get("cycle_time_dry") or customer.get("heating_time") or "",
        "firing_rate_fuel1_dry":     customer.get("firing_rate_fuel1_dry") or customer.get("fuel_consumption") or "",
        "max_consumption_fuel1_dry": customer.get("max_consumption_fuel1_dry") or customer.get("max_fuel_consumption1") or "",
        "firing_rate_fuel2_dry":     customer.get("firing_rate_fuel2_dry") or customer.get("fuel2_consumption") or "",
        "max_consumption_fuel2_dry": customer.get("max_consumption_fuel2_dry") or customer.get("max_fuel_consumption2") or "",
        "blower_size_dry":           customer.get("blower_size_dry") or customer.get("blower_capacity") or "",
        "max_electrical_load_dry":   customer.get("max_electrical_load_dry") or customer.get("max_electrical_load") or "",
        # Fuel-type flags drive {%p if is_oil %} / {%p if is_gas %} blocks
        # in the scope-of-supply section of the template.
        # Mutually-exclusive burner-type flags: exactly one of is_oil/is_gas/is_dual is true.
        "is_dual":               bool(customer.get("is_dual")),
        "is_oil":                bool(customer.get("is_oil")) and not bool(customer.get("is_dual")),
        "is_gas":                not bool(customer.get("is_oil")) and not bool(customer.get("is_dual")),
        # Label used in 'On the main ___ pipeline:' heading. Dual fuel keeps 'oil'
        # since the oil-line components (flow meter, manual ball valve, etc.) still apply.
        "fuel_line_label":       (
            "gas" if (not bool(customer.get("is_oil")) and not bool(customer.get("is_dual")))
            else "oil"
        ),
        # Full PIPELINE sentence tail — 'gas pipeline from gas train to burner' for
        # gas-only, otherwise 'LDO pipeline from LDO pumping unit to burner'.
        "fuel_delivery_text":    (
            "gas pipeline from gas train to burner"
            if (not bool(customer.get("is_oil")) and not bool(customer.get("is_dual")))
            else "LDO pipeline from LDO pumping unit to burner"
        ),
        # ─── Scope of Supply (Annexure I) parameters ─────────────────────────
        # Fuel tag next to 'Burner with Burner block & mounting Plate –'
        "fuel_short": (
            "NG"     if (not bool(customer.get("is_oil")) and not bool(customer.get("is_dual")))
            else "LDO/NG" if bool(customer.get("is_dual"))
            else "LDO"
        ),
        # Burner product name shown in row 3 of the scope table
        "burner_scope_name": (
            "ENCON Dual-Fuel Burner" if bool(customer.get("is_dual"))
            else "ENCON Gas Burner"  if not bool(customer.get("is_oil"))
            else "ENCON IIP Film Burner"
        ),
        # Row 5 — fuel delivery unit
        "pumping_scope_text": (
            "Gas train with pressure regulating and control valves"
            if (not bool(customer.get("is_oil")) and not bool(customer.get("is_dual")))
            else "LDO pumping unit with micro valve"
        ),
        # Row 6 (vertical column) — hood mechanism
        "hood_scope_vertical": (
            "Swiveling mechanism through bearing with mechanical locking"
            if (customer.get("hood_type") == "swivel")
            else "Hydraulic Cylinder for lifting and Lowering of the system"
        ),
        # Control-mode flags drive {%p if is_manual %} / {%p if is_automatic %}
        # for temperature-control and control-panel scope sections.
        "is_manual":             customer.get("control_mode") == "manual",
        "is_automatic":          customer.get("control_mode") != "manual",
        # BOM items list for the MAKE LIST table on the last page.
        # Each entry: {"item": "ITEM NAME", "make": "VENDOR" or "ENCON"}.
        # Pilot-burner rows are rewritten to match the selected pilot fuel
        # (LPG vs NG) — the BOM itself always carries a single fixed entry.
        "make_list":             [
            {
                "item": _rewrite_pilot_name(
                    x.get("item", ""),
                    (customer.get("pilot_gas_type") or "LPG").upper()
                ),
                "make": x.get("make") or "ENCON",
            }
            for x in (customer.get("bom_items") or [])
            if x.get("item")
        ],
    }

    # Use tundish-specific template if the product type indicates tundish
    is_tundish = any(
        "tundish" in (item.get("product_type") or "").lower()
        for item in quote_data.get("items", [])
    )
    template = TUNDISH_TEMPLATE_PATH if is_tundish and os.path.exists(TUNDISH_TEMPLATE_PATH) else TEMPLATE_PATH
    buffer = generate_word_offer(template, context)
    with open(output_path, "wb") as f:
        f.write(buffer.read())

    # Post-process: drop any tech-data row whose value cell ended up empty.
    _strip_empty_tech_rows(output_path)

    # Post-process: append BOM items (item, make) to the MAKE LIST table.
    _append_make_list(output_path, context.get("make_list", []))

    # Post-process: drop the unused column in the Scope of Supply (Annexure I)
    # table — VLPH-only quotes should not show the Horizontal column and vice
    # versa. total_vertical and total_horizontal were computed above.
    _prune_scope_columns(
        output_path,
        has_vertical=(total_vertical > 0),
        has_horizontal=(total_horizontal > 0),
    )


def _append_make_list(docx_path: str, items: list):
    """Find the table whose first row is ['ITEM', 'MAKE'] and append one
    row per BOM item with item name + vendor make."""
    if not items:
        return
    from docx import Document
    from copy import deepcopy
    from docx.oxml.ns import qn

    doc = Document(docx_path)
    target_table = None
    for t in doc.tables:
        if len(t.rows) >= 1 and len(t.rows[0].cells) >= 2:
            cells = [c.text.strip() for c in t.rows[0].cells]
            # Detect both 2-col ("ITEM","MAKE") and 3-col ("S. No.","ITEM","MAKE") layouts
            if cells[:2] == ["ITEM", "MAKE"] or cells[1:3] == ["ITEM", "MAKE"]:
                target_table = t
                break
    if target_table is None:
        return

    header_row = target_table.rows[0]
    seen = set()  # de-dupe by item name (case-insensitive)

    import re as _re
    # Known vendor names that may leak into item names
    KNOWN_VENDORS = {"BAUMER", "HGURU", "WIKA", "CAIR", "DEMBLA", "AIRA",
                     "MADAS", "ENCON", "L&T", "LT", "AUDCO", "LEADER",
                     "HONEYWELL", "BENGAL IND.", "BENGAL", "ROTEX", "VFLEX",
                     "SB INTERNATIONAL", "SKF", "FAG", "ABB", "BB",
                     "CROMPTON", "MASIBUS", "MURUGAPPA", "UNIFRAX",
                     "SWITZER", "DANFOSS", "JINDAL", "TATA",
                     "MAHARASHTRA SEAMLESS", "VARITECH", "JSD", "VANAZ",
                     "LINEAR SYSTEM", "THIRD PARTY"}
    # Location/usage qualifiers we want to drop
    LOC_QUALIFIERS = {"PILOT BURNER", "UV LINE", "PILOT", "UV"}

    def _clean_item(name: str, make_str: str) -> str:
        n = name.strip()
        # 1. 'GAS TRAIN <flow> NM3/Hr' → 'GAS TRAIN'
        n = _re.sub(r"^(GAS TRAIN)\s.*$", r"\1", n, flags=_re.IGNORECASE)
        # 2. Strip ALL trailing parentheticals that look like vendors or
        #    location qualifiers. Keep technical qualifiers like (Ball),
        #    (Globe), (Butterfly), (Lever), (Gear).
        while True:
            m = _re.search(r"\s*\(([^)]+)\)\s*$", n)
            if not m:
                break
            inside = m.group(1).strip()
            inside_u = inside.upper()
            is_vendor = (
                inside_u == make_str.upper()
                or inside_u in KNOWN_VENDORS
                or inside_u in LOC_QUALIFIERS
            )
            if not is_vendor:
                break
            n = n[:m.start()].rstrip()
        return n

    for entry in items:
        item = (entry.get("item") or "").strip()
        make = (entry.get("make") or "ENCON").strip() or "ENCON"
        if not item:
            continue
        item = _clean_item(item, make)
        key = item.lower()
        if key in seen:
            continue
        seen.add(key)

        # Clone the header row so formatting matches, then overwrite text
        new_row = deepcopy(header_row._element)
        # Clear text in each <w:t>
        for t in new_row.iter(qn("w:t")):
            t.text = ""
        # Set first run text in cells 0 and 1
        cells = new_row.findall(qn("w:tc"))
        if len(cells) >= 2:
            for cell, text in ((cells[0], item), (cells[1], make)):
                # find first <w:t> in first paragraph and set it
                t_elements = list(cell.iter(qn("w:t")))
                if t_elements:
                    t_elements[0].text = text
                    t_elements[0].set(qn("xml:space"), "preserve")
        target_table._element.append(new_row)

    doc.save(docx_path)


def _strip_empty_tech_rows(docx_path: str):
    """In the rendered offer, find the Technical Data table and remove rows
    whose value cell is blank (meaning the placeholder resolved to empty)."""
    from docx import Document
    doc = Document(docx_path)

    # Identify the tech-data table by looking for the unique label 'Refractory Weight'.
    tech_labels = {
        "Ladle Dimensions", "Ladle Type", "Reference TS",
        "Refractory Weight", "Weight of refractory lining",
        "Heating Schedule", "Heating Temperature", "Heating time",
        "Calorific Value of Fuel", "Calorific value of LDO", "Calorific value",
        "Fuel Consumption", "Firing Rate", "Firing Rate (R1)",
        "Burner Size & Capacity", "Burner Capacity", "Burner Capacity (R1)",
        "ENCON Burner", "Combustion Air Blower",
        "Blower Size", "Capacity of Blower", "Blower Capacity / Motor rating",
        "LDO Pumping Unit", "Pumping Unit",
        "Movement of Hood",
        "Motor recommended for Power Pack", "Maximum Electrical Load",
        "Maximum Electrical Load Required",
    }

    for table in doc.tables:
        labels_in_table = {row.cells[0].text.strip() for row in table.rows if len(row.cells) >= 2}
        if not (labels_in_table & {"Refractory Weight", "Weight of refractory lining",
                                    "Heating Schedule", "Heating Temperature"}):
            continue
        # This is the tech-data table — drop blank-value rows we own.
        rows_to_remove = []
        for row in table.rows:
            if len(row.cells) < 2:
                continue
            label = row.cells[0].text.strip()
            value = row.cells[1].text.strip()
            # 'Centrifugal type, ' alone (no model) is also empty
            if label in tech_labels and (not value or value == "Centrifugal type,"):
                rows_to_remove.append(row)
        for row in rows_to_remove:
            row._element.getparent().remove(row._element)
        break

    doc.save(docx_path)


def _prune_scope_columns(docx_path: str, has_vertical: bool, has_horizontal: bool):
    """Annexure I scope-of-supply has two columns: Vertical and Horizontal.
    If the quote is for only one of them, drop the other column so the
    offer shows just the relevant scope."""
    if has_vertical and has_horizontal:
        return  # both needed, nothing to do
    if not has_vertical and not has_horizontal:
        return  # nothing to base decision on

    from docx import Document
    from docx.oxml.ns import qn

    doc = Document(docx_path)

    # Find the scope table by its header row text.
    target = None
    for t in doc.tables:
        if len(t.rows) < 2 or len(t.rows[0].cells) != 2:
            continue
        header = " | ".join(c.text.strip() for c in t.rows[0].cells)
        if "Vertical Ladle" in header and "Horizontal Ladle" in header:
            target = t
            break
    if target is None:
        return

    col_to_drop = 1 if has_vertical else 0   # 1 = horizontal, 0 = vertical
    for row in list(target.rows):
        tcs = row._element.findall(qn('w:tc'))
        if col_to_drop < len(tcs):
            tcs[col_to_drop].getparent().remove(tcs[col_to_drop])

    # Also shrink the grid so the remaining column stretches across.
    tblgrid = target._element.find(qn('w:tblGrid'))
    if tblgrid is not None:
        grid_cols = tblgrid.findall(qn('w:gridCol'))
        if col_to_drop < len(grid_cols):
            tblgrid.remove(grid_cols[col_to_drop])

    doc.save(docx_path)
