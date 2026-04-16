import os
from export.word_writer import generate_word_offer

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_PATH = os.path.join(BASE_DIR, "Offer_Template.docx")


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
        "technical_person":      customer.get("technical_person", ""),
        "technical_phone":       customer.get("technical_phone", ""),
        # Pricing
        "total_price_vertical":  f"{total_vertical:,.2f}",
        "total_price_horizontal": f"{total_horizontal:,.2f}" if total_horizontal else "N/A",
        "total_in_words": (
            customer.get("total_in_words")
            or amount_in_words_indian(quote_data.get("grand_total", 0)) + " ONLY"
        ),
        "equipment_name": (
            f"{('Vertical' if not customer.get('is_horizontal') else 'Horizontal')} "
            f"Ladle Preheater – {customer.get('ladle_tons') or ''} Ton"
            if customer.get("ladle_tons")
            else (customer.get("project_name") or "")
        ),
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
        "fuel_cv":               customer.get("fuel_cv") or "",
        "fuel_consumption":      customer.get("fuel_consumption") or "",
        "burner_model":          customer.get("burner_model") or "",
        "blower_model":          customer.get("blower_model") or "",
        "blower_size":           customer.get("blower_size") or "",
        "blower_capacity":       customer.get("blower_capacity") or "",
        "hydraulic_motor_hp":    customer.get("hydraulic_motor_hp") or "",
        "max_electrical_load":   customer.get("max_electrical_load") or "",
        # Fuel-type flags drive {%p if is_oil %} / {%p if is_gas %} blocks
        # in the scope-of-supply section of the template.
        "is_oil":                bool(customer.get("is_oil")),
        "is_gas":                not bool(customer.get("is_oil")),
        # BOM items list for the MAKE LIST table on the last page.
        # Each entry: {"item": "ITEM NAME", "make": "VENDOR" or "ENCON"}.
        "make_list":             [
            {"item": x.get("item", ""), "make": x.get("make") or "ENCON"}
            for x in (customer.get("bom_items") or [])
            if x.get("item")
        ],
    }

    buffer = generate_word_offer(TEMPLATE_PATH, context)
    with open(output_path, "wb") as f:
        f.write(buffer.read())

    # Post-process: drop any tech-data row whose value cell ended up empty.
    _strip_empty_tech_rows(output_path)

    # Post-process: append BOM items (item, make) to the MAKE LIST table.
    _append_make_list(output_path, context.get("make_list", []))


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
            if cells[:2] == ["ITEM", "MAKE"]:
                target_table = t
                break
    if target_table is None:
        return

    header_row = target_table.rows[0]
    seen = set()  # de-dupe by item name (case-insensitive)
    for entry in items:
        item = (entry.get("item") or "").strip()
        make = (entry.get("make") or "ENCON").strip() or "ENCON"
        if not item:
            continue
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
        "Ladle Dimensions", "Refractory Weight", "Heating Schedule",
        "Calorific Value of Fuel", "Fuel Consumption",
        "Burner Size & Capacity", "Combustion Air Blower",
        "Blower Size", "Capacity of Blower",
        "Motor recommended for Power Pack", "Maximum Electrical Load",
    }

    for table in doc.tables:
        labels_in_table = {row.cells[0].text.strip() for row in table.rows if len(row.cells) >= 2}
        if not (labels_in_table & {"Refractory Weight", "Heating Schedule"}):
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
