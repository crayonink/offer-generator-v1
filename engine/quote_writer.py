import os
from export.word_writer import generate_word_offer

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_PATH = os.path.join(BASE_DIR, "Offer_Template.docx")


def generate_quote_docx(quote_data: dict, output_path: str):
    customer = quote_data["customer"]

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
        "total_price_vertical":  f"₹ {total_vertical:,.0f}",
        "total_price_horizontal": f"₹ {total_horizontal:,.0f}" if total_horizontal else "N/A",
        "total_in_words":        customer.get("total_in_words", ""),
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
    }

    buffer = generate_word_offer(TEMPLATE_PATH, context)
    with open(output_path, "wb") as f:
        f.write(buffer.read())

    # Post-process: drop any tech-data row whose value cell ended up empty.
    _strip_empty_tech_rows(output_path)


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
