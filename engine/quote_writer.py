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
        "project_name":          customer.get("project_name") or "",
        "company_name":          customer.get("company_name") or "",
        "company_address":       customer.get("address") or "",
        "mobile_no":             customer.get("mobile_no") or "",
        "poc_name":              customer.get("poc_name") or "",
        "poc_designation":       customer.get("poc_designation") or "",
        "total_price_vertical":  f"₹ {total_vertical:,.0f}",
        "total_price_horizontal": f"₹ {total_horizontal:,.0f}" if total_horizontal else "N/A",
        # Extra fields available if template is updated later
        "quote_no":    quote_data.get("quote_no", ""),
        "date":        quote_data.get("date", ""),
        "ref_no":      customer.get("ref_no", ""),
        "email":       customer.get("email", ""),
        "gstin":       customer.get("gstin", ""),
        "subtotal":    f"₹ {quote_data.get('subtotal', 0):,.0f}",
        "gst_percent": quote_data.get("gst_percent", 18),
        "gst_amount":  f"₹ {quote_data.get('gst_amount', 0):,.0f}",
        "freight":     f"₹ {quote_data.get('freight', 0):,.0f}",
        "grand_total": f"₹ {quote_data.get('grand_total', 0):,.0f}",
        "valid_days":  quote_data.get("valid_days", 30),
        "items":       quote_data.get("items", []),
    }

    buffer = generate_word_offer(TEMPLATE_PATH, context)
    with open(output_path, "wb") as f:
        f.write(buffer.read())
