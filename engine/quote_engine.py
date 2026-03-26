import sqlite3
import os
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.path.join(BASE_DIR, "vlph.db")


def _lookup_price(cursor, item: dict) -> float:
    product_type = item["product_type"]
    model = item["model"]
    with_motor = item.get("with_motor", False)
    variant = item.get("variant", "Duplex 1")

    if product_type == "Horizontal Ladle Preheater":
        cursor.execute(
            "SELECT COALESCE(SUM(amount), 0) FROM horizontal_master "
            "WHERE model=? AND amount IS NOT NULL "
            "AND particular NOT IN ('COMBUSTION EQUIPMENT:', 'S.NO.')",
            (model,),
        )
        return float(cursor.fetchone()[0])

    elif product_type == "Vertical Ladle Preheater":
        cursor.execute(
            "SELECT COALESCE(SUM(amount), 0) FROM vertical_master "
            "WHERE model=? AND amount IS NOT NULL "
            "AND particular NOT IN ('COMBUSTION EQUIPMENT:', 'S.NO.')",
            (model,),
        )
        return float(cursor.fetchone()[0])

    elif product_type == "Blower":
        col = "price__with_motor" if with_motor else "price_without_motor"
        cursor.execute(
            f"SELECT {col} FROM blower_pricelist_master WHERE model=? AND {col} IS NOT NULL LIMIT 1",
            (model,),
        )
        row = cursor.fetchone()
        return float(row[0]) if row and row[0] else 0.0

    elif product_type == "HPU":
        cursor.execute(
            "SELECT COALESCE(SUM(amount), 0) FROM hpu_master "
            "WHERE unit_kw=? AND variant=? AND amount IS NOT NULL",
            (model, variant),
        )
        return float(cursor.fetchone()[0])

    elif "Burner" in product_type:
        section_filter = "FILM" if "Film" in product_type else "DUAL"
        cursor.execute(
            "SELECT price FROM burner_pricelist_master "
            "WHERE burner_size=? AND component='BURNER ALONE' AND section LIKE ? LIMIT 1",
            (model, f"%{section_filter}%"),
        )
        row = cursor.fetchone()
        return float(row[0]) if row and row[0] else 0.0

    elif product_type == "Recuperator":
        cursor.execute(
            "SELECT COALESCE(SUM(amount), 0) FROM recuperator_master "
            "WHERE model=? AND amount IS NOT NULL",
            (model,),
        )
        return float(cursor.fetchone()[0])

    elif product_type == "Rad Heat":
        cursor.execute(
            "SELECT price_with_ss_tubing FROM rad_heat_master "
            "WHERE item=? AND section='MODEL' LIMIT 1",
            (model,),
        )
        row = cursor.fetchone()
        return float(row[0]) if row and row[0] else 0.0

    elif product_type == "GAIL Gas Burner":
        cursor.execute(
            "SELECT burner_set FROM gail_gas_burner_master WHERE burner_size=? LIMIT 1",
            (model,),
        )
        row = cursor.fetchone()
        return float(row[0]) if row and row[0] else 0.0

    return 0.0


def calculate_quote(form_data: dict) -> dict:
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    priced_items = []
    subtotal = 0.0

    for item in form_data["items"]:
        qty = item.get("qty", 1)
        # Use pre-set price if provided (e.g. from VLPH costing), else look up from DB
        if item.get("unit_price") is not None and float(item["unit_price"]) > 0:
            unit_price = float(item["unit_price"])
        else:
            unit_price = _lookup_price(cursor, item)
        total = unit_price * qty
        subtotal += total
        priced_items.append({
            "product_type": item["product_type"],
            "model": item["model"],
            "description": item.get("description") or item["model"],
            "qty": qty,
            "unit_price": round(unit_price, 2),
            "total": round(total, 2),
        })

    conn.close()

    gst_pct = form_data.get("gst_percent", 18)
    freight = float(form_data.get("freight", 0))
    gst_amount = subtotal * gst_pct / 100
    grand_total = subtotal + gst_amount + freight

    seq = form_data["quote_seq"]
    quote_no = f"ENCON/Q/{datetime.now().year}/{seq}"

    return {
        "quote_no": quote_no,
        "date": datetime.now().strftime("%d-%m-%Y"),
        "customer": form_data["customer"],
        "items": priced_items,
        "subtotal": round(subtotal, 2),
        "gst_percent": gst_pct,
        "gst_amount": round(gst_amount, 2),
        "freight": freight,
        "grand_total": round(grand_total, 2),
        "valid_days": form_data.get("valid_days", 30),
    }
