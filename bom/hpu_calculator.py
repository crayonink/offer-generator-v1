"""
HPU (Heating & Pumping Unit) cost calculator.
Pulls component amounts from hpu_master table and applies selling markup.

Formula (verified against Pricelist WorkBook):
  HPU selling price = sum(hpu_master amounts for kw, Duplex 1) × 1.8

For Kg-unit items (raw materials that reference the Rates sheet), amounts are
re-computed using live rates from component_price_master so that a pricebook
upload automatically picks up rate changes.

For Nos/Mtr/Ltr items (bought-out parts), the stored amount from the Excel
upload is used directly.
"""

import re
import sqlite3
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.path.join(BASE_DIR, "vlph.db")

HPU_MARKUP = 1.8  # selling price = material cost × 1.8

# Maps lookup key → LIKE pattern in component_price_master.item
_RATE_QUERIES = {
    "hardware_bolt": "hardware bolt",
    "ms_channel":    "m.s. chanel",
    "ms_pipe":       "m.s. tube b",
    "ms_sheet_3mm":  "m.s. sheet 3mm",
    "ms_sheet_5mm":  "m.s. sheet 5mm",
    "ms_plate_5mm":  "m.s. plate 5mm",
    "ms_flat":       "m.s. flat",
    "ms_round":      "m.s. round",
}


def _load_live_rates(conn: sqlite3.Connection) -> dict:
    """Fetch raw-material Rs/kg rates from component_price_master."""
    c = conn.cursor()
    rates = {}
    for key, pattern in _RATE_QUERIES.items():
        c.execute(
            "SELECT price FROM component_price_master WHERE LOWER(item) LIKE ? LIMIT 1",
            (f"%{pattern}%",),
        )
        row = c.fetchone()
        if row and row[0]:
            rates[key] = float(row[0])
    return rates


def _live_kg_rate(item_name: str, live: dict) -> float | None:
    """
    Return live Rs/kg rate for a known raw-material HPU item.
    Returns None if item doesn't match any known category (use stored rate).
    """
    u = item_name.upper()

    if "NUT" in u or "BOLT" in u:
        return live.get("hardware_bolt")

    if "CHANNEL" in u:
        return live.get("ms_channel")

    # M.S PIPE / M.S  PIPE (heavy tube) — all sizes use same rate
    if "PIPE" in u and "FLANGE" not in u:
        return live.get("ms_pipe")

    if "HR SHEET" in u:
        if re.search(r'5\s*MM', u):
            return live.get("ms_sheet_5mm")
        return live.get("ms_sheet_3mm")

    if "PLATE" in u:
        if re.search(r'3\s*MM', u):          # thin 3mm plate = sheet rate
            return live.get("ms_sheet_3mm")
        return live.get("ms_plate_5mm")

    if "FLAT" in u:
        return live.get("ms_flat")

    if "ROUND" in u:
        return live.get("ms_round")

    return None  # bought-out or unrecognised → use stored amount


def get_hpu_cost(required_kw: float) -> dict:
    """
    Calculate HPU selling price for given KW requirement.
    Selects the smallest available KW >= required_kw.

    Returns dict with keys: kw, model, material_cost, price
    """
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.execute(
        "SELECT DISTINCT unit_kw FROM hpu_master ORDER BY unit_kw ASC"
    )
    available = [r[0] for r in cursor.fetchall()]

    if not available:
        conn.close()
        raise ValueError("No HPU data found in hpu_master")

    selected_kw = next((k for k in available if k >= required_kw), available[-1])

    cursor.execute(
        """SELECT item, qty, unit, rate, amount
           FROM hpu_master
           WHERE unit_kw = ? AND variant = 'Duplex 1'""",
        (selected_kw,),
    )
    rows = cursor.fetchall()

    live = _load_live_rates(conn)
    conn.close()

    material_cost = 0.0
    for item, qty, unit, stored_rate, stored_amount in rows:
        unit_str = (unit or "").strip().lower()
        if unit_str == "kg" and qty:
            live_rate = _live_kg_rate(item or "", live)
            if live_rate is not None:
                material_cost += qty * live_rate
            else:
                # Fallback: use stored rate if available, else stored amount
                material_cost += qty * (stored_rate or 0) if stored_rate else (stored_amount or 0)
        else:
            material_cost += stored_amount or 0.0

    selling_price = round(material_cost * HPU_MARKUP, 2)

    return {
        "kw":            selected_kw,
        "model":         f"HPD-{selected_kw}",
        "material_cost": round(material_cost, 2),
        "price":         selling_price,
    }


if __name__ == "__main__":
    print(f"HPU Markup: {HPU_MARKUP}×\n")
    for kw in [3, 6, 9, 12, 16, 20, 24, 30, 36, 48]:
        result = get_hpu_cost(kw)
        print(f"  HPD-{kw:>2}KW  material={result['material_cost']:>10,.2f}  "
              f"selling={result['price']:>10,.2f}")
