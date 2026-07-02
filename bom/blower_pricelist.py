"""
Blower internal costing, priced from the pricelist.

Mirrors the HPU model (see bom/hpu_pricelist.py): blower_dm_idm_master holds the
BOM per blower model (a quantity per component), and every component's RATE is
pulled live from the pricelist (component_price_master) so editing a rate in the
Price Master reprices every blower that uses it.

blower_dm_idm_master is a WIDE table — one row per model with a <comp>_qty and
<comp>_cost column per component. We read only the qty and recompute the cost
from the live pricelist rate. All blower components are pre-existing raw
materials, so nothing new is seeded.

Cost chain (as in the source sheet):
    material subtotal = Σ(qty × pricelist rate)
    factory cost      = subtotal × 1.3   (30% overhead)
    selling price     = factory cost × 1.8
"""

import sqlite3

OVERHEAD = 1.3   # material subtotal -> factory cost
MARKUP   = 1.8   # factory cost -> selling price

# blower_dm_idm_master column prefix -> (pricelist item, display label, unit)
COMPONENTS = [
    ("angle65_50",    "M.S. Angle 65,50",            "M.S. Angle 65,50",           "kg"),
    ("channel",       "M.S. Channel",                "M.S. Channel",               "kg"),
    ("sheet8mm",      "M.S. Sheet 8mm",              "M.S. Sheet 8mm",             "kg"),
    ("sheet4mm",      "M.S. Sheet 4mm",              "M.S. Sheet 4mm",             "kg"),
    ("sheet2mm",      "M.S. Sheet 2mm",              "M.S. Sheet 2mm",             "kg"),
    ("flat",          "M.S. Flat",                   "M.S. Flat",                  "kg"),
    ("ms_round",      "M.S. Round",                  "M.S. Round",                 "kg"),
    ("ci_hub",        "C.I. Hub",                    "C.I. Hub",                   "kg"),
    ("coupling",      "Coupling",                    "Coupling",                   "kg"),
    ("plumber_block", "Plumber block with Bearing",  "Plumber Block with Bearing", "nos"),
    ("hardware",      "Hardware Bolt",               "Hardware Bolt",              "kg"),
]

# Display order of the sections (matches blower_dm_idm_master.section).
SECTION_ORDER = ["BLOWER DM 28", "BLOWER DM 40", "BLOWER IDM"]


def load_rates(conn: sqlite3.Connection) -> dict:
    """{pricelist item -> price} for every blower component."""
    names = tuple(dict.fromkeys(c[1] for c in COMPONENTS))
    q = ("SELECT item, price FROM component_price_master WHERE item IN (%s)"
         % ",".join("?" * len(names)))
    return {r[0]: (r[1] or 0.0) for r in conn.execute(q, names)}


def _f(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def compute_blower(row: dict, rates: dict) -> dict:
    """Build the priced breakdown for one blower_dm_idm_master row.

    Returns {model, items:[{s_no,item,qty,unit,rate,rate_ref,amount}],
             subtotal, factory_cost, selling}."""
    items = []
    subtotal = 0.0
    for prefix, price_item, label, unit in COMPONENTS:
        qty = _f(row.get(prefix + "_qty"))
        if qty <= 0:
            continue
        rate = rates.get(price_item, 0.0)
        amount = qty * rate
        subtotal += amount
        items.append({
            "s_no": len(items) + 1,
            "item": label,
            "qty": round(qty, 2),
            "unit": unit,
            "rate": round(rate, 2),
            "rate_ref": price_item,     # pricelist row the rate came from
            "amount": round(amount, 2),
        })
    factory = subtotal * OVERHEAD
    return {
        "model":        row.get("model"),
        "items":        items,
        "subtotal":     round(subtotal),
        "factory_cost": round(factory),
        "selling":      round(factory * MARKUP),
    }
