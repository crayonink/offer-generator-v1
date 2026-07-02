"""
HPU (Heating & Pumping Unit) cost calculator.
Pulls the BOM (qty per line) from hpu_master and prices every line from the
pricelist (component_price_master) via bom.hpu_pricelist, then applies markup.

Formula:
  HPU material cost  = Σ(qty × pricelist rate) over the Duplex 1 BOM
                       (LABOUR keeps its stored amount — not a material)
  HPU selling price  = material cost × 1.8

Because rates come from the pricelist, editing a rate in the Price Master
reprices HPU automatically — the same source the Internal-Costing HPU tab uses.
"""

import sqlite3
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.path.join(BASE_DIR, "vlph.db")

HPU_MARKUP = 1.8  # selling price = material cost × 1.8


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

    # Every line's rate now comes from the pricelist (component_price_master),
    # via the shared resolver, so this matches the Internal-Costing HPU tab.
    # LABOUR is not a material → keep its stored amount.
    try:
        from bom.hpu_pricelist import load_rates, resolve_rate, is_labour
    except ImportError:  # running this file directly (bom/ on path, not root)
        from hpu_pricelist import load_rates, resolve_rate, is_labour
    rates = load_rates(conn)
    conn.close()

    material_cost = 0.0
    for item, qty, unit, stored_rate, stored_amount in rows:
        if is_labour(item or ""):
            material_cost += stored_amount or 0.0
            continue
        rate, _src = resolve_rate(item or "", unit or "", rates)
        material_cost += (qty or 0.0) * rate

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
