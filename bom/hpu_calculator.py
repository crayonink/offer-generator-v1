"""
HPU (Heating & Pumping Unit) cost calculator.
Pulls component amounts from hpu_master table and applies selling markup.

Formula (verified against Pricelist WorkBook):
  HPU selling price = sum(hpu_master amounts for kw, Duplex 1) × 1.8
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

    # Available KW variants
    cursor.execute(
        "SELECT DISTINCT unit_kw FROM hpu_master ORDER BY unit_kw ASC"
    )
    available = [r[0] for r in cursor.fetchall()]

    if not available:
        conn.close()
        raise ValueError("No HPU data found in hpu_master")

    # Pick smallest KW >= required; fall back to largest if none qualify
    selected_kw = next((k for k in available if k >= required_kw), available[-1])

    # Sum material amounts for Duplex 1 variant
    cursor.execute(
        """
        SELECT SUM(amount)
        FROM hpu_master
        WHERE unit_kw = ?
          AND variant = 'Duplex 1'
        """,
        (selected_kw,),
    )
    row = cursor.fetchone()
    conn.close()

    material_cost = row[0] if row and row[0] else 0.0
    selling_price = round(material_cost * HPU_MARKUP, 2)

    return {
        "kw": selected_kw,
        "model": f"HPD-{selected_kw}",
        "material_cost": round(material_cost, 2),
        "price": selling_price,
    }


if __name__ == "__main__":
    for kw in [3, 6, 12, 20]:
        result = get_hpu_cost(kw)
        print(f"HPD-{kw}KW → material={result['material_cost']:,.2f}  "
              f"selling={result['price']:,.2f}")
