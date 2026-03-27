"""
bom/burner_fabrication.py

Calculates GAIL-type burner fabrication cost from gas_burner_parts_master.

How it works:
  1. gas_burner_parts_master stores each part's amount (pre-calculated in the
     Excel, either as qty×rate or a fixed charge like labour).
  2. We GROUP parts by section (the "BURNER PARTS X NM³ ..." header rows).
  3. We SUM amounts per section → fabrication cost.
  4. Selling price = fabrication cost × BURNER_MARKUP.

When you upload a new pricebook the amounts in gas_burner_parts_master update
automatically (Excel recalculates them from the Rates sheet before export),
so the final BOM price updates with no code changes.

BURNER_MARKUP: adjust this to match ENCON's actual selling margin.
               Default 1.25 → 25% margin over fabrication cost.
"""

import re
import sqlite3
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH  = os.path.join(BASE_DIR, "vlph.db")

BURNER_MARKUP = 1.25  # fabrication cost × markup = selling price


def _parse_nm3(section_title: str):
    """
    "BURNER PARTS 50 NM³ (GAIL DESIGN HOLE TYPE)" → 50
    "BURNER PARTS 100 NM³ POWERTRADE"              → 100
    """
    m = re.search(r'(\d+)\s*NM', section_title, re.I)
    return int(m.group(1)) if m else None


def get_all_sections() -> dict:
    """
    Returns dict: {nm3_capacity: {"title": ..., "parts": [...], "total_cost": ...}}
    Groups rows by the section-header rows in gas_burner_parts_master.
    """
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT particular, qty, unit, rate, amount FROM gas_burner_parts_master")
    rows = c.fetchall()
    conn.close()

    sections = {}
    current_nm3 = None
    current_title = None
    current_parts = []

    def _flush():
        if current_nm3 is not None:
            total = sum(p["amount"] for p in current_parts if p["amount"])
            sections[current_nm3] = {
                "title": current_title,
                "parts": current_parts,
                "fabrication_cost": round(total, 2),
                "selling_price":    round(total * BURNER_MARKUP, 2),
            }

    for particular, qty, unit, rate, amount in rows:
        # Detect section-header row: amount is None and particular contains "BURNER PARTS"
        if amount is None and particular and "BURNER PARTS" in particular.upper():
            _flush()
            current_nm3   = _parse_nm3(particular)
            current_title = particular.strip()
            current_parts = []
        elif current_nm3 is not None and particular:
            current_parts.append({
                "particular": particular,
                "qty":        qty,
                "unit":       unit,
                "rate":       rate,
                "amount":     amount or 0.0,
            })

    _flush()
    return sections


def get_burner_cost(nm3_capacity: int) -> dict:
    """
    Return fabrication cost + selling price for a given NM³ capacity.
    Picks the smallest available capacity >= nm3_capacity.

    Returns:
        {nm3, title, fabrication_cost, selling_price}
    """
    sections = get_all_sections()
    if not sections:
        raise ValueError("No burner parts data in gas_burner_parts_master")

    candidates = sorted(k for k in sections if k >= nm3_capacity)
    key = candidates[0] if candidates else max(sections)

    sec = sections[key]
    return {
        "nm3":              key,
        "title":            sec["title"],
        "fabrication_cost": sec["fabrication_cost"],
        "selling_price":    sec["selling_price"],
    }


if __name__ == "__main__":
    print("Gas Burner Fabrication Costs")
    print(f"Markup: {BURNER_MARKUP}x\n")
    sections = get_all_sections()
    if not sections:
        print("No data in gas_burner_parts_master.")
    for nm3, sec in sorted(sections.items()):
        print(f"  {nm3} NM3/hr — {sec['title']}")
        print(f"    Fabrication cost : Rs {sec['fabrication_cost']:,.0f}")
        print(f"    Selling price    : Rs {sec['selling_price']:,.0f}")
        for p in sec["parts"]:
            print(f"      {p['particular']:<45} amount={p['amount']:>8,.0f}")
        print()
