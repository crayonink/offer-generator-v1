# bom/snsf_brf_builder.py
"""
BOM builder for SNSF BRF (Billet Reheating Furnace, 30 Ton).
Per-item markup multipliers as per the costing breakup sheet.
"""

import pandas as pd


# ── Main BOM items (from Breakup sheet) ─────────────────────────────────────
BRF_ITEMS = [
    # Combustion Equipment
    {"sno": "1a", "section": "Combustion Equipment",
     "item": "Burner Set 85 Nm3/hr (Dual Fuel)", "qty": 14, "unit": "Set",
     "unit_price": 31000, "markup": 1.8},
    {"sno": "1b", "section": "Combustion Equipment",
     "item": "Blower 120 HP 40\"", "qty": 2, "unit": "Nos",
     "unit_price": 652800, "markup": 1.6},
    {"sno": "1c", "section": "Combustion Equipment",
     "item": "Gas Train (1200 Nm3/hr)", "qty": 1, "unit": "No.",
     "unit_price": 475000, "markup": 1.5},
    {"sno": "1d", "section": "Combustion Equipment",
     "item": "Recuperator", "qty": 1, "unit": "No.",
     "unit_price": 1449830, "markup": 1.8},

    # Hydraulics
    {"sno": "2", "section": "Hydraulics",
     "item": "Hydraulic Powerpack Charging Grid with Cylinder", "qty": 1, "unit": "Set",
     "unit_price": 2500000, "markup": 1.8},

    # CI Casting
    {"sno": "3", "section": "CI Casting",
     "item": "CI Casting for Doors (Inspection, Bottom Plate & Discharge)", "qty": 3000, "unit": "Kg",
     "unit_price": 150, "markup": 1.8},

    # CI Hanger
    {"sno": "4", "section": "CI Hanger",
     "item": "CI Hanger 2400 PC (3 kg per piece)", "qty": 7300, "unit": "kg",
     "unit_price": 150, "markup": 1.8},

    # Pneumatic Cylinder
    {"sno": "5", "section": "Door Mechanism",
     "item": "Pneumatic Cylinder with Door Lifting Arrangement", "qty": 3, "unit": "Set",
     "unit_price": 80000, "markup": 1.8},

    # Pinch Roll
    {"sno": "6a", "section": "Material Handling",
     "item": "Pinch Roll", "qty": 1, "unit": "Set",
     "unit_price": 1000000, "markup": 1.5},
    {"sno": "6b", "section": "Material Handling",
     "item": "Ejector with Operator Seating Arrangement", "qty": 1, "unit": "Set",
     "unit_price": 1000000, "markup": 1.8},

    # Automation
    {"sno": "7a", "section": "Automation",
     "item": "Control Panel with PLC", "qty": 1, "unit": "Set",
     "unit_price": 1600000, "markup": 1.8},
    {"sno": "7b", "section": "Automation",
     "item": "Mass Flow Control", "qty": 1, "unit": "Set",
     "unit_price": 2851020, "markup": 1.8},

    # Design & Engineering
    {"sno": "8", "section": "Engineering",
     "item": "Design, Engineering & Purchase Support", "qty": 1, "unit": "Set",
     "unit_price": 2000000, "markup": 1.0},
]

# ── Optional NG items ───────────────────────────────────────────────────────
BRF_NG_OPTIONAL = [
    {"sno": "9", "section": "NG Optional",
     "item": "Pilot Burner with Accessories", "qty": 14, "unit": "Set",
     "unit_price": 25000, "markup": 1.8},
    {"sno": "10", "section": "NG Optional",
     "item": "UV Sensor", "qty": 14, "unit": "Set",
     "unit_price": 8000, "markup": 1.8},
]

# ── Client scope items ──────────────────────────────────────────────────────
BRF_CLIENT_SCOPE = [
    {"sno": "11", "section": "Client Scope",
     "item": "Mild Steel", "qty": 107.846, "unit": "Ton",
     "unit_price": 55000, "markup": 1.8},
    {"sno": "12", "section": "Client Scope",
     "item": "Refractory", "qty": 681.872, "unit": "Ton",
     "unit_price": 25300, "markup": 1.3},
]


def build_snsf_brf_df(
    include_ng_optional: bool = False,
    include_client_scope: bool = False,
) -> pd.DataFrame:
    """
    Build BOM for SNSF BRF 30 Ton Billet Reheating Furnace.
    Each item has its own markup multiplier.
    """
    items = list(BRF_ITEMS)

    if include_ng_optional:
        items.extend(BRF_NG_OPTIONAL)

    if include_client_scope:
        items.extend(BRF_CLIENT_SCOPE)

    rows = []
    for it in items:
        cost = round(it["qty"] * it["unit_price"], 0)
        sell = round(cost * it["markup"], 0)
        rows.append((
            it["section"], it["item"], it["qty"], it["unit"],
            it["unit_price"], cost, it["markup"], sell
        ))

    df = pd.DataFrame(rows, columns=[
        "SECTION", "ITEM", "QTY", "UNIT",
        "UNIT PRICE", "COST PRICE", "MARKUP", "SELL PRICE"
    ])

    # Summary
    main_cost = df.loc[~df["SECTION"].isin(["NG Optional", "Client Scope"]), "COST PRICE"].sum()
    main_sell = df.loc[~df["SECTION"].isin(["NG Optional", "Client Scope"]), "SELL PRICE"].sum()

    ng_cost = df.loc[df["SECTION"] == "NG Optional", "COST PRICE"].sum() if include_ng_optional else 0
    ng_sell = df.loc[df["SECTION"] == "NG Optional", "SELL PRICE"].sum() if include_ng_optional else 0

    client_cost = df.loc[df["SECTION"] == "Client Scope", "COST PRICE"].sum() if include_client_scope else 0
    client_sell = df.loc[df["SECTION"] == "Client Scope", "SELL PRICE"].sum() if include_client_scope else 0

    summary = {
        "main_cost": round(main_cost, 0),
        "main_sell": round(main_sell, 0),
        "ng_optional_cost": round(ng_cost, 0),
        "ng_optional_sell": round(ng_sell, 0),
        "client_scope_cost": round(client_cost, 0),
        "client_scope_sell": round(client_sell, 0),
        "grand_cost": round(main_cost + ng_cost + client_cost, 0),
        "grand_sell": round(main_sell + ng_sell + client_sell, 0),
    }

    return df, summary
