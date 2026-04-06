# bom/snsf_brf_builder.py
"""
BOM builder for SNSF BRF (Billet Reheating Furnace, 30 Ton).
Per-item markup multipliers as per the costing breakup sheet.
Includes full calculation breakdown matching legacy Excel.
"""

import pandas as pd


# ── Main BOM items (from Breakup sheet) ─────────────────────────────────────
BRF_ITEMS = [
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
    {"sno": "2", "section": "Hydraulics",
     "item": "Hydraulic Powerpack Charging Grid with Cylinder", "qty": 1, "unit": "Set",
     "unit_price": 2500000, "markup": 1.8},
    {"sno": "3", "section": "CI Casting",
     "item": "CI Casting for Doors (Inspection, Bottom Plate & Discharge)", "qty": 3000, "unit": "Kg",
     "unit_price": 150, "markup": 1.8},
    {"sno": "4", "section": "CI Hanger",
     "item": "CI Hanger 2400 PC (3 kg per piece)", "qty": 7300, "unit": "kg",
     "unit_price": 150, "markup": 1.8},
    {"sno": "5", "section": "Door Mechanism",
     "item": "Pneumatic Cylinder with Door Lifting Arrangement", "qty": 3, "unit": "Set",
     "unit_price": 80000, "markup": 1.8},
    {"sno": "6a", "section": "Material Handling",
     "item": "Pinch Roll", "qty": 1, "unit": "Set",
     "unit_price": 1000000, "markup": 1.5},
    {"sno": "6b", "section": "Material Handling",
     "item": "Ejector with Operator Seating Arrangement", "qty": 1, "unit": "Set",
     "unit_price": 1000000, "markup": 1.8},
    {"sno": "7a", "section": "Automation",
     "item": "Control Panel with PLC", "qty": 1, "unit": "Set",
     "unit_price": 1600000, "markup": 1.8},
    {"sno": "7b", "section": "Automation",
     "item": "Mass Flow Control", "qty": 1, "unit": "Set",
     "unit_price": 2851020, "markup": 1.8},
    {"sno": "8", "section": "Engineering",
     "item": "Design, Engineering & Purchase Support", "qty": 1, "unit": "Set",
     "unit_price": 2000000, "markup": 1.0},
]

BRF_NG_OPTIONAL = [
    {"sno": "9", "section": "NG Optional",
     "item": "Pilot Burner with Accessories", "qty": 14, "unit": "Set",
     "unit_price": 25000, "markup": 1.8},
    {"sno": "10", "section": "NG Optional",
     "item": "UV Sensor", "qty": 14, "unit": "Set",
     "unit_price": 8000, "markup": 1.8},
]

BRF_CLIENT_SCOPE = [
    {"sno": "11", "section": "Client Scope",
     "item": "Mild Steel", "qty": 107.846, "unit": "Ton",
     "unit_price": 55000, "markup": 1.8},
    {"sno": "12", "section": "Client Scope",
     "item": "Refractory", "qty": 681.872, "unit": "Ton",
     "unit_price": 25300, "markup": 1.3},
]


# ── Supplementary calculation data (from legacy Excel) ──────────────────────

FURNACE_CALC = {
    "furnace_capacity_tph": 30.0,
    "billet_L_mm": 4000, "billet_W_mm": 100, "billet_H_mm": 100,
    "ms_density_kg_m3": 7850,
    "billet_weight_kg": 471000,
    "inside_material_kg": 90000,
    "effective_length_m": 17.0,
    "effective_width_m": 4.5,
    "overall_width_mm": 6036,
    "overall_length_mm": 21022,
    "calorific_value_kcal_nm3": 9500,
    "reheat_temp_c": 100,
    "time_hr": 1.0,
}

MASS_FLOW_CONTROL = [
    {"zone": "Heating Zone 1 (Air)", "items": [
        {"item": "Pneumatic Flow Control Valve 200 NB", "qty": 1, "price": 213240},
        {"item": "Orifice Plate 250 NB", "qty": 1, "price": 19000},
        {"item": "DPT", "qty": 1, "price": 45000},
    ]},
    {"zone": "Heating Zone 1 (Gas)", "items": [
        {"item": "Pneumatic Flow Control Valve 80 NB", "qty": 1, "price": 105000},
        {"item": "Orifice Plate 100 NB", "qty": 1, "price": 7000},
        {"item": "DPT", "qty": 1, "price": 45000},
    ]},
    {"zone": "Heating Zone 1 (Oil)", "items": [
        {"item": "Motorized Control Valve with Actuator 25 NB", "qty": 1, "price": 40000},
        {"item": "Mass Flow Meter (Nagman) 25 NB", "qty": 1, "price": 90000},
    ]},
    {"zone": "Heating Zone 2 (Air)", "items": [
        {"item": "Pneumatic Flow Control Valve 200 NB", "qty": 1, "price": 213240},
        {"item": "Orifice Plate 250 NB", "qty": 1, "price": 19000},
        {"item": "DPT", "qty": 1, "price": 45000},
    ]},
    {"zone": "Heating Zone 2 (Gas)", "items": [
        {"item": "Pneumatic Flow Control Valve 80 NB", "qty": 1, "price": 105000},
        {"item": "Orifice Plate 100 NB", "qty": 1, "price": 7000},
        {"item": "DPT", "qty": 1, "price": 45000},
    ]},
    {"zone": "Heating Zone 2 (Oil)", "items": [
        {"item": "Motorized Control Valve with Actuator 25 NB", "qty": 1, "price": 40000},
        {"item": "Mass Flow Meter (Nagman) 25 NB", "qty": 1, "price": 90000},
    ]},
    {"zone": "Soaking Zone (Air)", "items": [
        {"item": "Pneumatic Flow Control Valve 150 NB", "qty": 1, "price": 189540},
        {"item": "Orifice Plate 200 NB", "qty": 1, "price": 15000},
        {"item": "DPT", "qty": 1, "price": 45000},
    ]},
    {"zone": "Soaking Zone (Gas)", "items": [
        {"item": "Pneumatic Flow Control Valve 80 NB", "qty": 1, "price": 105000},
        {"item": "Orifice Plate 80 NB", "qty": 1, "price": 7000},
        {"item": "DPT", "qty": 1, "price": 45000},
    ]},
    {"zone": "Soaking Zone (Oil)", "items": [
        {"item": "Motorized Control Valve with Actuator 25 NB", "qty": 1, "price": 40000},
        {"item": "Mass Flow Meter (Nagman) 25 NB", "qty": 1, "price": 90000},
    ]},
    {"zone": "Furnace", "items": [
        {"item": "RTD", "qty": 3, "price": 2000},
        {"item": "DPT", "qty": 1, "price": 45000},
        {"item": "Motorized Damper", "qty": 1, "price": 250000},
        {"item": "Solenoid Valve (Gas Line) 65 NB", "qty": 15, "price": 46000},
        {"item": "Solenoid Valve (Oil Line) 20 NB", "qty": 15, "price": 5000},
        {"item": "TT in Gas Line", "qty": 1, "price": 15000},
        {"item": "PT in Gas Line", "qty": 1, "price": 45000},
        {"item": "TT in Air Line", "qty": 1, "price": 15000},
        {"item": "PT in Air Line", "qty": 1, "price": 45000},
        {"item": "PLC S7-1500", "qty": 1, "price": 800000},
        {"item": "Control Panel", "qty": 1, "price": 800000},
    ]},
]

RECUPERATOR_CALC = {
    "title": "30 TPH NG Based Furnace",
    "total_flue_gas_nm3hr": 7700,
    "total_mass_flue_gas_kghr": 9240,
    "specific_heat_flue_gas": 0.23,
    "initial_temp_flue_gas_c": 650.0,
    "final_temp_flue_gas_c": 337.59,
    "heat_transfer_coeff": 30.0,
    "air_volume_nm3hr": 7000,
    "initial_air_temp_c": 30.0,
    "final_air_temp_c": 350.0,
    "specific_heat_air": 0.247,
    "heat_required_kcal": 663936,
    "lmtd_c": 273.40,
    "surface_area_m2": 80.95,
    "bank_length_mm": 615.8,
    "bank_width_mm": 856.7,
    "bank_gap_mm": 150.0,
    "pipes_total": 216,
    "pipes_per_row": 12,
    "pipes_per_column": 18,
    "pipe_dia_mm": 48.3,
    "pipe_length_m": 2.575,
    "pipe_thickness_mm": 3.6,
    "pipe_weight_kg_m": 4.72,
    "hot_bank_weight_kg": 1312.63,
    "material": "SS-316 Seamless",
}

COMBUSTION_ITEMS = [
    {"item": "Recuperator", "qty": 1, "price": 1518440},
    {"item": "Blower 50 HP/40\"", "qty": 2, "price": 652800},
    {"item": "Burner Set 85 L/hr (Dual Fuel)", "qty": 15, "price": 31000},
    {"item": "Gas Train (700 Nm3/hr)", "qty": 1, "price": 475000},
    {"item": "Mass Flow Control", "qty": 1, "price": 4451020},
    {"item": "H&P Unit", "qty": 1, "price": 105000},
    {"item": "Pneumatically Operated Doors", "qty": 3, "price": 100000},
    {"item": "Ejector + Operator Seating", "qty": 1, "price": 1000000},
    {"item": "Pinch Roll", "qty": 1, "price": 1000000},
    {"item": "Hydraulic Powerpack Charging Grid", "qty": 1, "price": 2500000},
]

REFRACTORY_ITEMS = [
    {"item": "Hanging Brick 60% (Soaking & Heating)", "qty": 1207, "weight_kg": 47077, "rate": 1555, "cost": 1877041},
    {"item": "Hanging Brick 40% (Pre-Heating)", "qty": 319, "weight_kg": 11814, "rate": 1075, "cost": 343256},
    {"item": "Holding Brick 60% (Soaking & Heating)", "qty": 1028, "weight_kg": 28970, "rate": 945, "cost": 971460},
    {"item": "Holding Brick 60% (Pre-Heating)", "qty": 282, "weight_kg": 7025, "rate": 710, "cost": 200407},
    {"item": "Fire Brick 60% Bottom (Soaking/Heating)", "qty": 7237, "weight_kg": 5140, "rate": 105, "cost": 759868},
    {"item": "Fire Brick 50%+40% Bottom", "qty": 7081, "weight_kg": 39803, "rate": 160, "cost": 1132882},
    {"item": "Fire Brick IS 8 Bottom Soaking", "qty": 5947, "weight_kg": 28322, "rate": 20, "cost": 118947},
    {"item": "Hot Face", "qty": 11658, "weight_kg": 10705, "rate": 48, "cost": 559580},
    {"item": "Cold Face", "qty": 11658, "weight_kg": 20984, "rate": 20, "cost": 233158},
    {"item": "Side Arch Brick 60%", "qty": 750, "weight_kg": 52461, "rate": 105, "cost": 78750},
    {"item": "Hysil Board 900x600x50", "qty": 178, "weight_kg": 1485, "rate": 1230, "cost": 218774},
    {"item": "Castable for Roof (Insulite 4)", "qty": 7000, "weight_kg": 7000, "rate": 35, "cost": 245000},
    {"item": "Fibre for Roof", "qty": 50, "weight_kg": 700, "rate": 2200, "cost": 110000},
    {"item": "Aluminium Foil", "qty": 180, "weight_kg": 180, "rate": 350, "cost": 63000},
    {"item": "Refractory Block for Discharge (600x200x150)", "qty": 50, "weight_kg": 40500, "rate": 18000, "cost": 900000},
    {"item": "Skid Block LC-90 PCPF", "qty": 210, "weight_kg": 10206, "rate": 11000, "cost": 2310000},
    {"item": "LC-60 Castable on Hearth", "qty": 50000, "weight_kg": 50000, "rate": 100, "cost": 5000000},
]


def get_supplementary():
    """Return all calculation data for display in UI."""
    return {
        "furnace_calc": FURNACE_CALC,
        "mass_flow_control": MASS_FLOW_CONTROL,
        "recuperator": RECUPERATOR_CALC,
        "combustion_items": COMBUSTION_ITEMS,
        "refractory_items": REFRACTORY_ITEMS,
    }


def build_snsf_brf_df(
    include_ng_optional: bool = False,
    include_client_scope: bool = False,
) -> pd.DataFrame:
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

    main_cost = df.loc[~df["SECTION"].isin(["NG Optional", "Client Scope"]), "COST PRICE"].sum()
    main_sell = df.loc[~df["SECTION"].isin(["NG Optional", "Client Scope"]), "SELL PRICE"].sum()
    ng_cost = df.loc[df["SECTION"] == "NG Optional", "COST PRICE"].sum() if include_ng_optional else 0
    ng_sell = df.loc[df["SECTION"] == "NG Optional", "SELL PRICE"].sum() if include_ng_optional else 0
    client_cost = df.loc[df["SECTION"] == "Client Scope", "COST PRICE"].sum() if include_client_scope else 0
    client_sell = df.loc[df["SECTION"] == "Client Scope", "SELL PRICE"].sum() if include_client_scope else 0

    summary = {
        "main_cost": float(main_cost),
        "main_sell": float(main_sell),
        "ng_optional_cost": float(ng_cost),
        "ng_optional_sell": float(ng_sell),
        "client_scope_cost": float(client_cost),
        "client_scope_sell": float(client_sell),
        "grand_cost": float(main_cost + ng_cost + client_cost),
        "grand_sell": float(main_sell + ng_sell + client_sell),
    }

    return df, summary
