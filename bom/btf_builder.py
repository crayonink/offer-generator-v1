# bom/btf_builder.py
"""
BOM builder for Box Type Furnace (10 Ton Reheating Furnace).
Two combustion modes: ON/OFF and Mass Flow.
Includes full calculation breakdown matching legacy Excel.
"""

import pandas as pd


# ── Structure items (common to both modes) ──────────────────────────────────
STRUCTURE_ITEMS = [
    {"item": "MS Structure",                  "qty": 8700, "unit": "kg",    "rate": 65.0,      "cost": 565500},
    {"item": "Piping",                        "qty": 4000, "unit": "kg",    "rate": 65.0,      "cost": 260000},
    {"item": "Ceramic Fibre",                 "qty": 340,  "unit": "Rolls", "rate": 2200.0,    "cost": 748000},
    {"item": "Refractory (Walls and Bogie)",  "qty": 1,    "unit": "Set",   "rate": 1514080,   "cost": 1514080},
    {"item": "Bogie with Drives",             "qty": 1,    "unit": "Set",   "rate": 1100700,   "cost": 1100700},
    {"item": "Door with Drives",              "qty": 1,    "unit": "Set",   "rate": 831998,    "cost": 831998},
]

# ── Combustion ON/OFF (22 items) ────────────────────────────────────────────
COMBUSTION_ONOFF = [
    {"item": "Burner HV-G3A",                           "qty": 6,  "unit_price": 66100},
    {"item": "After Burner (3000 kW Regenerator)",       "qty": 1,  "unit_price": 480000},
    {"item": "Blower 40\", 10HP",                        "qty": 1,  "unit_price": 61667},
    {"item": "Gas Train 100 Nm3/hr",                     "qty": 1,  "unit_price": 110000},
    {"item": "R-Type TC with TT (Zone + Furnace)",       "qty": 6,  "unit_price": 44000},
    {"item": "DPT for Furnace Pressure Control",         "qty": 1,  "unit_price": 45000},
    {"item": "Motorized Damper",                         "qty": 1,  "unit_price": 200000},
    {"item": "Solenoid Valve (Gas Line)",                "qty": 7,  "unit_price": 6500},
    {"item": "Ball Valve (Gas Line)",                    "qty": 7,  "unit_price": 2200},
    {"item": "Pilot Skid (Air/Gas)",                     "qty": 7,  "unit_price": 50000},
    {"item": "Flowmeter (Gas Line)",                     "qty": 1,  "unit_price": 150000},
    {"item": "Pressure Switch Low (Air Line)",           "qty": 1,  "unit_price": 12000},
    {"item": "Pressure Gauge (Air Line)",                "qty": 1,  "unit_price": 4000},
    {"item": "Shut-off Valve (Air Line) 65 NB",          "qty": 7,  "unit_price": 47000},
    {"item": "Butterfly Valve (Air Line) 65 NB",         "qty": 7,  "unit_price": 3000},
    {"item": "PLC S7-1200 with HMI 7\"",                 "qty": 1,  "unit_price": 200000},
    {"item": "Recuperator",                              "qty": 1,  "unit_price": 215424},
    {"item": "Dilution Air Blower",                      "qty": 1,  "unit_price": 25000},
    {"item": "Control Panel",                            "qty": 1,  "unit_price": 200000},
    {"item": "Temperature Recorder (Chino)",             "qty": 1,  "unit_price": 150000},
    {"item": "Oxygen Analyzer",                          "qty": 1,  "unit_price": 250000},
    {"item": "Cablings",                                 "qty": 1,  "unit_price": 100000},
]

# ── Combustion Mass Flow (28 items — adds zone-wise control) ────────────────
COMBUSTION_MASSFLOW = [
    {"item": "Burner HVG-3A",                           "qty": 6,  "unit_price": 66100},
    {"item": "After Burner (3000 kW Regenerator)",       "qty": 1,  "unit_price": 480000},
    {"item": "Blower 40\", 10HP",                        "qty": 1,  "unit_price": 61667},
    {"item": "Gas Train 100 Nm3/hr",                     "qty": 1,  "unit_price": 110000},
    {"item": "R-Type TC with TT (Zone + Furnace)",       "qty": 6,  "unit_price": 44000},
    {"item": "DPT for Furnace Pressure Control",         "qty": 1,  "unit_price": 45000},
    {"item": "Motorized Damper",                         "qty": 1,  "unit_price": 200000},
    {"item": "Solenoid Valve (Gas Line)",                "qty": 7,  "unit_price": 6500},
    {"item": "Ball Valve (Gas Line)",                    "qty": 7,  "unit_price": 2200},
    {"item": "Orifice Plate (Gas Line) 25 NB",           "qty": 2,  "unit_price": 3000},
    {"item": "DPT (Gas Line per Zone)",                  "qty": 2,  "unit_price": 45000},
    {"item": "Control Valve (Gas Line) 25 NB",           "qty": 2,  "unit_price": 83000},
    {"item": "Pilot Skid (Air/Gas)",                     "qty": 7,  "unit_price": 50000},
    {"item": "Flowmeter (Gas Line)",                     "qty": 1,  "unit_price": 150000},
    {"item": "Pressure Switch Low (Air Line)",           "qty": 1,  "unit_price": 12000},
    {"item": "Pressure Gauge (Air Line)",                "qty": 1,  "unit_price": 4000},
    {"item": "Shut-off Valve (Air Line) 65 NB",          "qty": 7,  "unit_price": 47000},
    {"item": "Butterfly Valve (Air Line) 65 NB",         "qty": 7,  "unit_price": 3000},
    {"item": "Orifice Plate (Air Line) 100 NB",          "qty": 2,  "unit_price": 7000},
    {"item": "DPT (Air Line per Zone)",                  "qty": 2,  "unit_price": 45000},
    {"item": "Control Valve (Air Line) 100 NB",          "qty": 2,  "unit_price": 110450},
    {"item": "Recuperator",                              "qty": 1,  "unit_price": 215424},
    {"item": "Dilution Air Blower",                      "qty": 1,  "unit_price": 25000},
    {"item": "PLC S7-1200 with HMI 7\"",                 "qty": 1,  "unit_price": 200000},
    {"item": "Control Panel",                            "qty": 1,  "unit_price": 200000},
    {"item": "Oxygen Analyzer",                          "qty": 1,  "unit_price": 250000},
    {"item": "Temperature Recorder (Chino)",             "qty": 1,  "unit_price": 150000},
    {"item": "Cablings",                                 "qty": 1,  "unit_price": 120000},
]


# ── Supplementary calculation data (from legacy Excel) ──────────────────────

FURNACE_DIMENSIONS = {
    "internal_L_mm": 2590, "internal_W_mm": 2660, "internal_H_mm": 1550,
    "ceramic_fibre_L_mm": 7300, "ceramic_fibre_W_mm": 600, "ceramic_fibre_thk_mm": 25,
    "total_ceramic_rolls": 340,
    "ms_sheet_vol_m3": 0.0756, "ms_density_kg_m3": 7850, "ms_sheet_wt_kg": 593.46,
    "ms_sheets_reqd": 3, "ms_wt_total_kg": 2640,
}

HEAT_LOAD_CALC = [
    {"item": "Heat to Charge",      "value": 158400, "unit": "kcal"},
    {"item": "Heat to Pier",        "value": 19800,  "unit": "kcal"},
    {"item": "Heat to Refractory",  "value": 72000,  "unit": "kcal"},
    {"item": "Heat to Insulation",  "value": 18480,  "unit": "kcal"},
    {"item": "Surface Loss",        "value": 10000,  "unit": "kcal"},
    {"item": "To Casting",          "value": 4125,   "unit": "kcal"},
    {"item": "Total",               "value": 282805, "unit": "kcal"},
    {"item": "Gross",               "value": 377073, "unit": "kcal"},
]

FURNACE_PARAMS = {
    "furnace_capacity_tph": 2.0,
    "total_load_tonne": 10.0,
    "std_gas_consumption_nm3_ton": 40.0,
    "fuel_consumption_nm3hr": 80,
    "air_flow_nm3hr": 840,
    "cfm": 494.12,
    "blower_hp_calc": 6.18,
    "blower_hp_selected": 10,
    "no_of_burners": 6,
    "no_of_zones": 2,
    "heating_zone_1_burners": 3,
    "heating_zone_2_burners": 3,
    "rating_per_zone_kcal": 188537,
    "rating_per_zone_kw": 219.23,
    "std_burner_rating_lph": 40.0,
}

PIPE_SIZING = {
    "air_zone1": {"flow_nm3hr": 166.67, "velocity_ms": 15.0, "inner_dia_mm": 62.70},
    "air_zone2": {"flow_nm3hr": 500.0,  "velocity_ms": 15.0, "inner_dia_mm": 108.61},
    "gas_zone1": {"flow_nm3hr": 15.0,   "velocity_ms": 13.0, "inner_dia_mm": 20.21},
    "gas_zone2": {"flow_nm3hr": 25.0,   "velocity_ms": 13.0, "inner_dia_mm": 26.09},
}

RECUPERATOR_CALC = {
    "total_flue_gas_nm3hr": 920,
    "total_mass_flue_gas_kghr": 1104,
    "specific_heat_flue_gas": 0.23,
    "initial_temp_flue_gas_c": 600.0,
    "final_temp_flue_gas_c": 389.19,
    "heat_transfer_coeff": 30.0,
    "air_volume_nm3hr": 840,
    "initial_air_temp_c": 35.0,
    "final_air_temp_c": 250.0,
    "specific_heat_air": 0.247,
    "heat_required_kcal": 53530,
    "lmtd_c": 316.88,
    "surface_area_m2": 5.63,
    "bank_length_mm": 526.4,
    "bank_width_mm": 535.5,
    "bank_gap_mm": 150.0,
    "pipes_total": 120,
    "pipes_per_row": 12,
    "pipes_per_column": 10,
    "pipe_dia_mm": 33.4,
    "pipe_length_m": 0.5,
    "pipe_thickness_mm": 4.5,
    "pipe_weight_kg_m": 3.24,
    "hot_bank_weight_kg": 97.2,
}


def get_supplementary():
    """Return all calculation data for display in UI."""
    return {
        "furnace_dimensions": FURNACE_DIMENSIONS,
        "heat_load": HEAT_LOAD_CALC,
        "furnace_params": FURNACE_PARAMS,
        "pipe_sizing": PIPE_SIZING,
        "recuperator": RECUPERATOR_CALC,
    }


def build_btf_df(combustion_mode: str = "onoff", markup: float = 1.8) -> pd.DataFrame:
    """
    Build BOM for Box Type Furnace.
    combustion_mode: 'onoff' or 'massflow'
    markup: sell price multiplier (default 1.8)
    """
    rows = []

    # Structure items
    for s in STRUCTURE_ITEMS:
        rows.append(("Structure", s["item"], s["qty"], s["unit"], s["rate"], s["cost"]))

    # Combustion system items
    comb_items = COMBUSTION_ONOFF if combustion_mode == "onoff" else COMBUSTION_MASSFLOW
    for c in comb_items:
        total = c["qty"] * c["unit_price"]
        rows.append(("Combustion System", c["item"], c["qty"], "", c["unit_price"], total))

    df = pd.DataFrame(rows, columns=["SECTION", "ITEM", "QTY", "UNIT", "UNIT PRICE", "COST PRICE"])

    # Add sell price columns
    df["SELL PRICE"] = (df["COST PRICE"] * markup).round(0)

    # Summary
    structure_cost = df.loc[df["SECTION"] == "Structure", "COST PRICE"].sum()
    combustion_cost = df.loc[df["SECTION"] == "Combustion System", "COST PRICE"].sum()
    total_cost = structure_cost + combustion_cost
    total_sell = round(total_cost * markup, 0)
    designing = round(total_sell * 0.10, 0)
    negotiation = round(total_sell * 0.10, 0)
    quoted = round(total_sell + designing + negotiation, -3)

    summary = {
        "structure_cost": float(structure_cost),
        "combustion_cost": float(combustion_cost),
        "total_cost": float(total_cost),
        "sell_price": float(total_sell),
        "designing_10pct": float(designing),
        "negotiation_10pct": float(negotiation),
        "quoted_price": float(quoted),
        "markup": markup,
        "combustion_mode": combustion_mode,
    }

    return df, summary
