# bom/btf_builder.py
"""
BOM builder for Box Type Furnace (10 Ton Reheating Furnace).
Two combustion modes: ON/OFF and Mass Flow.
Includes full calculation breakdown matching legacy Excel.
All prices read from btf_price_master DB table.
"""

import pandas as pd
import sqlite3
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.path.join(BASE_DIR, "vlph.db")


def _load_btf_items(category: str) -> list:
    """Load items from btf_price_master for a given category."""
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute(
        "SELECT item, qty, unit, rate FROM btf_price_master WHERE category=? ORDER BY rowid",
        (category,)
    ).fetchall()
    conn.close()
    return [{"item": r[0], "qty": r[1], "unit": r[2] or "", "rate": r[3]} for r in rows]


# Fallback hardcoded data (used only if DB is empty)
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
    # Cold Bank
    "cold_bank_material": "CS Boiler Grade",
    "cold_bank_dia_mm": 33.4,
    "cold_bank_length_m": 0.5,
    "cold_bank_thickness_mm": 4.5,
    "cold_bank_weight_kg_m": 3.24,
    "cold_bank_pipe_wt_kg": 1.62,
    "cold_bank_total_wt_kg": 97.2,
    # Cost breakdown
    "cost_hot_bank_pipe_rate": 134.80,
    "cost_cold_bank_pipe_rate": 134.80,
    "cost_ms_rate": 60.0,
    "cost_all_pipes": 26204,
    "ms_outershell_kg": 49.86,
    "ms_air_inlet_duct_kg": 102.36,
    "ms_hot_air_outlet_duct_kg": 112.79,
    "ms_pipe_holding_plate_kg": 356.57,
    "ms_bottom_box_kg": 125.58,
    "ms_machining_flanges_kg": 100,
    "ms_side_hood_kg": 700,
    "cost_ms_outer_shell": 20000,
    "cost_ms_combustion_air_inlet": 92830,
    "cost_ms_channel_150x75": 10200,
    "cost_angle_65": 13200,
    "cost_angle_75": 5400,
    "cost_angle_50": 4050,
    "cost_total_material": 171884,
    "cost_pipe_bending": 7700,
    "cost_welding_rod": 3840,
    "cost_hole_fabrication": 24000,
    "cost_thermocouple_tt": 8000,
    "cost_total_recuperator": 215424,
}

FURNACE_VOLUMES = [
    {"item": "Vol. of Side wall (Right)", "value": 0.047212, "unit": "m3"},
    {"item": "Vol. of Side wall (Left)", "value": 0.047212, "unit": "m3"},
    {"item": "Vol. of Back Side", "value": 0.048248, "unit": "m3"},
    {"item": "Vol. of Top Side", "value": 0.083195, "unit": "m3"},
]

GAS_CONSUMPTION_CALC = [
    {"item": "No. of Zones", "value": "2"},
    {"item": "Rating / Zone", "value": "1,88,537 kcal (219.23 kW)"},
    {"item": "Gross Heat Load", "value": "3,77,073 kcal"},
    {"item": "Gross (kW)", "value": "438.46 kW"},
    {"item": "Heating Zone Ratio", "value": "25:8"},
    {"item": "Consumption (Heating)", "value": "2,82,805 kcal"},
    {"item": "Consumption (Soaking)", "value": "70,701 kcal"},
    {"item": "GROSS CON. (Regen)", "value": "70,70,125 kcal → 10,180,980 kcal"},
    {"item": "GROSS CON. (Conv.)", "value": "5,65,610 kcal → 13,883,155 kcal"},
    {"item": "CV of Gas", "value": "8,600 kcal/m3"},
    {"item": "Gas Required (Regen)", "value": "1,183.83 m3"},
    {"item": "Gas Required (Conv.)", "value": "1,614.32 m3"},
    {"item": "Gas / MT (Regen)", "value": "118.38 m3/MT"},
    {"item": "Gas / MT (Conv.)", "value": "161.43 m3/MT"},
    {"item": "KG/MT (Regen)", "value": "87.60 kg/MT"},
    {"item": "KG/MT (Conv.)", "value": "119.46 kg/MT"},
    {"item": "Guarantee (Regen)", "value": "96.36 kg/MT (+10%)"},
    {"item": "Guarantee (Conv.)", "value": "131.41 kg/MT (+10%)"},
]

GEAR_BOX_CALC = [
    {"item": "Total Load", "value": "10 Tonne"},
    {"item": "Max Required Speed", "value": "20 m/min"},
    {"item": "Fractional Resistance", "value": "0.03"},
    {"item": "Efficiency", "value": "0.5"},
    {"item": "Pull Load", "value": "300 Kgs"},
    {"item": "Required Power", "value": "2.67 kW"},
    {"item": "Required HP", "value": "3.57 HP"},
    {"item": "10% Extra", "value": "3.93 HP"},
]


MS_STRUCTURE_BREAKDOWN = [
    {"section": "MS Sheet", "items": [
        {"item": "Std. sheet Vol. (1500×6300×8)mm", "value": "0.0756 m3"},
        {"item": "Density of MS Sheet", "value": "7850 kg/m3"},
        {"item": "Wt. of each sheet", "value": "593.46 kg"},
        {"item": "No. of sheets Required", "value": "3 Nos"},
        {"item": "Wt. of Required MS sheet", "value": "2,640 Kg", "highlight": True},
    ]},
    {"section": "C-Channel 150×75×7.2 mm", "items": [
        {"item": "Wt. Per metre", "value": "17 kg"},
        {"item": "Outer walls × 3", "value": "90 m"},
        {"item": "Back and Front", "value": "30 m"},
        {"item": "Wt. of C-Channel", "value": "2,500 kg", "highlight": True},
    ]},
    {"section": "I-Beam 500×180 mm", "items": [
        {"item": "Wt. Per metre", "value": "87 kg"},
        {"item": "Total I-Beam", "value": "33 m"},
        {"item": "Wt. of I-Beam", "value": "3,500 kg", "highlight": True},
    ]},
    {"section": "TOTAL", "items": [
        {"item": "Total MS Structure", "value": "8,640 kg", "highlight": True},
    ]},
]

DOOR_BREAKDOWN = [
    {"item": "Plate (2.6×2.7)m thk 8mm", "qty": 593.46, "rate": 60, "total": 35608},
    {"item": "Angle (65×65×8)", "qty": 407, "rate": 60, "total": 24420},
    {"item": "Angle (40×40×6)", "qty": 297, "rate": 60, "total": 17820},
    {"item": "CI Casting HK for Door", "qty": 495, "rate": 650, "total": 321750},
    {"item": "Door Guide Plate", "qty": 165, "rate": 60, "total": 9900},
    {"item": "Counter Weight", "qty": 880, "rate": 60, "total": 52800},
    {"item": "Door Locking Cylinder", "qty": 2, "rate": 25000, "total": 50000},
    {"item": "Door Column (L,R,Top) 3 nos", "qty": 500, "rate": 62, "total": 31000},
    {"item": "Shafts", "qty": 200, "rate": 110, "total": 22000},
    {"item": "Chain 5 mtr", "qty": 5, "rate": 5000, "total": 25000},
    {"item": "Sprocket 4 Nos (20 teeth)", "qty": 4, "rate": 3500, "total": 14000},
    {"item": "Sprocket 6 Nos (15 teeth)", "qty": 6, "rate": 2000, "total": 12000},
    {"item": "Rope 22 mtr", "qty": 22, "rate": 100, "total": 2200},
    {"item": "Pulley", "qty": 50, "rate": 110, "total": 5500},
    {"item": "Guide Roller", "qty": 4, "rate": 2000, "total": 8000},
    {"item": "Gear Box Motor 5 HP (Nord)", "qty": 1, "rate": 200000, "total": 200000},
]

CERAMIC_FIBRE_CALC = [
    {"wall": "Side wall (Right)", "L": 3290, "W": 1000, "thk": 350, "area": 3.29},
    {"wall": "Side wall (Left)", "L": 3290, "W": 1000, "thk": 350, "area": 3.29},
    {"wall": "Top", "L": 3290, "W": 3360, "thk": 350, "area": 11.05},
    {"wall": "Back wall", "L": 2660, "W": 1000, "thk": 350, "area": 2.66},
    {"wall": "Door", "L": 3456, "W": 2066, "thk": 300, "area": 0},
]
CERAMIC_FIBRE_SUMMARY = {
    "vol_std_ceramic_m3": 0.1095,
    "total_vol_m3": 14.045,
    "total_rolls": 340,
    "wt_per_roll_kg": 15,
}

TROLLEY_BOGIE_BREAKDOWN = [
    {"item": "MS Beam 200×125 - 90 mtr", "qty": 1500, "rate": 60, "total": 90000},
    {"item": "MS Channel 200×75 - 60 mtr", "qty": 500, "rate": 60, "total": 30000},
    {"item": "MS Plate 10mm × 6300×1500", "qty": 450, "rate": 60, "total": 27000},
    {"item": "MS Round 85 DIA 2.5 mtr", "qty": 70, "rate": 60, "total": 4200},
    {"item": "Plummer Block Plate 0.3×0.3×0.09", "qty": 500, "rate": 60, "total": 30000},
    {"item": "MS Plate 30mm for Block Fixing", "qty": 150, "rate": 60, "total": 9000},
    {"item": "Rack", "qty": 5, "rate": 40000, "total": 200000},
    {"item": "Sprocket", "qty": 2, "rate": 50000, "total": 100000},
    {"item": "Shaft for Connecting Gearbox", "qty": 2, "rate": 15000, "total": 30000},
    {"item": "Bearing Assembly (12 Nos)", "qty": 12, "rate": 10000, "total": 120000},
    {"item": "Wheel (6 Nos)", "qty": 400, "rate": 110, "total": 44000},
    {"item": "CI Plate for Trolley", "qty": 800, "rate": 100, "total": 80000},
    {"item": "MS Angle 100×100×10 - 36 mtr", "qty": 200, "rate": 65, "total": 13000},
    {"item": "MS Round 45 DIA 20 mtr", "qty": 150, "rate": 70, "total": 10500},
    {"item": "Gear Box Motor 5 HP with Geared Motor", "qty": 1, "rate": 300000, "total": 300000},
    {"item": "Anchor/Stud 50 Kg", "qty": 50, "rate": 260, "total": 13000},
]

REFRACTORY_BREAKDOWN = {
    "walls": [
        {"type": "C/F Bricks", "qty": 368, "cost_per": 26, "total": 9568, "wt_per": 1.8, "total_wt": 46.8},
        {"type": "H/F Bricks", "qty": 368, "cost_per": 70, "total": 25760, "wt_per": 2.0, "total_wt": 140},
        {"type": "IS-8", "qty": 368, "cost_per": 90, "total": 33120, "wt_per": 3.8, "total_wt": 342},
        {"type": "HA 60%", "qty": 368, "cost_per": 90, "total": 33120, "wt_per": 5.0, "total_wt": 450},
    ],
    "bogie": [
        {"type": "C/F Bricks", "qty": 512, "cost_per": 26, "total": 13312, "wt_per": 1.8, "total_wt": 46.8},
        {"type": "H/F Bricks", "qty": 512, "cost_per": 70, "total": 35840, "wt_per": 2.0, "total_wt": 140},
        {"type": "IS-8", "qty": 512, "cost_per": 90, "total": 46080, "wt_per": 3.8, "total_wt": 342},
        {"type": "HA 60%", "qty": 512, "cost_per": 90, "total": 46080, "wt_per": 5.0, "total_wt": 450},
        {"type": "Hysil Board", "qty": 38, "cost_per": 1260, "total": 47880, "wt_per": 6.0, "total_wt": 228},
    ],
    "misc": [
        {"type": "Accoset 50", "qty": 12, "rate": 950, "total": 11400},
        {"type": "Fireclay", "qty": 12, "rate": 240, "total": 2880},
        {"type": "Firecreat", "qty": 10, "rate": 765, "total": 7650},
    ],
    "arch": [
        {"type": "End Arch Brick", "qty": 125, "cost_per": 90, "total": 11250},
        {"type": "Skew Block Brick", "qty": 12, "cost_per": 170, "total": 2040},
        {"type": "Side Arch Brick 01", "qty": 125, "cost_per": 1260, "total": 157500},
        {"type": "Side Arch Brick 02", "qty": 60, "cost_per": 90, "total": 5400},
    ],
}


def get_supplementary():
    """Return all calculation data for display in UI."""
    return {
        "furnace_dimensions": FURNACE_DIMENSIONS,
        "heat_load": HEAT_LOAD_CALC,
        "furnace_params": FURNACE_PARAMS,
        "pipe_sizing": PIPE_SIZING,
        "recuperator": RECUPERATOR_CALC,
        "furnace_volumes": FURNACE_VOLUMES,
        "gas_consumption": GAS_CONSUMPTION_CALC,
        "gear_box": GEAR_BOX_CALC,
        "ms_structure": MS_STRUCTURE_BREAKDOWN,
        "door": DOOR_BREAKDOWN,
        "ceramic_fibre": CERAMIC_FIBRE_CALC,
        "ceramic_fibre_summary": CERAMIC_FIBRE_SUMMARY,
        "trolley_bogie": TROLLEY_BOGIE_BREAKDOWN,
        "refractory": REFRACTORY_BREAKDOWN,
    }


def build_btf_df(combustion_mode: str = "onoff", markup: float = 1.8) -> pd.DataFrame:
    """
    Build BOM for Box Type Furnace.
    combustion_mode: 'onoff' or 'massflow'
    markup: sell price multiplier (default 1.8)
    """
    rows = []

    # Structure items — from DB, fallback to hardcoded
    db_structure = _load_btf_items("Structure")
    structure_data = db_structure if db_structure else STRUCTURE_ITEMS
    for s in structure_data:
        cost = round(s["qty"] * s["rate"], 0) if "cost" not in s else s["cost"]
        rows.append(("Structure", s["item"], s["qty"], s.get("unit", ""), s["rate"], cost))

    # Combustion system items — from DB, fallback to hardcoded
    comb_category = "Combustion-Massflow" if combustion_mode == "massflow" else "Combustion-ONOFF"
    db_comb = _load_btf_items(comb_category)
    if db_comb:
        for c in db_comb:
            total = round(c["qty"] * c["rate"], 0)
            rows.append(("Combustion System", c["item"], c["qty"], c.get("unit", ""), c["rate"], total))
    else:
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
