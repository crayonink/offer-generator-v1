# bom/snsf_brf_builder.py
"""
BOM builder for SNSF BRF (Billet Reheating Furnace, 30 Ton).
Per-item markup multipliers as per the costing breakup sheet.
All prices read from snsf_brf_price_master DB table.
"""

import pandas as pd
import sqlite3
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.path.join(BASE_DIR, "vlph.db")


def _load_snsf_items(include_ng=False, include_client=False) -> list:
    """Load items from snsf_brf_price_master."""
    conn = sqlite3.connect(DB_PATH)
    exclude = []
    if not include_ng:
        exclude.append("NG Optional")
    if not include_client:
        exclude.append("Client Scope")

    if exclude:
        placeholders = ",".join("?" * len(exclude))
        rows = conn.execute(
            f"SELECT category, item, qty, unit, unit_price, markup FROM snsf_brf_price_master WHERE category NOT IN ({placeholders}) ORDER BY rowid",
            exclude
        ).fetchall()
    else:
        rows = conn.execute(
            "SELECT category, item, qty, unit, unit_price, markup FROM snsf_brf_price_master ORDER BY rowid"
        ).fetchall()

    if include_ng:
        ng_rows = conn.execute(
            "SELECT category, item, qty, unit, unit_price, markup FROM snsf_brf_price_master WHERE category='NG Optional' ORDER BY rowid"
        ).fetchall()
        rows = list(rows) + list(ng_rows)
    if include_client:
        cl_rows = conn.execute(
            "SELECT category, item, qty, unit, unit_price, markup FROM snsf_brf_price_master WHERE category='Client Scope' ORDER BY rowid"
        ).fetchall()
        rows = list(rows) + list(cl_rows)

    conn.close()
    return [{"section": r[0], "item": r[1], "qty": r[2], "unit": r[3] or "", "unit_price": r[4], "markup": r[5]} for r in rows]


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

# Full furnace calculation data — every row from legacy "Furnace (2)" sheet
FURNACE_FULL_DATA = [
    ['REHEATING FURNACE CALCULATION'],
    ['Fluid Data', '', '', '', 'Technical Data'],
    ['Gas =', '', '', '', 'Max  Reheat Temp. (T2) =', 0, '°C'],
    ['Pressure =', '', 'bar', '', 'Time =', 1, 'hours'],
    ['Calorific Value =', 9500, 'Kcal/Nm³', '', 'weight of the billet =', 471000, 'kg'],
    ['Gas Temp. =', '', '°C', '', 'From Ambient to T2 =', 100, '°C'],
    ['1. BILLET CALCULATION & DIMENSION OF FURNACE'],
    ['Billet Dimension in metre'],
    ['Length (mm) =', 4000, 4, 'm', 'DATA FILLING ='],
    ['Width (mm) =', 100, 150, 'm', 'OUTPUT ='],
    ['Height (mm) =', 100, 0.1, 'm'],
    ['Volume of Billet', 60, 'm³'],
    ['Density of MS =', 7850, 'kg/m³'],
    ['Wt. of a billet =', 471000, 'kg'],
    ['For Furnace Dimension Calculation 3 times of the material is inside the furnace'],
    ['Furnace Cap. =', 30, 'Tonne/hr'],
    ['Inside Material =', 90000, 'Kg'],
    ['No. of Billets inside the Furnace =', 0.1911, 1],
    ['Effective Length of the Furnace (Centre Distance)'],
    ['EFFECTIVE LENGTH =', '', 0.1, 'metre'],
    ['EFFECTIVE WIDTH =', '', 4.5, 'metre'],
    ['', 'Overall Width with Refractory', '', '', 'Overall Length with Refractory'],
    ['115+115+230+50=510', 'EFFECTIVE WIDTH =', 4500, 'mm', 'EFFECTIVE LENGTH =', 17000, 'mm'],
    ['460 + 50 = 510', 'Right Refractory =', 510, 'mm', 'Discharge side with Refractory =', 2010, 'mm'],
    ['460 Refractory Lining', 'Left Refractory =', 510, 'mm', 'Charge side with Refractory =', 1500, 'mm'],
    ['50 HAY', 'Sheet =', 16, 'mm', 'Sheet on charging side =', 6, 'mm'],
    ['', 'Channel =', 500, 'mm', 'Sheet on refractory side =', 6, 'mm'],
    ['', 'Overall Width =', 6036, 'mm', 'Channel =', 500, 'mm'],
    ['', '', 6.036, 'metre', 'Overall Length =', 21.022, 'metre'],
    ['PREHEATING ZONE', '', 'HEATING ZONE', '', '', 'SOAKING ZONE'],
    ['Length(m) =', 4.2044, 'Length (metre) =', '', 10.231, 'Length(m) =', 6.069],
    ['2. ROOF REFRACTORY & STRUCTURE CALCULATION'],
    ['ROOF REFRACTORY', '', '', '', 'SIDE WALL REFRACTORY (R&L)'],
    ['Hanging and holding Brick (High Alumina)', '', '', '', '50mm HAY + Cold face + hot Face + 60% Fire Brick'],
    ['Hanging Brick Per width', 6.6133, 7],
    ['Total Hanging Brick (Soaking+Heating 60% H/A)', 1207.1, 'Nos'],
    ['Wt. of Each Hanging Brick (60%H/A)', 39.36, 'kg'],
    ['Total Hanging Brick in Preheating (40% H/A)', 319.308, 'Nos'],
    ['Wt. of Each Hanging Brick (40%H/A)', 36.85, 'kg'],
    ['Total Hanging Brick Wt.', 59278, 'kg'],
    ['Holding Brick per width', 6, 'Nos'],
    ['Total Holding Brick (Soaking+Heating 60%)', 1028, 'Nos'],
    ['Wt. of Each Holding Brick (60%H/A)', 23.6, 'kg'],
    ['Total Wt. of Holding Brick (60%H/A)', 24261, 'kg'],
    ['Total Holding Brick in Preheating (40%)', 282.264, 'Nos'],
    ['Wt. of Each Holding brick (40%)', 22.1, 'Kg'],
    ['Total Wt. of Holding Brick (40%H/A)', 6238, 'kg'],
    ['STD. Ceramic Fiber', '(7300 x 600 x 25)', 'mm'],
    ['Wt. of Each Fiber', 15, 'kg'],
    ['Total Ceramic Fibre rolls (roof)', 22.48, 'Rolls'],
    ['Total Rolls', 33, 'Rolls'],
    ['Total Ceramic Weight', 495, 'kg'],
    ['Castable volume', 1.3332, 'm³'],
    ['Castable Wt.', 733.23, 'Kg'],
    ['Total Castable wt.', 6452.45, 'Kg'],
    ['Standard Bag wt.', 25, 'kg'],
    ['Total Bag Required', 258.1, 'Nos'],
    ['Total Al. Foil Required', 133.06, 'metre'],
    ['', 107.78, 'kg'],
    ['FLUE DUCT CALCULATION Before Recuperator1'],
    ['Internal Dia. =', 1.9, 'metre'],
    ['Outer Dia. =', 2.2, 'metre'],
    ['Length of the Duct =', 20, 'metre'],
    ['Volume of Castable =', 19.311, 'm3'],
    ['Castable (LC40) Wt. =', 48857, 'kg'],
    ['Total 25 kg Bag Required', 1954.27, 'Nos'],
    ['FLUE DUCT CALCULATION After Recuperator 2'],
    ['Internal Dia. =', 1.9, 'metre'],
    ['Outer Dia. =', 2.1, 'metre'],
    ['Length of the Duct =', 7.5, 'metre'],
    ['Volume of Castable =', 4.71, 'm3'],
    ['Castable (LC40) Wt. =', 11916.3, 'kg'],
    ['Total 25 kg Bag Required', 476.65, 'Nos'],
    ['STRUCTURAL CALCULATION'],
    ['STRUCTURAL CALCULATION ON SIDE WALL', '', '', '', 'I-BEAM CALCULATION (500×180)'],
    ['C-CHANNEL std. size =', '250*80*7.2 thk', 'mm', '', 'STD. Size of I-Beam =', '500*180*', 'mm'],
    ['Wt. per metre =', 30.6, 'kg', '', 'Wt. per metre =', 87, 'Kg'],
    ['Single Size Required =', 3, 'metre', '', 'Qty Required on Roof =', 14.01, 'Nos'],
    ['Overall Length (4 side) =', 82.04, 'metre', '', 'Required Width =', 6.036, 'metre'],
    ['Overall Width (7 side) =', 42.252, 'metre', '', 'Total 500×180 beam =', 84.59, 'Metre'],
    ['3m Pcs Channel b/w outers =', 156, 'metre', '', 'Total Wt. =', 8096, 'Kg'],
    ['Total C-channel =', 283.29, 'metre', '', 'I-BEAM (200×100) for Top & Bottom'],
    ['Total Wt. of C-channel =', 9536, 'kg', '', 'Wt. per metre =', 24.5, 'Kg'],
    ['CHARGING DOOR SIDE', '', '', '', 'Qty Required on Roof =', 7, 'Nos'],
    ['volume of Flat Size =', 0.0481, 'm3', '', 'Required Length =', 21.022, 'metre'],
    ['Wt. of the Flat =', 377.93, 'Kg', '', '200×100 beam for Top =', 147.15, 'Metre'],
    ['16mm Plate for Door Sliding =', '(7x1x0.016)', 'Metre', '', 'Total Wt. =', 3965.8, 'Kg'],
    ['Wt. of the Plate =', 692.18, 'Kg', '', 'Qty Required on Bottom =', 14.01, 'Nos'],
    ['Plate For Bottom Tie Rod (8mm)', '', '', '', '200×100 for Bottom =', 294.62, 'Metre'],
    ['Std. Size =', '(8036x250x8)', 'mm', '', 'Total 200×100 I-Beam =', 441.77, 'Metre'],
    ['Wt. of Plate =', 1460.91, 'kg', '', 'Total 200×100 I-Beam Wt. =', 12302, 'Kg'],
    ['CASTING MATERIAL CALCULATION'],
    ['CI HANGER FOR HANGING BRICK', '', '', '', 'CI SKID FOR PREHEATING ZONE'],
    ['Qty. Required =', 1451.4, 'Nos', '', 'Std. length of skid =', 1.5, 'Metre'],
    ['Wt. of single piece =', 3, 'kg', '', 'use of 4 lines skid =', 4.27, 'Nos'],
    ['Total Wt. of Hanger =', 4354.22, 'Kg', '', 'Total Skid =', 17.08, 'Nos'],
    ['', '', '', '', 'Wt. of each skid =', 250, 'kg'],
    ['Total Weight of Casting =', 10.72, 'Tonne', '', 'Total Skid weight =', 4269.6, 'kg'],
    ['PIPE CALCULATION'],
    ['Pipe Required For Flue Duct (5mm thk.)', '', '', '', 'PIPE DIA. FOR COLD AIR'],
    ['Internal Dia. =', 2.19, 'metre', '', 'Flow =', 400, 'Nm³/hr'],
    ['Outer Dia. =', 2.2, 'metre', '', 'Velocity =', 15, 'm/s'],
    ['Length of Pipe =', 27.5, 'metre', '', 'Inner Dia. =', 283.21, 'mm'],
    ['Volume of 5mm pipe =', 0.9477, 'm3', '', 'PIPE DIA. FOR HOT AIR (After Recuperator)'],
    ['', '', '', '', 'Temp. After Recuperator =', 300, '°C'],
    ['', '', '', '', 'Density =', 1.9556, 'Kg/m³'],
    ['PIPE DIA. FOR BF GAS', '', '', '', 'Flow =', 782.25, 'Nm³/hr'],
    ['Flow =', 650, 'Nm³/hr', '', 'Velocity =', 12, 'm/s'],
    ['Velocity for BF Gas =', 12, 'm/s', '', 'Inner Dia. =', 151.88, 'mm'],
    ['Inner Dia. =', 138.45, 'mm'],
]

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

MILD_STEEL_ITEMS = [
    {"sno": 2, "item": "ISMB 500×180 (Roof)", "qty_m": 130, "wt_per_m": 87, "total_wt": 11310, "rate": 55, "cost": 622050},
    {"sno": 3, "item": "ISMB 200×100 (Top & Bottom)", "qty_m": 630, "wt_per_m": 24.5, "total_wt": 15435, "rate": 55, "cost": 848925},
    {"sno": 4, "item": "ISMC 250×80", "qty_m": 310, "wt_per_m": 30.6, "total_wt": 9486, "rate": 55, "cost": 521730},
    {"sno": 5, "item": "MS Plate 16mm flat (Charging Door)", "qty_m": 1, "wt_per_m": 970, "total_wt": 970, "rate": 55, "cost": 53350},
    {"sno": 6, "item": "MS Plate 8mm (Bottom Tie Rod)", "qty_m": 1, "wt_per_m": 2250, "total_wt": 2250, "rate": 55, "cost": 123750},
    {"sno": 7, "item": "MS Plate 12mm (join I-Beam)", "qty_m": 4.7, "wt_per_m": 350, "total_wt": 1645, "rate": 55, "cost": 90475},
    {"sno": 8, "item": "MS Plate 8mm (Side Walls)", "qty_m": 8, "wt_per_m": 4550, "total_wt": 36400, "rate": 55, "cost": 2002000},
    {"sno": 9, "item": "MS Plate 8mm (Discharge Side)", "qty_m": 1, "wt_per_m": 1550, "total_wt": 1550, "rate": 55, "cost": 85250},
    {"sno": 10, "item": "Duct Flue Gas 16+6=20 mtr", "qty_m": 1, "wt_per_m": 8300, "total_wt": 8300, "rate": 55, "cost": 456500},
    {"sno": 11, "item": "Chimney per Tonne", "qty_m": 1, "wt_per_m": 16000, "total_wt": 16000, "rate": 55, "cost": 880000},
    {"sno": 12, "item": "Supporting Structure", "qty_m": 1, "wt_per_m": 4500, "total_wt": 4500, "rate": 55, "cost": 247500},
]
MILD_STEEL_TOTAL_WT = 107846  # kg = 107.846 Ton
MILD_STEEL_TOTAL_COST = 5931530

PIPELINE_CASTING = [
    {"item": "Air, Gas, Oil Pipeline", "wt": 8000, "rate": 75, "cost": 600000},
    {"item": "Inspection Door, Discharge Door, Plate", "wt": 3000, "rate": 150, "cost": 450000},
    {"item": "CI Skid for Preheating Zone", "qty": 16, "wt_each": 250, "wt": 4000, "rate": 150, "cost": 600000},
    {"item": "CI Hanger 2400 PC (3 kg per piece)", "wt": 7300, "rate": 150, "cost": 1095000},
]

REFRACTORY_EXTRA = [
    {"item": "Flue Duct Castable (LC-40) 2.2m OD×1.9m ID", "qty": 49000, "wt": 49000, "rate": 35, "cost": 1715000},
    {"item": "Side Arch Brick IS-8 (Chimney)", "qty": 2000, "wt": 8000, "rate": 20, "cost": 40000},
    {"item": "Cold Face (Chimney)", "qty": 5000, "wt": 4000, "rate": 20, "cost": 100000},
    {"item": "Fireclay (bags)", "qty": 170, "wt": 250000, "rate": 765, "cost": 130050},
    {"item": "Accosset 50 (bags)", "qty": 165, "wt": 8500, "rate": 865, "cost": 142725},
]

RECUPERATOR_COST_BREAKDOWN = {
    "cost_ss316_pipe_rate": 470.0,
    "cost_all_pipes": 1226057,
    "ms_outershell_kg": 296.60,
    "ms_air_inlet_duct_kg": 115.25,
    "ms_hot_air_outlet_duct_kg": 127.00,
    "ms_pipe_holding_plate_kg": 655.25,
    "ms_bottom_box_kg": 176.25,
    "ms_flanges_kg": 45,
    "ms_side_hood_kg": 1600,
    "cost_ms_outer_shell": 10000,
    "cost_ms_combustion_air_inlet": 180921,
    "cost_ms_channel_150x75": 10200,
    "cost_angle_65": 13200,
    "cost_angle_75": 5400,
    "cost_angle_50": 4050,
    "cost_total_material": 1449828,
    "cost_pipe_bending": 10500,
    "cost_welding_rod": 6912,
    "cost_hole_fabrication": 43200,
    "cost_thermocouple_tt": 8000,
    "cost_total_recuperator": 1518440,
    # Cold bank
    "cold_bank_material": "SS-316 Seamless",
    "cold_bank_dia_mm": 48.3,
    "cold_bank_length_m": 2.575,
    "cold_bank_thickness_mm": 2.7,
    "cold_bank_weight_kg_m": 4.72,
    "cold_bank_pipe_wt_kg": 12.0,
    "cold_bank_total_wt_kg": 1296,
}

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
        "furnace_full": FURNACE_FULL_DATA,
        "mild_steel": MILD_STEEL_ITEMS,
        "mild_steel_total_wt": MILD_STEEL_TOTAL_WT,
        "mild_steel_total_cost": MILD_STEEL_TOTAL_COST,
        "pipeline_casting": PIPELINE_CASTING,
        "mass_flow_control": MASS_FLOW_CONTROL,
        "recuperator": RECUPERATOR_CALC,
        "recuperator_cost": RECUPERATOR_COST_BREAKDOWN,
        "combustion_items": COMBUSTION_ITEMS,
        "refractory_items": REFRACTORY_ITEMS,
        "refractory_extra": REFRACTORY_EXTRA,
    }


def build_snsf_brf_df(
    include_ng_optional: bool = False,
    include_client_scope: bool = False,
) -> pd.DataFrame:
    # Load from DB, fallback to hardcoded
    db_items = _load_snsf_items(include_ng=include_ng_optional, include_client=include_client_scope)
    if db_items:
        items = db_items
    else:
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
