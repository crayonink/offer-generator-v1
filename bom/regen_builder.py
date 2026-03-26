"""
Regenerative Burner BOM builder.

Two categories:
  REGEN_BOM_PER_PAIR — qty scales with num_pairs  (e.g. 2 pairs → qty × 2)
  REGEN_BOM_FIXED    — always qty=1 regardless of pairs  (PLC, gas train, E&C)

build_regen_df(num_pairs, markup) → DataFrame with cost + selling columns
"""

import pandas as pd


# (section, item_name, specification, qty_per_pair, cost_per_unit ₹)
REGEN_BOM_PER_PAIR = [
    # ── BURNER SET ────────────────────────────────────────────────────────────
    ("BURNER SET",   "Regenerative Burner Set",     "1000 KW (2×500 KW burners) with refractory heat-storage media, complete",    1, 162998.63),

    # ── AIR LINE ─────────────────────────────────────────────────────────────
    ("AIR LINE",     "Combustion Air Blower",        "3 HP, SWSI centrifugal, with motor & base frame",                            1, 120000.00),
    ("AIR LINE",     "Regen (Exhaust) Blower",       "3 HP, SWSI centrifugal, with motor & base frame",                            1, 120000.00),
    ("AIR LINE",     "Air Butterfly Valve",           "DN80, SS body, pneumatic actuator operated",                                 4,   8500.00),
    ("AIR LINE",     "Air Pressure Gauge",            "0–500 mmWC, glycerine filled, 63 mm dial",                                   2,   2500.00),

    # ── GAS LINE ─────────────────────────────────────────────────────────────
    ("GAS LINE",     "Gas Solenoid Valve",            "DN25, 2-port NC, 230 VAC, rated for NG",                                     2,  18500.00),
    ("GAS LINE",     "Quick Exhaust Valve",            "DN25, pilot operated, SS body",                                              2,   9500.00),
    ("GAS LINE",     "Gas Butterfly Valve",            "DN40, SS body, actuator operated",                                           2,  12500.00),
    ("GAS LINE",     "Gas Pressure Regulator",         "DN25, adjustable 10–50 mbar, Itron or equivalent",                           1,  24000.00),
    ("GAS LINE",     "Gas Pressure Gauge",             "0–1 bar, glycerine filled, 63 mm dial",                                      2,   2500.00),
    ("GAS LINE",     "Gas Pressure Switch",            "0–500 mbar, SPDT, adjustable set-point",                                     2,   6500.00),

    # ── TEMP CONTROL ─────────────────────────────────────────────────────────
    ("TEMP CONTROL", "Type K Thermocouple",           "0–1400°C, mineral insulated, 3 m, SS sheath",                                2,   4500.00),
    ("TEMP CONTROL", "Temperature Controller (PID)", "Autonics TZ4ST or equivalent, 96×96 mm",                                    2,  12500.00),
    ("TEMP CONTROL", "Radiation Shield",              "SS316, for thermocouple hot-face protection",                                2,   4500.00),

    # ── COMBUSTION ───────────────────────────────────────────────────────────
    ("COMBUSTION",   "UV Flame Detector",             "Honeywell C7027A or equivalent",                                             2,  18500.00),
    ("COMBUSTION",   "Igniter + HV Transformer",      "Spark igniter with 10 kV transformer",                                       2,  22500.00),
    ("COMBUSTION",   "Proportional Control Valve",    "DN20, modulating, 0–10 V / 4–20 mA input",                                   1,  45000.00),
    ("COMBUSTION",   "Burner Switching Controller",   "Dedicated cycle-timer module, 24 VDC",                                        1,  65000.00),
]


# (section, item_name, specification, qty, cost_per_unit ₹)
REGEN_BOM_FIXED = [
    # ── CONTROLS ─────────────────────────────────────────────────────────────
    ("CONTROLS",    "PLC + HMI",                     "Siemens S7-1200 + KTP700 7\" touch panel",                                   1, 300000.00),
    ("CONTROLS",    "Main Control Panel",             "Powder-coated MS enclosure, MCBs, contactors, relays, 24 VDC SMPS",         1, 300000.00),
    ("CONTROLS",    "MCC / Motor Starters",           "DOL starters for all blower motors",                                         1,  85000.00),
    ("CONTROLS",    "Field Junction Box (FJB)",       "IP65, SS enclosure, with terminal strips",                                    1,  45000.00),
    ("CONTROLS",    "Cables & Conduits",              "Power + signal cables, conduit, glands, lugs (lump-sum estimate)",           1,  95000.00),

    # ── GAS TRAIN ────────────────────────────────────────────────────────────
    ("GAS TRAIN",   "Main Gas Train",                 "2\" NB, SS body, with filter, pressure gauge & SSOV, complete",              1, 110000.00),
    ("GAS TRAIN",   "Safety Shut Off Valve (SSOV)",   "DN50, spring-return NC, CE-marked",                                          1,  55000.00),
    ("GAS TRAIN",   "Gas Pressure Transmitter",       "0–500 mbar, 4–20 mA output, SS wetted",                                     2,  22000.00),

    # ── STRUCTURAL ───────────────────────────────────────────────────────────
    ("STRUCTURAL",  "Support Frame & Piping",         "MS structural steel, painted, with pipe support clamps",                     1, 120000.00),
    ("STRUCTURAL",  "Refractory Sealing Plugs",       "Cast refractory plugs for burner port sealing",                              1,  48000.00),

    # ── SERVICES ─────────────────────────────────────────────────────────────
    ("SERVICES",    "Erection & Commissioning",       "Site erection, wiring, commissioning, trial runs, handover",                 1, 225000.00),
]


def build_regen_df(num_pairs: int, markup: float = 1.80) -> pd.DataFrame:
    """Return a DataFrame with all BOM rows including cost and selling prices."""
    rows = []

    for section, item, spec, qty_per_pair, cost_unit in REGEN_BOM_PER_PAIR:
        qty = qty_per_pair * num_pairs
        rows.append(_make_row(section, item, spec, qty, cost_unit, markup))

    for section, item, spec, qty, cost_unit in REGEN_BOM_FIXED:
        rows.append(_make_row(section, item, spec, qty, cost_unit, markup))

    return pd.DataFrame(rows)


def _make_row(section, item, spec, qty, cost_unit, markup):
    total_cost  = qty * cost_unit
    sell_unit   = cost_unit * markup
    total_sell  = qty * sell_unit
    return {
        "SECTION":       section,
        "ITEM NAME":     item,
        "SPECIFICATION": spec,
        "QTY":           qty,
        "COST/UNIT":     round(cost_unit, 2),
        "TOTAL COST":    round(total_cost, 2),
        "SELL/UNIT":     round(sell_unit, 2),
        "TOTAL SELLING": round(total_sell, 2),
    }
