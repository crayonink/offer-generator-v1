"""
Regenerative Burner BOM builder — Excel-matched, KW-model-driven.

Each entry in REGEN_MODELS corresponds to 1 pair (2 burners) at that KW rating.
Valve sizes and costs are extracted directly from the legacy Excel costing sheets.

build_regen_df(kw, markup, num_pairs) → DataFrame with cost + selling columns
select_model(required_kw)             → nearest model KW >= required_kw
"""

import pandas as pd


# ── Per-model data (1 pair = 2 burners) ──────────────────────────────────────
# All values verified against "Copy of Regen Standard Costing_COG of DSP.xlsx"
# 6000 KW uses markup=2.0; all others use markup=1.8
REGEN_MODELS = {
    500: dict(
        burner_cost=124429.13, blower_cost=120000, markup=1.8,
        panel_cost=300000, gas_train_cost=88500,
        # Gas line — burner section
        gas_sol_nb=32, gas_sol_cost=13700,           # Solenoid valve per burner
        gas_bv_nb=32,  gas_bv_cost=4925,  gas_bv_qty=10,   # Ball valves
        gas_hose_nb=32, gas_hose_cost=1750, gas_hose_qty=10, # Flex hoses
        pg_burner=3000,
        # Air line — burner section
        air_sov_nb=125, air_sov_cost=50050,
        air_mbv_nb=125, air_mbv_cost=12498,
        flue_sov_nb=200, flue_sov_cost=80000,
        # Temperature control
        air_cv_nb=100, air_cv_cost=110450,
        air_fm_nb=125, air_fm_cost=54000,
        gas_cv_nb=25,  gas_cv_cost=83000,
        gas_fm_nb=32,  gas_fm_cost=48000,
        pneu_damp_nb=200, pneu_damp_cost=80000,
    ),
    1000: dict(
        burner_cost=162998.63, blower_cost=120000, markup=1.8,
        panel_cost=300000, gas_train_cost=110000,
        gas_sol_nb=40, gas_sol_cost=14720,
        gas_bv_nb=40,  gas_bv_cost=5100,  gas_bv_qty=10,
        gas_hose_nb=40, gas_hose_cost=2000, gas_hose_qty=10,
        pg_burner=3000,
        air_sov_nb=200, air_sov_cost=80000,
        air_mbv_nb=200, air_mbv_cost=31178,
        flue_sov_nb=250, flue_sov_cost=125000,
        air_cv_nb=150, air_cv_cost=125600,
        air_fm_nb=200, air_fm_cost=57000,
        gas_cv_nb=32,  gas_cv_cost=83000,
        gas_fm_nb=40,  gas_fm_cost=49000,
        pneu_damp_nb=250, pneu_damp_cost=125000,
    ),
    # 1500 KW: Excel sheet has corrupt qty (24 instead of 2).
    # Burner/blower/panel/gas_train verified; valve sizes same as 1000 KW (same pipe range ≤1500 KW).
    1500: dict(
        burner_cost=196797.43, blower_cost=180000, markup=1.8,
        panel_cost=450000, gas_train_cost=139300,
        gas_sol_nb=40, gas_sol_cost=14720,
        gas_bv_nb=40,  gas_bv_cost=5100,  gas_bv_qty=10,
        gas_hose_nb=40, gas_hose_cost=2000, gas_hose_qty=10,
        pg_burner=3000,
        air_sov_nb=200, air_sov_cost=80000,
        air_mbv_nb=200, air_mbv_cost=31178,
        flue_sov_nb=250, flue_sov_cost=125000,
        air_cv_nb=150, air_cv_cost=125600,
        air_fm_nb=200, air_fm_cost=57000,
        gas_cv_nb=32,  gas_cv_cost=83000,
        gas_fm_nb=40,  gas_fm_cost=49000,
        pneu_damp_nb=250, pneu_damp_cost=125000,
    ),
    2000: dict(
        burner_cost=346253.53, blower_cost=190000, markup=1.8,
        panel_cost=450000, gas_train_cost=144100,
        gas_sol_nb=65, gas_sol_cost=43000,
        gas_bv_nb=65,  gas_bv_cost=13400, gas_bv_qty=10,
        gas_hose_nb=65, gas_hose_cost=4200, gas_hose_qty=10,
        pg_burner=4000,
        air_sov_nb=250, air_sov_cost=125000,
        air_mbv_nb=250, air_mbv_cost=38378,
        flue_sov_nb=350, flue_sov_cost=177000,
        air_cv_nb=200, air_cv_cost=144000,
        air_fm_nb=250, air_fm_cost=58000,
        gas_cv_nb=50,  gas_cv_cost=96960,
        gas_fm_nb=65,  gas_fm_cost=50000,
        pneu_damp_nb=350, pneu_damp_cost=177000,
    ),
    2500: dict(
        burner_cost=356349.12, blower_cost=220000, markup=1.8,
        panel_cost=600000, gas_train_cost=224000,
        gas_sol_nb=65, gas_sol_cost=43000,
        gas_bv_nb=65,  gas_bv_cost=13400, gas_bv_qty=10,
        gas_hose_nb=65, gas_hose_cost=4200, gas_hose_qty=10,
        pg_burner=4000,
        air_sov_nb=250, air_sov_cost=125000,
        air_mbv_nb=250, air_mbv_cost=38378,
        flue_sov_nb=400, flue_sov_cost=227500,
        air_cv_nb=200, air_cv_cost=144000,
        air_fm_nb=250, air_fm_cost=58000,
        gas_cv_nb=50,  gas_cv_cost=96960,
        gas_fm_nb=65,  gas_fm_cost=50000,
        pneu_damp_nb=400, pneu_damp_cost=350000,
    ),
    3000: dict(
        burner_cost=474790.79, blower_cost=220000, markup=1.8,
        panel_cost=600000, gas_train_cost=224000,
        gas_sol_nb=80, gas_sol_cost=44000,
        gas_bv_nb=80,  gas_bv_cost=17000, gas_bv_qty=10,
        gas_hose_nb=80, gas_hose_cost=6900, gas_hose_qty=10,
        pg_burner=4000,
        air_sov_nb=300, air_sov_cost=148000,
        air_mbv_nb=300, air_mbv_cost=48055,
        flue_sov_nb=450, flue_sov_cost=361020,
        air_cv_nb=250, air_cv_cost=189540,
        air_fm_nb=300, air_fm_cost=60000,
        gas_cv_nb=65,  gas_cv_cost=97810,
        gas_fm_nb=80,  gas_fm_cost=51000,
        pneu_damp_nb=450, pneu_damp_cost=350000,
    ),
    4500: dict(
        burner_cost=663539.31, blower_cost=320000, markup=1.8,
        panel_cost=600000, gas_train_cost=295200,
        gas_sol_nb=80, gas_sol_cost=44000,
        gas_bv_nb=80,  gas_bv_cost=17000, gas_bv_qty=2,   # qty=2 in legacy (not 10)
        gas_hose_nb=80, gas_hose_cost=6900, gas_hose_qty=2,
        pg_burner=4000,
        air_sov_nb=350, air_sov_cost=177000,
        air_mbv_nb=350, air_mbv_cost=61700,
        flue_sov_nb=500, flue_sov_cost=453470,
        air_cv_nb=300, air_cv_cost=213240,
        air_fm_nb=350, air_fm_cost=64000,
        gas_cv_nb=65,  gas_cv_cost=97810,
        gas_fm_nb=80,  gas_fm_cost=51000,
        pneu_damp_nb=500, pneu_damp_cost=350000,
    ),
    # 6000 KW: markup=2.0, different gas line structure (shut-off + butterfly per burner)
    # plus separate gas skid items (gate valve, PSOV, pressure switch)
    6000: dict(
        burner_cost=868568.06, blower_cost=450000, markup=2.0,
        panel_cost=700000, gas_train_cost=0,  # gas skid items listed separately below
        gas_sol_nb=350, gas_sol_cost=177000,        # Shut-Off Valve per burner (not solenoid)
        gas_bv_nb=350,  gas_bv_cost=61700,  gas_bv_qty=2,   # Butterfly Valve (not ball valve)
        gas_hose_nb=350, gas_hose_cost=90000, gas_hose_qty=2,
        pg_burner=4000,
        air_sov_nb=400, air_sov_cost=227500,
        air_mbv_nb=400, air_mbv_cost=83750,
        flue_sov_nb=700, flue_sov_cost=1048800,
        air_cv_nb=350, air_cv_cost=242250,
        air_fm_nb=400, air_fm_cost=70500,
        gas_cv_nb=300, gas_cv_cost=213240,
        gas_fm_nb=350, gas_fm_cost=64000,
        pneu_damp_nb=650, pneu_damp_cost=625771.2,
    ),
}

MODEL_KWS = sorted(REGEN_MODELS.keys())  # [500, 1000, 1500, 2000, 2500, 3000, 4500, 6000]

# PLC cost by num_pairs (Siemens S7-1200 for 1-2 pairs, S7-1500 for 3+)
_PLC_COST = {1: 300000, 2: 300000, 3: 600000, 4: 750000, 5: 800000, 6: 900000}


def select_model(required_kw: float) -> int:
    """Return the smallest model KW >= required_kw (caps at 6000 KW)."""
    for kw in MODEL_KWS:
        if kw >= required_kw:
            return kw
    return MODEL_KWS[-1]


def build_regen_df(kw: int, markup: float = None, num_pairs: int = 1) -> pd.DataFrame:
    """
    Build full BOM DataFrame for the given KW model.

    kw        : one of MODEL_KWS (500, 1000, 1500, 2000, 2500, 3000, 4500, 6000)
    markup    : selling price multiplier; if None uses model default (1.8 or 2.0)
    num_pairs : number of pairs — affects PLC selection only (default 1)
    """
    m  = REGEN_MODELS[kw]
    mk = markup if markup is not None else m['markup']
    rows = []

    def add(section, item, spec, qty, cost_unit):
        rows.append(_make_row(section, item, spec, qty, cost_unit, mk))

    # ── 1. BURNER SET ─────────────────────────────────────────────────────────
    add("BURNER SET", f"Burner with Regenerator ({kw} KW)",
        f"Regenerative burner with heat-storage media, complete",         2, m['burner_cost'])
    add("BURNER SET", "Pilot Burner",        "7 KW",                     2, 10000)
    add("BURNER SET", "Burner Controller",   "",                          2,  3600)
    add("BURNER SET", "Ignition Transformer","",                          2,  3300)
    add("BURNER SET", "UV Sensor",           "",                          2,  5500)

    # ── 2. GAS LINE — Pilot ───────────────────────────────────────────────────
    add("GAS LINE — PILOT", "Pilot Regulator",       "NB15",             2,  4400)
    add("GAS LINE — PILOT", "Pilot Solenoid Valve",  "NB15",             2,  4300)
    add("GAS LINE — PILOT", "Flexible Hose",         "NB15",             2,  1500)
    add("GAS LINE — PILOT", "Ball Valve",            "NB15",             2,  1400)
    add("GAS LINE — PILOT", "Pressure Gauge 0-500",  "",                 2,  3000)

    # ── 3. GAS LINE — Burner ──────────────────────────────────────────────────
    # 6000 KW uses shut-off valve + butterfly valve (not solenoid + ball valve × 10)
    sol_label = "Shut-Off Valve" if kw == 6000 else "Solenoid Valve"
    bv_label  = "Butterfly Valve" if kw == 6000 else "Ball Valve"
    add("GAS LINE — BURNER", sol_label,
        f"NB{m['gas_sol_nb']}",              2,                          m['gas_sol_cost'])
    add("GAS LINE — BURNER", bv_label,
        f"NB{m['gas_bv_nb']}",               m['gas_bv_qty'],            m['gas_bv_cost'])
    add("GAS LINE — BURNER", "Flexible Hose",
        f"NB{m['gas_hose_nb']}",             m['gas_hose_qty'],          m['gas_hose_cost'])
    add("GAS LINE — BURNER", "Pressure Gauge 0-500",  "",                2,  m['pg_burner'])

    # ── 4. AIR LINE — Pilot / UV / UV Cooling ─────────────────────────────────
    add("AIR LINE — PILOT/UV", "Ball Valve UV",       "NB15",            8,  1400)
    add("AIR LINE — PILOT/UV", "Flexible Hose UV",    "NB15",            4,  1500)
    add("AIR LINE — PILOT/UV", "Ball Valve Pilot",    "NB15",            4,  1400)
    add("AIR LINE — PILOT/UV", "Flexible Hose Pilot", "NB15",            4,  1500)

    # ── 5. AIR LINE — Burner ──────────────────────────────────────────────────
    add("AIR LINE — BURNER", "Shut-Off Valve Air",
        f"DN{m['air_sov_nb']}",              2,                          m['air_sov_cost'])
    add("AIR LINE — BURNER", "Manual Butterfly Valve Air",
        f"DN{m['air_mbv_nb']}",              2,                          m['air_mbv_cost'])
    add("AIR LINE — BURNER", "Pressure Gauge 0-1000", "",                2,  4000)
    add("AIR LINE — BURNER", "Shut-Off Valve Flue Gas",
        f"DN{m['flue_sov_nb']}",             2,                          m['flue_sov_cost'])
    add("AIR LINE — BURNER", "Thermocouple with TT",  "",                4,  5000)

    # ── 6. TEMPERATURE CONTROL ────────────────────────────────────────────────
    add("TEMP CONTROL", "Air Control Valve",
        f"DN{m['air_cv_nb']}",               1,                          m['air_cv_cost'])
    add("TEMP CONTROL", "Air Flow Meter (DPT)",
        f"DN{m['air_fm_nb']}",               1,                          m['air_fm_cost'])
    add("TEMP CONTROL", "Gas Control Valve",
        f"DN{m['gas_cv_nb']}",               1,                          m['gas_cv_cost'])
    add("TEMP CONTROL", "Gas Flow Meter (DPT)",
        f"DN{m['gas_fm_nb']}",               1,                          m['gas_fm_cost'])
    add("TEMP CONTROL", "Thermocouple with TT (Furnace)", "",            1,  25000)
    add("TEMP CONTROL", "DPT",               "",                         1,  45000)
    add("TEMP CONTROL", "Pneumatic Damper",
        f"DN{m['pneu_damp_nb']}",            1,                          m['pneu_damp_cost'])
    add("TEMP CONTROL", "Manual Damper",     "",                         1,  40000)

    # ── 7. BLOWER ─────────────────────────────────────────────────────────────
    add("BLOWER", "Combustion Blower (40\" WG)",
        f"With motor, for {kw} KW",           2,                          m['blower_cost'])

    # ── 8. CONTROLS ───────────────────────────────────────────────────────────
    plc_cost = _PLC_COST.get(num_pairs, 900000)
    add("CONTROLS", "PLC with HMI",
        "Siemens S7-1200/1500 with touch panel",  1,                     plc_cost)
    add("CONTROLS", "Control Panel",          "",                         1,  m['panel_cost'])

    # ── 9. GAS TRAIN ─────────────────────────────────────────────────────────
    if m['gas_train_cost'] > 0:
        add("GAS TRAIN", "NG Gas Train",
            f"Complete, for {kw} KW",             1,                     m['gas_train_cost'])
    else:
        # 6000 KW uses a custom gas skid instead of a packaged gas train
        add("GAS TRAIN", "Gate Valve",                "DN350",            1, 275000)
        add("GAS TRAIN", "Pressure Gauge with Manual Cock", "",           1,   4000)
        add("GAS TRAIN", "Pneumatic Shut-Off Valve",  "DN350",            1, 177000)
        add("GAS TRAIN", "Pressure Switch Low/High",  "",                 2,  12000)

    return pd.DataFrame(rows)


def _make_row(section, item, spec, qty, cost_unit, markup):
    total_cost = qty * cost_unit
    sell_unit  = cost_unit * markup
    total_sell = qty * sell_unit
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
