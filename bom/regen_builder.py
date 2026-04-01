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

# ── Burner + Regenerator material weights (kg) per KW model ──────────────────
# Source: "Burner Sizing and costing" sheet, rows 35-42
# Columns: burner_ms, burner_refrac, regen_ms, regen_ss, regen_refrac, regen_ceramic, block_refrac
_BURNER_WEIGHTS = {
    500:  dict(burner_ms=167.53, burner_refrac=200.39, regen_ms=139.07, regen_ss=11.06,  regen_refrac=422.45, regen_ceramic=227.30, block_refrac=128.76),
    1000: dict(burner_ms=198.55, burner_refrac=237.50, regen_ms=139.08, regen_ss=23.70,  regen_refrac=559.67, regen_ceramic=349.07, block_refrac=158.96),
    1500: dict(burner_ms=232.05, burner_refrac=277.58, regen_ms=156.47, regen_ss=26.07,  regen_refrac=662.48, regen_ceramic=448.70, block_refrac=192.34),
    2000: dict(burner_ms=314.50, burner_refrac=376.21, regen_ms=284.47, regen_ss=28.44,  regen_refrac=1065.43, regen_ceramic=1028.09, block_refrac=276.98),
    2500: dict(burner_ms=347.46, burner_refrac=415.63, regen_ms=284.47, regen_ss=37.92,  regen_refrac=1065.43, regen_ceramic=1028.09, block_refrac=311.57),
    3000: dict(burner_ms=390.89, burner_refrac=467.59, regen_ms=349.28, regen_ss=44.24,  regen_refrac=1375.01, regen_ceramic=1556.60, block_refrac=357.67),
    4500: dict(burner_ms=485.20, burner_refrac=580.40, regen_ms=505.71, regen_ss=55.30,  regen_refrac=1776.45, regen_ceramic=2373.08, block_refrac=459.40),
    6000: dict(burner_ms=645.28, burner_refrac=771.89, regen_ms=660.56, regen_ss=55.30,  regen_refrac=2152.18, regen_ceramic=3193.34, block_refrac=635.85),
}

# Material cost rates (with 10% wastage applied)
_RATES = dict(
    ms_total=82.50,        # (50 material + 25 labour) × 1.10
    ss_total=82.50,
    refrac_total=89.10,    # (56 material + 25 labour) × 1.10
    ceramic_total=137.50,  # 125 × 1.10
)

# ── Pipe sizes per KW model (Natural Gas, 0.05 barg) ─────────────────────────
# Source: "Burner Pipe Size" sheet, rows 9-16
_PIPE_SIZES = {
    500:  dict(ng_flow=50,  air_flow=500,  flue_flow=550,  air_dn=125, gas_dn=30,  flue_dn=200),
    1000: dict(ng_flow=100, air_flow=1000, flue_flow=1100, air_dn=200, gas_dn=40,  flue_dn=250),
    1500: dict(ng_flow=150, air_flow=1500, flue_flow=1650, air_dn=200, gas_dn=50,  flue_dn=300),
    2000: dict(ng_flow=200, air_flow=2000, flue_flow=2200, air_dn=250, gas_dn=65,  flue_dn=350),
    2500: dict(ng_flow=250, air_flow=2500, flue_flow=2750, air_dn=250, gas_dn=65,  flue_dn=400),
    3000: dict(ng_flow=300, air_flow=3000, flue_flow=3300, air_dn=300, gas_dn=80,  flue_dn=450),
    4500: dict(ng_flow=450, air_flow=4500, flue_flow=4950, air_dn=350, gas_dn=80,  flue_dn=500),
    6000: dict(ng_flow=600, air_flow=6000, flue_flow=6600, air_dn=400, gas_dn=100, flue_dn=600),
}

# ── ENCON 40" WG Blower catalogue ────────────────────────────────────────────
# Source: "Blower" sheet
BLOWER_CATALOGUE = [
    dict(model="ENCON 40/5",   hp="5HP",   cfm=400,  nm3hr=680,  price_without_motor=56500,  price_with_motor=81000),
    dict(model="ENCON 40/7.5", hp="7.5HP", cfm=600,  nm3hr=1020, price_without_motor=60500,  price_with_motor=99500),
    dict(model="ENCON 40/10",  hp="10HP",  cfm=800,  nm3hr=1360, price_without_motor=76000,  price_with_motor=111000),
    dict(model="ENCON 40/15",  hp="15HP",  cfm=1200, nm3hr=2040, price_without_motor=87000,  price_with_motor=158000),
    dict(model="ENCON 40/20",  hp="20HP",  cfm=1600, nm3hr=2730, price_without_motor=91000,  price_with_motor=178000),
    dict(model="ENCON 40/25",  hp="25HP",  cfm=2000, nm3hr=3400, price_without_motor=111000, price_with_motor=215000),
    dict(model="ENCON 40/30",  hp="30HP",  cfm=2400, nm3hr=4000, price_without_motor=131000, price_with_motor=250000),
    dict(model="ENCON 40/40",  hp="40HP",  cfm=3200, nm3hr=5200, price_without_motor=151500, price_with_motor=316500),
    dict(model="ENCON 40/50",  hp="50HP",  cfm=4000, nm3hr=6500, price_without_motor=175000, price_with_motor=361000),
    dict(model="ENCON 40/60",  hp="60HP",  cfm=4800, nm3hr=7800, price_without_motor=198000, price_with_motor=441000),
]

# KW → blower HP mapping (from costing sheets)
_BLOWER_HP = {500:"10HP", 1000:"10HP", 1500:"15HP", 2000:"20HP", 2500:"25HP", 3000:"25HP", 4500:"40HP", 6000:"60HP"}


def get_supplementary_data(kw: int) -> dict:
    """Return burner sizing, pipe sizes, and blower selection for the given KW model."""
    w  = _BURNER_WEIGHTS[kw]
    r  = _RATES
    p  = _PIPE_SIZES[kw]
    hp = _BLOWER_HP[kw]
    blower = next(b for b in BLOWER_CATALOGUE if b['hp'] == hp)

    # Per-unit material cost breakdown (1 burner)
    burner_cost_detail = [
        dict(component="Burner Body", material="MS",         weight_kg=w['burner_ms'],    rate=r['ms_total'],     cost=round(w['burner_ms']    * r['ms_total'],    2)),
        dict(component="Burner Body", material="Refractory", weight_kg=w['burner_refrac'],rate=r['refrac_total'], cost=round(w['burner_refrac'] * r['refrac_total'],2)),
        dict(component="Regenerator", material="MS",         weight_kg=w['regen_ms'],     rate=r['ms_total'],     cost=round(w['regen_ms']     * r['ms_total'],    2)),
        dict(component="Regenerator", material="SS",         weight_kg=w['regen_ss'],     rate=r['ss_total'],     cost=round(w['regen_ss']     * r['ss_total'],    2)),
        dict(component="Regenerator", material="Refractory", weight_kg=w['regen_refrac'], rate=r['refrac_total'], cost=round(w['regen_refrac'] * r['refrac_total'],2)),
        dict(component="Regenerator", material="Ceramic Balls", weight_kg=w['regen_ceramic'], rate=r['ceramic_total'], cost=round(w['regen_ceramic'] * r['ceramic_total'],2)),
        dict(component="Burner Block", material="Refractory", weight_kg=w['block_refrac'], rate=r['refrac_total'], cost=round(w['block_refrac'] * r['refrac_total'],2)),
    ]
    total_unit_cost = sum(d['cost'] for d in burner_cost_detail)

    return dict(
        burner_sizing=dict(
            kw=kw,
            material_rates=dict(ms=r['ms_total'], ss=r['ss_total'], refractory=r['refrac_total'], ceramic_balls=r['ceramic_total']),
            cost_detail=burner_cost_detail,
            total_unit_cost=round(total_unit_cost, 2),
            total_pair_cost=round(total_unit_cost * 2, 2),
        ),
        pipe_sizes=dict(
            fuel="Natural Gas (NG)",
            pressure="0.05 barg",
            kw=kw,
            ng_flow_nm3hr=p['ng_flow'],
            air_flow_nm3hr=p['air_flow'],
            flue_flow_nm3hr=p['flue_flow'],
            air_line_dn=p['air_dn'],
            gas_line_dn=p['gas_dn'],
            flue_line_dn=p['flue_dn'],
        ),
        blower_selection=dict(
            kw=kw,
            selected_model=blower['model'],
            hp=blower['hp'],
            cfm=blower['cfm'],
            nm3hr=blower['nm3hr'],
            price_without_motor=blower['price_without_motor'],
            price_with_motor=blower['price_with_motor'],
            qty_per_pair=2,
            costing_price=REGEN_MODELS[kw]['blower_cost'],
        ),
        blower_catalogue=BLOWER_CATALOGUE,
    )


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
