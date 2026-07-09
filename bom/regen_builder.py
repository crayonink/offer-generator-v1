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

# ── Flat line-item prices (same across all KW models) ────────────────────────
# These, REGEN_MODELS, _PLC_COST and _GAS_SKID_6000 are the code-side source of
# truth; they seed the Pricelist (component_price_master, REGEN_* categories) and
# are the fallback build_regen_df uses when no DB price is supplied. See
# bom/regen_pricelist.py.
# Flat bought-out prices — matched by name to the centralised Pricelist
# (component_price_master). The mapped Pricelist item is noted per line; values
# are a snapshot of the current Pricelist. Items marked TBD are unconfirmed and
# kept at their prior value.
_FLAT = {
    "pilot_burner":         11000,   # ENCON-PB-LPG-10KW (10 KW pilot)
    "burner_controller":    10000,   # Sequence Controller (LINEAR)
    "ignition_transformer":  5500,   # Ignition Transformer (DANFOSS)
    "uv_sensor":            13000,   # UV Sensor with Air Jacket (LINEAR)
    "pilot_regulator":     7313.59, # Gas Pressure Regulator 025 NB, 5 Bar (MADAS, −45%)
    "pilot_solenoid":        3317.03,# Solenoid Valve 15 NB (MADAS, −45%)
    "pilot_pg_500":          3000,   # Pressure Gauge with TNV (HGURU)
    "ball_valve_nb15":       1953,   # Ball Valve 15 NB (L&T)
    "flex_hose_nb15":         940,   # Flexible Hose 15 NB 1500mm (BENGAL)
    "air_pg_1000":           4000,   # Pressure Gauge with TNV (BAUMER)
    "thermocouple_tt":       5000,   # Thermocouple Small (Pricelist)
    "furnace_thermocouple": 36000,   # THERMOCOUPLE (TEMPSENS)
    "dpt":                  43150,   # DPT (HONEYWELL)
    "manual_damper":        50000,   # DAMPER MANUAL (ENCON)
}

# 6000 KW custom gas skid (used instead of a packaged gas train)
_GAS_SKID_6000 = {
    "gate_valve":      275000,
    "pg_cock":           4000,
    "pneu_sov":        177000,
    "pressure_switch":  12000,
}

# Oil fuel line — fuel='Oil' swaps these in for the gas fuel line, gas train,
# and the gas control valve / gas flow meter. All NB25 (oil lines are small and
# size-invariant across KW). Prices from Regen_BOM.xlsx (OIL sheet).
# Oil grades — all build the same regen oil line.
_OIL_FUELS = {"oil", "hsd", "ldo", "hdo", "fo", "sko", "cfo", "lshs"}

_OIL = {
    "solenoid_valve_oil":  7000,   # NB25, per burner
    "oil_control_valve":  25000,   # NB25
    "oil_flow_meter_dpt": 48000,   # NB25, DPT + orifice + flanges
    "tt_oil_line":         5000,   # temperature transmitter in oil line
    "pt_oil_line":        48000,   # pressure transmitter in oil line
}

# ── Per-fuel gas lines (low-CV gases: BFG / COG / Producer Gas) ───────────────
# Low-CV gases need far more volume than NG, so the gas + flue lines step up in
# NB (from regen_pipe_sizes), and they use a BUILT-UP line (no packaged NG gas
# train). Size-indexed prices below are the gas-regen master (Regen_BOM.xlsx,
# GAS sheet). The built-up gas line reuses the shut-off + manual butterfly valve
# prices (the same valves the air line uses), priced to NB400.
_GAS_CATALOG = {
    "solenoid":  {32: 13700, 40: 14720, 50: 17900, 65: 43000, 80: 44000, 100: 76000},
    "ball_valve":{32: 4925,  40: 5100,  50: 7200,  65: 13400, 80: 17000, 100: 26600},
    "flex_hose": {32: 1750,  40: 2000,  50: 3000,  65: 4200,  80: 6900,  100: 7650},
    "gas_cv":    {25: 83000, 32: 83000, 40: 83000, 50: 96960, 65: 97810, 80: 101900,
                  100: 110450, 150: 125600, 200: 144000, 250: 189540, 300: 213240, 350: 242250, 400: 261000},
    "gas_fm":    {32: 48000, 40: 49000, 50: 49700, 65: 50000, 80: 51000, 100: 52000,
                  150: 54000, 200: 57000, 250: 58000, 300: 60000, 350: 64000, 400: 70500},
    "shutoff":   {125: 50050, 200: 80000, 250: 125000, 300: 148000, 350: 177000, 400: 227500},
    "butterfly": {125: 12498, 200: 31178, 250: 38378, 300: 48055, 350: 61700, 400: 83750},
    # DN900 flue = same price as DN700 (per ENCON). Pneumatic damper is a flat
    # rate above DN400, extended to cover DN650/700/900.
    "flue_sov":  {200: 80000, 250: 125000, 300: 148000, 350: 177000, 400: 227500,
                  450: 361020, 500: 453470, 600: 838150, 700: 1048800, 900: 1048800},
    "pneu_damp": {200: 80000, 250: 125000, 300: 148000, 350: 177000,
                  400: 350000, 450: 350000, 500: 350000, 600: 350000,
                  650: 350000, 700: 350000, 900: 350000},
}

# UI fuel name -> regen_pipe_sizes.gas_type (for the per-fuel DN lookup)
_FUEL_PIPE_NAME = {
    "natural gas":       "Natural Gas (NG) 8600 Kcal/Nm³",
    "coke oven gas":     "Coke Oven Gas 4000 Kcal/Nm³",
    "producer gas":      "Producer Gas 1250 Kcal/Nm³",
    "blast furnace gas": "Blast Furnace Gas 720 Kcal/Nm³",
}


def _snap_price(cat_key, dn):
    """Price for a valve at the smallest catalog NB >= the pipe DN.
    Returns (nb_used, price, gap) — gap=True when DN exceeds the catalog's max
    (we fall back to the largest priced size and the caller should flag it)."""
    d = _GAS_CATALOG[cat_key]
    ge = sorted(n for n in d if n >= dn)
    if ge:
        return ge[0], d[ge[0]], False
    mx = max(d)
    return mx, d[mx], True


def _fuel_pipe_dn(db_path, fuel, kw):
    """(gas DN, flue DN) for a fuel + KW from regen_pipe_sizes, or (None, None)."""
    name = _FUEL_PIPE_NAME.get((fuel or "").strip().lower())
    if not (name and db_path):
        return None, None
    try:
        import sqlite3 as _s
        with _s.connect(db_path) as _c:
            r = _c.execute("SELECT dn_gas_mm, dn_flue_mm FROM regen_pipe_sizes "
                           "WHERE gas_type=? AND burner_size_kw=?", (name, kw)).fetchone()
        if r:
            return (int(r[0]) if r[0] else None, int(r[1]) if r[1] else None)
    except Exception:
        pass
    return None, None

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

# Material cost rates (with 10% wastage applied) — FALLBACK ONLY.
# Production overrides these from DB regen_material_rates (parsed from the REGEN
# costing workbook, "Costing Consideration" block) in main.py. Kept in sync with
# that workbook so the fallback can't diverge if the DB read fails.
# Formula: (material_cost + 25 labour) × 1.10 wastage
_RATES = dict(
    ms_total=82.50,        # MS: (50 + 25 labour) × 1.10
    ss_total=82.50,        # SS: (50 + 25 labour) × 1.10
    refrac_total=89.10,    # Refractory (Whyte Heat K castable): (56 + 25 labour) × 1.10
    ceramic_total=137.50,  # Ceramic balls: 125 × 1.10 (no labour line)
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


def build_regen_df(kw: int, markup: float = None, num_pairs: int = 1,
                   db_path: str = None, fuel: str = "Natural Gas") -> pd.DataFrame:
    """
    Build full BOM DataFrame for the given KW model.

    kw        : one of MODEL_KWS (500, 1000, 1500, 2000, 2500, 3000, 4500, 6000)
    markup    : selling price multiplier; if None uses model default (1.8 or 2.0)
    num_pairs : number of pairs — affects PLC selection only (default 1)
    db_path   : if given, every line price is sourced LIVE from the Pricelist
                (component_price_master, REGEN_* categories), falling back to the
                code constant per field. If None, the code constants are used.
    fuel      : "Oil" swaps the gas fuel line + gas train + gas control/flow
                meter for the oil line (Solenoid Valve Oil, Oil Control Valve,
                Oil Flow Meter, TT/PT). Any gas fuel builds the standard NG BOM.
    """
    _fuel_l = (fuel or "").strip().lower()
    # Any oil grade (HSD/LDO/HDO/FO/SKO/CFO/LSHS) builds the oil line; the
    # regen oil line is the same for every grade.
    is_oil = _fuel_l in _OIL_FUELS
    # Low-CV gases (BFG / COG / Producer Gas) resize the gas + flue lines per
    # fuel and use a built-up line (no packaged NG gas train). NG (and the
    # default) keep the standard gas BOM.
    is_lowcv = (not is_oil) and _fuel_l in ("coke oven gas", "producer gas", "blast furnace gas")
    gas_dn, flue_dn = (_fuel_pipe_dn(db_path, fuel, kw) if is_lowcv else (None, None))
    # Resolve prices — the centralised Pricelist wins over the code constants.
    flat, plc_map, skid, oil = _FLAT, _PLC_COST, _GAS_SKID_6000, _OIL
    m = REGEN_MODELS[kw]
    _conn = None
    if db_path:
        try:
            import sqlite3 as _sql
            _conn = _sql.connect(db_path)
            from bom.regen_pricelist import load_regen_prices
            _pr = load_regen_prices(_conn, kw)
            m, flat, plc_map, skid, oil = (_pr['model'], _pr['flat'], _pr['plc'],
                                           _pr['gas_skid'], _pr['oil'])
        except Exception:
            m, flat, plc_map, skid, oil = REGEN_MODELS[kw], _FLAT, _PLC_COST, _GAS_SKID_6000, _OIL

    # Sized-valve pricing for the per-fuel (low-CV) lines: Pricelist first, code
    # catalog (_GAS_CATALOG) as the fallback.
    def _snap(pl_type, cat_key, dn):
        if _conn:
            try:
                from bom.regen_pricelist import valve_price
                nbu, p, gap = valve_price(_conn, pl_type, dn)
                if p is not None:
                    return nbu, p, gap
            except Exception:
                pass
        return _snap_price(cat_key, dn)

    mk = markup if markup is not None else m['markup']
    rows = []

    def add(section, item, spec, qty, cost_unit, scale=True):
        # Base quantities are per pair; multiply by num_pairs unless the item is
        # one-per-system (scale=False) — matches the RegenCosting sheet formulas.
        q = qty * num_pairs if scale else qty
        rows.append(_make_row(section, item, spec, q, cost_unit, mk))

    # ── 1. BURNER SET ─────────────────────────────────────────────────────────
    add("BURNER SET", f"Burner with Regenerator ({kw} KW)",
        f"Regenerative burner with heat-storage media, complete",         2, m['burner_cost'])
    add("BURNER SET", "Pilot Burner",        "10 KW (LPG)",              2, flat['pilot_burner'], scale=False)
    add("BURNER SET", "Sequence Controller", "",                          2, flat['burner_controller'])
    add("BURNER SET", "Ignition Transformer","",                          2, flat['ignition_transformer'])
    add("BURNER SET", "UV Sensor",           "",                          2, flat['uv_sensor'])

    # ── 2. GAS LINE — Pilot ───────────────────────────────────────────────────
    add("GAS LINE — PILOT", "Pilot Regulator",       "NB15",             2, flat['pilot_regulator'])
    add("GAS LINE — PILOT", "Pilot Solenoid Valve",  "NB15",             2, flat['pilot_solenoid'])
    add("GAS LINE — PILOT", "Flexible Hose",         "NB15",             2, flat['flex_hose_nb15'])
    add("GAS LINE — PILOT", "Ball Valve",            "NB15",             2, flat['ball_valve_nb15'])
    add("GAS LINE — PILOT", "Pressure Gauge 0-500",  "",                 2, flat['pilot_pg_500'])

    # ── 3. FUEL LINE — Burner ─────────────────────────────────────────────────
    if is_oil:
        # Oil fuel line (NB25) — replaces the gas solenoid/ball-valve/hose bank.
        add("OIL LINE — BURNER", "Solenoid Valve (Oil)",          "NB25", 2, oil['solenoid_valve_oil'])
        add("OIL LINE — BURNER", "Temperature Transmitter (Oil)", "",     1, oil['tt_oil_line'])
        add("OIL LINE — BURNER", "Pressure Transmitter (Oil)",    "",     1, oil['pt_oil_line'])
    elif is_lowcv and gas_dn:
        # Built-up gas line sized to the fuel's gas DN. Up to NB100 it's the
        # solenoid + ball-valve + hose bank; above that (low-CV, large flow) it
        # becomes a shut-off + manual butterfly valve line (air-line valves).
        if gas_dn <= 100:
            nb, p, _  = _snap("solenoid",   "solenoid",   gas_dn); add("GAS LINE — BURNER", "Solenoid Valve",  f"NB{nb}",  2,  p)
            nb, p, _  = _snap("ball_valve", "ball_valve", gas_dn); add("GAS LINE — BURNER", "Ball Valve",      f"NB{nb}",  10, p)
            nb, p, _  = _snap("flex_hose",  "flex_hose",  gas_dn); add("GAS LINE — BURNER", "Flexible Hose",   f"NB{nb}",  10, p)
        else:
            nb, p, _  = _snap("shutoff",    "shutoff",    gas_dn); add("GAS LINE — BURNER", "Shut-Off Valve",  f"DN{nb}",  2,  p)
            nb, p, _  = _snap("butterfly",  "butterfly",  gas_dn); add("GAS LINE — BURNER", "Manual Butterfly Valve", f"DN{nb}", 2, p)
        add("GAS LINE — BURNER", "Pressure Gauge 0-500",  "",                2,  m['pg_burner'])
    else:
        # 6000 KW uses shut-off valve + butterfly valve (not solenoid + ball valve × 10)
        sol_label = "Shut-Off Valve" if kw == 6000 else "Solenoid Valve"
        bv_label  = "Butterfly Valve" if kw == 6000 else "Ball Valve"
        add("GAS LINE — BURNER", sol_label,
            f"NB{m['gas_sol_nb']}",              2,                          m['gas_sol_cost'])
        add("GAS LINE — BURNER", bv_label,
            f"NB{m['gas_bv_nb']}",               m['gas_bv_qty'],            m['gas_bv_cost'])
        add("GAS LINE — BURNER", "Flexible Hose",
            f"NB{m['gas_hose_nb']}",             m['gas_hose_qty'],          m['gas_hose_cost'], scale=False)
        add("GAS LINE — BURNER", "Pressure Gauge 0-500",  "",                2,  m['pg_burner'])

    # ── 4. AIR LINE — Pilot / UV / UV Cooling ─────────────────────────────────
    add("AIR LINE — PILOT/UV", "Ball Valve UV",       "NB15",            8, flat['ball_valve_nb15'])
    add("AIR LINE — PILOT/UV", "Flexible Hose UV",    "NB15",            4, flat['flex_hose_nb15'])
    add("AIR LINE — PILOT/UV", "Ball Valve Pilot",    "NB15",            4, flat['ball_valve_nb15'])
    add("AIR LINE — PILOT/UV", "Flexible Hose Pilot", "NB15",            4, flat['flex_hose_nb15'])

    # ── 5. AIR LINE — Burner ──────────────────────────────────────────────────
    add("AIR LINE — BURNER", "Shut-Off Valve Air",
        f"DN{m['air_sov_nb']}",              2,                          m['air_sov_cost'])
    add("AIR LINE — BURNER", "Manual Butterfly Valve Air",
        f"DN{m['air_mbv_nb']}",              2,                          m['air_mbv_cost'])
    add("AIR LINE — BURNER", "Pressure Gauge 0-1000", "",                2, flat['air_pg_1000'])
    if is_lowcv and flue_dn:
        nb, p, gap = _snap("shutoff", "flue_sov", flue_dn)
        spec = f"DN{flue_dn}" + (f" (priced at DN{nb} — verify)" if gap else "")
        add("AIR LINE — BURNER", "Shut-Off Valve Flue Gas", spec, 2, p)
    else:
        add("AIR LINE — BURNER", "Shut-Off Valve Flue Gas",
            f"DN{m['flue_sov_nb']}",             2,                          m['flue_sov_cost'])
    add("AIR LINE — BURNER", "Thermocouple with TT",  "",                4, flat['thermocouple_tt'])

    # ── 6. TEMPERATURE CONTROL ────────────────────────────────────────────────
    add("TEMP CONTROL", "Air Control Valve",
        f"DN{m['air_cv_nb']}",               1,                          m['air_cv_cost'])
    add("TEMP CONTROL", "Air Flow Meter (DPT)",
        f"DN{m['air_fm_nb']}",               1,                          m['air_fm_cost'], scale=False)
    if is_oil:
        add("TEMP CONTROL", "Oil Control Valve",   "NB25",              1, oil['oil_control_valve'])
        add("TEMP CONTROL", "Oil Flow Meter (DPT)","NB25",              1, oil['oil_flow_meter_dpt'])
    elif is_lowcv and gas_dn:
        nb, p, _ = _snap("control",    "gas_cv", gas_dn); add("TEMP CONTROL", "Gas Control Valve",    f"DN{nb}", 1, p)
        nb, p, _ = _snap("flow_meter", "gas_fm", gas_dn); add("TEMP CONTROL", "Gas Flow Meter (DPT)", f"DN{nb}", 1, p)
    else:
        add("TEMP CONTROL", "Gas Control Valve",
            f"DN{m['gas_cv_nb']}",               1,                          m['gas_cv_cost'])
        add("TEMP CONTROL", "Gas Flow Meter (DPT)",
            f"DN{m['gas_fm_nb']}",               1,                          m['gas_fm_cost'])
    add("TEMP CONTROL", "Thermocouple with TT (Furnace)", "",            1, flat['furnace_thermocouple'])
    add("TEMP CONTROL", "DPT",               "",                         1, flat['dpt'], scale=False)
    if is_lowcv and flue_dn:
        # Pneumatic damper = DAMPER MANUAL (flat Pricelist price, not per-NB).
        add("TEMP CONTROL", "Pneumatic Damper", f"DN{flue_dn}", 1, flat['manual_damper'])
    else:
        add("TEMP CONTROL", "Pneumatic Damper",
            f"DN{m['pneu_damp_nb']}",            1,                          m['pneu_damp_cost'])
    add("TEMP CONTROL", "Manual Damper",     "",                         1, flat['manual_damper'], scale=False)

    # ── 7. BLOWER ─────────────────────────────────────────────────────────────
    add("BLOWER", "Combustion Blower (40\" WG)",
        f"With motor, for {kw} KW",           2,                          m['blower_cost'], scale=False)

    # ── 8. CONTROLS ───────────────────────────────────────────────────────────
    plc_cost = plc_map.get(num_pairs, plc_map.get(6, 900000))
    add("CONTROLS", "PLC with HMI",
        "Siemens S7-1200/1500 with touch panel",  1,                     plc_cost, scale=False)
    add("CONTROLS", "Control Panel",          "",                         1,  m['panel_cost'], scale=False)

    # ── 9. GAS TRAIN ─────────────────────────────────────────────────────────
    if is_oil or is_lowcv:
        pass  # oil / low-CV gases (BFG/COG/PG) use a built-up line — no packaged train
    elif m['gas_train_cost'] > 0:
        add("GAS TRAIN", "NG Gas Train",
            f"Complete, for {kw} KW",             1,                     m['gas_train_cost'], scale=False)
    else:
        # 6000 KW uses a custom gas skid instead of a packaged gas train
        add("GAS TRAIN", "Gate Valve",                "DN350",            1, skid['gate_valve'], scale=False)
        add("GAS TRAIN", "Pressure Gauge with Manual Cock", "",           1, skid['pg_cock'], scale=False)
        add("GAS TRAIN", "Pneumatic Shut-Off Valve",  "DN350",            1, skid['pneu_sov'], scale=False)
        add("GAS TRAIN", "Pressure Switch Low/High",  "",                 2, skid['pressure_switch'], scale=False)

    if _conn is not None:
        _conn.close()
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
