"""
Regenerative-burner pricing — the single DB-backed source for every price the
regen offer BOM uses (mirrors the HPU / Blower / Burner pricelist architecture).

Every price that build_regen_df() puts on a line now lives in the Pricelist
(component_price_master) under a REGEN_* category, seeded idempotently at
startup from the code constants in regen_builder. The offer LAYOUT and logic
stay in build_regen_df; only the numbers are sourced here, so they become
editable in the same Price-Master UI as the other products.

    seed_regen_pricelist(conn)   -> insert missing rows (idempotent, non-destructive)
    load_regen_prices(conn, kw)  -> {model, flat, plc, gas_skid} with DB overrides
                                     (falls back to the code constant per field)
"""

from bom.regen_builder import (
    REGEN_MODELS, _FLAT, _PLC_COST, _GAS_SKID_6000, MODEL_KWS,
)

# ── Pricelist categories (shown as sections in the Price-Master UI) ───────────
CAT_BURNER   = "REGEN Burner Set"
CAT_GAS      = "REGEN Gas Line"
CAT_AIR      = "REGEN Air Line"
CAT_TEMP     = "REGEN Temp Control"
CAT_BLOWER   = "REGEN Blower"
CAT_CONTROLS = "REGEN Controls"
CAT_GASTRAIN = "REGEN Gas Train"
CAT_FITTINGS = "REGEN Fittings"

# ── Per-KW model fields: REGEN_MODELS[kw][field] -> (label, category) ─────────
# gas_train_cost is skipped for any model whose value is 0 (6000 KW uses a skid).
PERKW_FIELDS = [
    ("burner_cost",     "Burner with Regenerator",         CAT_BURNER),
    ("blower_cost",     "Combustion Blower",               CAT_BLOWER),
    ("panel_cost",      "Control Panel",                   CAT_CONTROLS),
    ("gas_train_cost",  "NG Gas Train",                    CAT_GASTRAIN),
    ("gas_sol_cost",    "Gas Solenoid / Shut-Off Valve",   CAT_GAS),
    ("gas_bv_cost",     "Gas Ball / Butterfly Valve",      CAT_GAS),
    ("gas_hose_cost",   "Gas Flexible Hose",               CAT_GAS),
    ("pg_burner",       "Burner Gas Pressure Gauge 0-500", CAT_GAS),
    ("air_sov_cost",    "Air Shut-Off Valve",              CAT_AIR),
    ("air_mbv_cost",    "Air Manual Butterfly Valve",      CAT_AIR),
    ("flue_sov_cost",   "Flue Gas Shut-Off Valve",         CAT_AIR),
    ("air_cv_cost",     "Air Control Valve",               CAT_TEMP),
    ("air_fm_cost",     "Air Flow Meter (DPT)",            CAT_TEMP),
    ("gas_cv_cost",     "Gas Control Valve",               CAT_TEMP),
    ("gas_fm_cost",     "Gas Flow Meter (DPT)",            CAT_TEMP),
    ("pneu_damp_cost",  "Pneumatic Damper",                CAT_TEMP),
]

# NB/size field that carries the spec for a per-KW price field (display only).
_PERKW_NB = {
    "gas_sol_cost": "gas_sol_nb", "gas_bv_cost": "gas_bv_nb", "gas_hose_cost": "gas_hose_nb",
    "air_sov_cost": "air_sov_nb", "air_mbv_cost": "air_mbv_nb", "flue_sov_cost": "flue_sov_nb",
    "air_cv_cost": "air_cv_nb", "air_fm_cost": "air_fm_nb", "gas_cv_cost": "gas_cv_nb",
    "gas_fm_cost": "gas_fm_nb", "pneu_damp_cost": "pneu_damp_nb",
}

# ── Flat fields (same price across all KW models): _FLAT[key] -> (label, cat) ──
FLAT_FIELDS = [
    ("pilot_burner",         "Pilot Burner",                    CAT_BURNER),
    ("burner_controller",    "Burner Controller",               CAT_BURNER),
    ("ignition_transformer", "Ignition Transformer",            CAT_BURNER),
    ("uv_sensor",            "UV Sensor",                       CAT_BURNER),
    ("pilot_regulator",      "Pilot Regulator NB15",            CAT_GAS),
    ("pilot_solenoid",       "Pilot Solenoid Valve NB15",       CAT_GAS),
    ("pilot_pg_500",         "Pilot Gas Pressure Gauge 0-500",  CAT_GAS),
    ("ball_valve_nb15",      "Ball Valve NB15",                 CAT_FITTINGS),
    ("flex_hose_nb15",       "Flexible Hose NB15",              CAT_FITTINGS),
    ("air_pg_1000",          "Air Pressure Gauge 0-1000",       CAT_AIR),
    ("thermocouple_tt",      "Thermocouple with TT",            CAT_AIR),
    ("furnace_thermocouple", "Furnace Thermocouple with TT",    CAT_TEMP),
    ("dpt",                  "DPT (flow / pressure / temp)",    CAT_TEMP),
    ("manual_damper",        "Manual Damper",                   CAT_TEMP),
]

# ── 6000 KW gas-skid fields: _GAS_SKID_6000[key] -> (label, category) ─────────
GASSKID_FIELDS = [
    ("gate_valve",      "Gate Valve DN350",                     CAT_GASTRAIN),
    ("pg_cock",         "Gas Pressure Gauge with Manual Cock",  CAT_GASTRAIN),
    ("pneu_sov",        "Pneumatic Shut-Off Valve DN350",       CAT_GASTRAIN),
    ("pressure_switch", "Pressure Switch Low / High",           CAT_GASTRAIN),
]


# ── Item-name helpers (canonical, must match between seed and load) ───────────
# Every regen item name is prefixed so it can't collide with — or silently share
# a row with — a generic item another product owns (e.g. "Ignition Transformer",
# "UV Sensor"). Regen owns its own editable rows.
PREFIX = "Regen: "


def _perkw_item(label, kw):
    return f"{PREFIX}{label} ({kw}KW)"


def _flat_item(label):
    return f"{PREFIX}{label}"


def _plc_item(pairs):
    return f"{PREFIX}PLC with HMI ({'≤2 pairs' if pairs <= 2 else f'{pairs} pairs'})"


# PLC prices seed one row per distinct pair-count (1 and 2 share ≤2 pairs).
_PLC_SEED_PAIRS = [2, 3, 4, 5, 6]


def _f(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def seed_regen_pricelist(conn) -> int:
    """Insert any missing REGEN pricelist rows into component_price_master.
    Idempotent and non-destructive — existing rows (and live edits) are kept."""
    inserted = 0

    def _ins(item, cat, price, spec=None):
        nonlocal inserted
        if price is None:
            return
        if conn.execute("SELECT 1 FROM component_price_master WHERE item=? LIMIT 1",
                        (item,)).fetchone():
            return
        # Store the exact value (burner_cost carries a computed .xx) so the seed
        # reproduces the current offer to the paisa.
        p = float(price)
        conn.execute(
            "INSERT INTO component_price_master (item, category, company, unit, "
            "price, previous_price, specification) VALUES (?,?,?,?,?,?,?)",
            (item, cat, "ENCON", "nos", p, p, spec))
        inserted += 1

    # Per-KW model prices
    for kw in MODEL_KWS:
        m = REGEN_MODELS[kw]
        for field, label, cat in PERKW_FIELDS:
            val = m.get(field)
            if field == "gas_train_cost" and not val:
                continue          # 6000 KW: no packaged gas train
            nb = m.get(_PERKW_NB.get(field, ""), None)
            spec = f"{kw} KW" + (f" · DN{int(nb)}" if nb else "")
            _ins(_perkw_item(label, kw), cat, val, spec)

    # Flat prices (one row each, shared across KW models)
    for key, label, cat in FLAT_FIELDS:
        _ins(_flat_item(label), cat, _FLAT.get(key))

    # PLC prices (per pair-count)
    for pairs in _PLC_SEED_PAIRS:
        _ins(_plc_item(pairs), CAT_CONTROLS, _PLC_COST.get(pairs),
             f"{'1-2' if pairs <= 2 else pairs} pair(s)")

    # 6000 KW gas-skid prices
    for key, label, cat in GASSKID_FIELDS:
        _ins(_perkw_item(label, 6000), cat, _GAS_SKID_6000.get(key), "6000 KW")

    conn.commit()
    return inserted


def _price(conn, item):
    r = conn.execute("SELECT price FROM component_price_master WHERE item=? LIMIT 1",
                     (item,)).fetchone()
    return _f(r[0]) if r else None


def load_regen_prices(conn, kw: int) -> dict:
    """Resolve every price build_regen_df needs for `kw`, DB value winning over
    the code constant. Returns {model, flat, plc, gas_skid} — same shapes as the
    regen_builder constants, so build_regen_df can consume it directly."""
    model = dict(REGEN_MODELS[kw])
    for field, label, cat in PERKW_FIELDS:
        if field == "gas_train_cost" and not model.get(field):
            continue
        p = _price(conn, _perkw_item(label, kw))
        if p is not None:
            model[field] = p

    flat = dict(_FLAT)
    for key, label, cat in FLAT_FIELDS:
        p = _price(conn, _flat_item(label))
        if p is not None:
            flat[key] = p

    plc = dict(_PLC_COST)
    for pairs in _PLC_SEED_PAIRS:
        p = _price(conn, _plc_item(pairs))
        if p is not None:
            plc[pairs] = p
    plc[1] = plc.get(2, plc.get(1))     # 1 pair shares the ≤2-pair price

    gas_skid = dict(_GAS_SKID_6000)
    for key, label, cat in GASSKID_FIELDS:
        p = _price(conn, _perkw_item(label, 6000))
        if p is not None:
            gas_skid[key] = p

    return dict(model=model, flat=flat, plc=plc, gas_skid=gas_skid)
