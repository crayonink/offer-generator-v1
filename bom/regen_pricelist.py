"""
Regen price resolver — sources every regen BOM price from the CENTRALISED
Pricelist (component_price_master). Regen adds no product-specific pricing of
its own; each item is name-matched to its Pricelist row (flat items by exact
name, sized valves by category + NB + make). If a row is missing, the resolver
returns None so build_regen_df falls back to its code constant.

    flat_price(conn, key)          -> price for a flat bought-out item
    valve_price(conn, vtype, nb)   -> (nb_used, price, gap) for a sized valve
    plc_price(conn, pairs)         -> PLC-with-HMI tier price
    control_panel_price(conn)      -> Control Panel price
    load_regen_prices(conn, kw)    -> {model, flat, plc, gas_skid, oil} for build_regen_df
"""

import re

from bom.regen_builder import (
    REGEN_MODELS, _FLAT, _PLC_COST, _GAS_SKID_6000, _OIL,
)

# ── regen flat key -> exact Pricelist item name ──────────────────────────────
FLAT_ITEM = {
    "pilot_burner":         "ENCON-PB-LPG-10KW",
    "burner_controller":    "Sequence Controller",
    "ignition_transformer": "Ignition Transformer",
    "uv_sensor":            "UV Sensor with Air Jacket",
    "pilot_regulator":      "GAS PRESSURE REGULATORS ( Pmax = 5 Bar) 025 RCS04V0000-M5XXX",
    "pilot_solenoid":       "Solenoid Valve 15 NB",
    "pilot_pg_500":         "PRESSURE GAUGE WITH TNV (HGURU)",
    "ball_valve_nb15":      "BALL VALVE 15 NB #01 L3RBTC/L3RSWC",
    "flex_hose_nb15":       "FLEXIBLE HOSE 15 NB 1500mm",
    "air_pg_1000":          "PRESSURE GAUGE WITH TNV (BAUMER)",
    "thermocouple_tt":      "Thermocouple Small",
    "furnace_thermocouple": "THERMOCOUPLE",
    "dpt":                  "DPT",
    "manual_damper":        "DAMPER MANUAL",
}

# ── sized valve type -> (Pricelist category, make) ; item carries "<nb> NB" ───
SIZED = {
    "solenoid":   ("Solenoid Valve - Automatic Reset", "MADAS"),
    "ball_valve": ("Ball Valve",                        "L&T"),
    "flex_hose":  ("Flexible Hose",                     "BENGAL"),
    "butterfly":  ("Butterfly Valve",                   "L&T"),
    "shutoff":    ("Pneumatic Shut Off Valve",          "DEMBLA"),
    "control":    ("Pneumatic Control Valve",           "DEMBLA"),
    "flow_meter": ("Flow Meter (DPT)",                  "HONEYWELL"),
}


def _f(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def flat_price(conn, key):
    name = FLAT_ITEM.get(key)
    if not name:
        return None
    r = conn.execute("SELECT price FROM component_price_master WHERE item=? LIMIT 1",
                     (name,)).fetchone()
    return _f(r[0]) if r else None


def _nb_options(conn, category, company):
    q = "SELECT item, price FROM component_price_master WHERE category=?"
    args = [category]
    if company:
        q += " AND company=?"
        args.append(company)
    out = {}
    for item, price in conn.execute(q, args):
        m = re.search(r"(\d+)\s*NB", str(item))
        p = _f(price)
        if m and p is not None:
            nb = int(m.group(1))
            if nb not in out or p < out[nb]:   # cheapest wins on dup
                out[nb] = p
    return out


def valve_price(conn, vtype, nb):
    """(nb_used, price, gap) for a sized valve — snaps to the smallest Pricelist
    NB >= nb. gap=True when nb exceeds the category's max (price falls back to
    the largest available)."""
    cat, company = SIZED[vtype]
    opts = _nb_options(conn, cat, company)
    if not opts:
        return None, None, False
    ge = sorted(n for n in opts if n >= nb)
    if ge:
        return ge[0], opts[ge[0]], False
    mx = max(opts)
    return mx, opts[mx], True


def plc_price(conn, pairs):
    label = "1-2 Pair" if pairs <= 2 else f"{pairs} Pair"
    r = conn.execute("SELECT price FROM component_price_master WHERE item=? LIMIT 1",
                     (f"PLC with HMI ({label})",)).fetchone()
    return _f(r[0]) if r else None


def control_panel_price(conn):
    r = conn.execute("SELECT price FROM component_price_master WHERE item='CONTROL PANEL' "
                     "LIMIT 1").fetchone()
    return _f(r[0]) if r else None


# per-KW model valve field -> (valve type, NB field on the model)
_MODEL_VALVE = {
    "gas_sol_cost":  ("solenoid",   "gas_sol_nb"),
    "gas_bv_cost":   ("ball_valve", "gas_bv_nb"),
    "gas_hose_cost": ("flex_hose",  "gas_hose_nb"),
    "air_sov_cost":  ("shutoff",    "air_sov_nb"),
    "air_mbv_cost":  ("butterfly",  "air_mbv_nb"),
    "flue_sov_cost": ("shutoff",    "flue_sov_nb"),
    "air_cv_cost":   ("control",    "air_cv_nb"),
    "air_fm_cost":   ("flow_meter", "air_fm_nb"),
    "gas_cv_cost":   ("control",    "gas_cv_nb"),
    "gas_fm_cost":   ("flow_meter", "gas_fm_nb"),
}


def load_regen_prices(conn, kw: int) -> dict:
    """Resolve every price build_regen_df needs for `kw` from the Pricelist,
    falling back to the code constant when a row is missing. Returns
    {model, flat, plc, gas_skid, oil} — same shapes build_regen_df consumes."""
    model = dict(REGEN_MODELS[kw])

    # Per-KW sized valves -> Pricelist by (type, NB at the NG size for this KW).
    for field, (vtype, nbf) in _MODEL_VALVE.items():
        nb = model.get(nbf)
        if nb:
            _, p, _ = valve_price(conn, vtype, int(nb))
            if p is not None:
                model[field] = p
    # Pneumatic Damper = the (manual) DAMPER MANUAL row; Control Panel from Bought
    # Out; burner PG 0-500 = the small pressure gauge. Gas train / burner /
    # blower have no bought-out Pricelist row — kept from the code constant.
    dm = flat_price(conn, "manual_damper")
    if dm is not None:
        model["pneu_damp_cost"] = dm
    cp = control_panel_price(conn)
    if cp is not None:
        model["panel_cost"] = cp
    pg = flat_price(conn, "pilot_pg_500")
    if pg is not None:
        model["pg_burner"] = pg

    flat = dict(_FLAT)
    for key in FLAT_ITEM:
        p = flat_price(conn, key)
        if p is not None:
            flat[key] = p

    plc = dict(_PLC_COST)
    for pairs in (1, 2, 3, 4, 5, 6):
        p = plc_price(conn, pairs)
        if p is not None:
            plc[pairs] = p

    return dict(model=model, flat=flat, plc=plc,
                gas_skid=dict(_GAS_SKID_6000), oil=dict(_OIL))
