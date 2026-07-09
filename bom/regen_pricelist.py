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
    "pneu_damp":  ("Pneumatic Damper",                  "ENCON"),
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


def control_panel_price(conn, kw):
    """Regen control panel, per burner KW ('Control Panel {kw} KW')."""
    r = conn.execute("SELECT price FROM component_price_master WHERE item=? LIMIT 1",
                     (f"Control Panel {kw} KW",)).fetchone()
    return _f(r[0]) if r else None


def burner_price(conn, kw):
    """Burner + Regenerator, per KW ('Burner with Regenerator {kw} KW')."""
    r = conn.execute("SELECT price FROM component_price_master WHERE item=? LIMIT 1",
                     (f"Burner with Regenerator {kw} KW",)).fetchone()
    return _f(r[0]) if r else None


def blower_price_ic(conn, kw):
    """Combustion blower from the internal-costing blower pricelist (with motor),
    for the ENCON 40" WG model at this KW's HP."""
    from bom.regen_builder import _BLOWER_HP
    from bom.blower_pricelist import blower_price as _bp
    hp = (_BLOWER_HP.get(kw) or "").replace("HP", "").strip()
    if not hp:
        return None
    try:
        return _f(_bp(conn, f"ENCON 40/{hp}", with_motor=True))
    except Exception:
        return None


def gas_train_price(conn, flow_nm3hr):
    """NG gas train from the 'Gas Train' pricelist — the smallest flow band whose
    upper bound covers the required Nm³/hr."""
    best = None
    for _item, spec, price in conn.execute(
            "SELECT item, specification, price FROM component_price_master "
            "WHERE category='Gas Train'"):
        m = re.search(r"-\s*(\d+)", str(spec))   # upper bound of "X-Y Nm³/hr"
        p = _f(price)
        if not m or p is None:
            continue
        hi = int(m.group(1))
        if flow_nm3hr <= hi and (best is None or hi < best[0]):
            best = (hi, p)
    return best[1] if best else None


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
    "pneu_damp_cost":("pneu_damp",  "pneu_damp_nb"),
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
    # Burner PG 0-500 = the small pressure gauge (pneumatic damper is a sized
    # valve, resolved in the loop above — separate from the manual damper).
    pg = flat_price(conn, "pilot_pg_500")
    if pg is not None:
        model["pg_burner"] = pg
    # Control panel + burner+regenerator: dedicated per-KW Pricelist rows.
    cp = control_panel_price(conn, kw)
    if cp is not None:
        model["panel_cost"] = cp
    bp = burner_price(conn, kw)
    if bp is not None:
        model["burner_cost"] = bp
    # Combustion blower from the internal-costing blower pricelist.
    bl = blower_price_ic(conn, kw)
    if bl is not None:
        model["blower_cost"] = bl
    # NG gas train from the 'Gas Train' pricelist, by NG flow (only KW that have one).
    if model.get("gas_train_cost"):
        from bom.regen_builder import _PIPE_SIZES
        flow = _PIPE_SIZES.get(kw, {}).get("ng_flow")
        if flow:
            gt = gas_train_price(conn, flow)
            if gt is not None:
                model["gas_train_cost"] = gt

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
