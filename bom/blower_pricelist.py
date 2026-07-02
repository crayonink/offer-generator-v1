"""
Blower pricing — the PERKIN Rates basis (the single source for the Internal-
Costing Blower tab AND the equipment offer).

Each blower model (blower_pricelist_master, MEDIUM / HIGH PRESSURE sections)
carries a stored Amount (blower cost, weight-based) and a Motor Price (ABB).
Price is computed from those:

    PRICE without motor = Amount × 1.8
    PRICE with motor    = (Amount × 1.8) + (Motor × 1.5)

MEDIUM PRESSURE = ENCON 28" WG series, HIGH PRESSURE = ENCON 40" WG series.
"""

import sqlite3

WITHOUT_MARKUP = 1.8   # Amount -> price without motor
MOTOR_MARKUP   = 1.5   # motor added on top at ×1.5

# The two PERKIN tables, shown (and priced) in this order.
SECTIONS = ["MEDIUM PRESSURE", "HIGH PRESSURE"]


def _f(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def price_without_motor(amount) -> float:
    return round(_f(amount) * WITHOUT_MARKUP, 2)


def price_with_motor(amount, motor) -> float:
    return round(_f(amount) * WITHOUT_MARKUP + _f(motor) * MOTOR_MARKUP, 2)


# The blower-alone (without-motor) price lives in the pricelist Rates tab under
# these categories (company PERKINS), so it's editable there and fetched here.
ALONE_CATEGORIES = ["Blower Alone (28 inch)", "Blower Alone (40 inch)"]
_SECTION_TO_ALONE_CAT = {
    "MEDIUM PRESSURE": "Blower Alone (28 inch)",
    "HIGH PRESSURE":   "Blower Alone (40 inch)",
}


def alone_prices(conn: sqlite3.Connection) -> dict:
    """{model -> blower-alone price} from the pricelist Rates rows."""
    q = ("SELECT item, price FROM component_price_master WHERE category IN (%s)"
         % ",".join("?" * len(ALONE_CATEGORIES)))
    return {r[0]: _f(r[1]) for r in conn.execute(q, ALONE_CATEGORIES)}


def seed_blower_alone(conn: sqlite3.Connection) -> int:
    """Idempotently create the 'Blower Alone' pricelist rows (one per model,
    company PERKINS), seeded with the without-motor price (Amount × 1.8). After
    that the Rates row is the editable source that the Blower tab/offer fetch."""
    rows = conn.execute(
        "SELECT section, model, per_kg_amount, hp FROM blower_pricelist_master "
        "WHERE section IN ('MEDIUM PRESSURE','HIGH PRESSURE')").fetchall()
    inserted = 0
    for section, model, amount, hp in rows:
        cat = _SECTION_TO_ALONE_CAT.get(section)
        if not cat:
            continue
        if conn.execute("SELECT 1 FROM component_price_master WHERE item=? AND "
                        "category=? LIMIT 1", (model, cat)).fetchone():
            continue
        price = price_without_motor(amount)
        hpv = _f(hp)
        spec = None
        if hpv:
            spec = f"{int(hpv) if hpv == int(hpv) else hpv} HP"
        conn.execute(
            "INSERT INTO component_price_master (item, category, company, unit, "
            "price, previous_price, specification) VALUES (?,?,?,?,?,?,?)",
            (model, cat, "PERKINS", "nos", price, price, spec))
        inserted += 1
    conn.commit()
    return inserted


def blower_price(conn: sqlite3.Connection, model: str, with_motor: bool = False) -> float:
    """Blower price for a model. Blower-alone (without motor) is fetched from the
    pricelist Rates row; with_motor adds Motor × 1.5. Single source for every
    blower offer path."""
    r = conn.execute(
        "SELECT per_kg_amount, motor_price_abb FROM blower_pricelist_master "
        "WHERE model=? AND section IN ('MEDIUM PRESSURE','HIGH PRESSURE') LIMIT 1",
        (model,)).fetchone()
    if not r:
        return 0.0
    amount, motor = r[0], r[1]
    alone = alone_prices(conn).get(model)
    wo = alone if alone is not None else price_without_motor(amount)  # fallback pre-seed
    return round(wo + (_f(motor) * MOTOR_MARKUP if with_motor else 0.0), 2)


def blower_models(conn: sqlite3.Connection) -> dict:
    """The two PERKIN tables for the Internal-Costing Blower tab.
    Returns {section: [ {model, hp, weight, amount, motor, price_wo, price_w} ]}."""
    rows = conn.execute(
        "SELECT section, model, hp, blower_weight, per_kg_amount, motor_price_abb, "
        "cfm, nm3_per_hr, pressure FROM blower_pricelist_master "
        "WHERE section IN ('MEDIUM PRESSURE','HIGH PRESSURE') "
        "ORDER BY section, CAST(hp AS REAL)").fetchall()
    alone = alone_prices(conn)   # blower-alone (without motor) from Rates
    data = {}
    for section, model, hp, weight, amount, motor, cfm, nm3, pressure in rows:
        wo = alone.get(model)
        wo = round(wo) if wo is not None else round(price_without_motor(amount))
        data.setdefault(section, []).append({
            "model":     model,
            "hp":        _f(hp),
            "weight":    _f(weight),
            "amount":    round(_f(amount)),
            "motor":     round(_f(motor)),
            "cfm":       _f(cfm),
            "nm3_per_hr": _f(nm3),
            "price_wo":  wo,
            "price_w":   round(wo + _f(motor) * MOTOR_MARKUP),
            "alone_cat": _SECTION_TO_ALONE_CAT.get(section, ""),
        })
    return data


# ── Legacy: DM/IDM fabrication cost (pricelist-linked) ──────────────────────
# The fabricated mounting/drive structure from blower_dm_idm_master, priced
# from raw materials. Shown as "Legacy tables" below the PERKIN tables — a cost
# reference only; NOT the blower offer basis.
LEGACY_OVERHEAD = 1.3   # material subtotal -> factory cost
LEGACY_MARKUP   = 1.8   # factory cost -> selling
LEGACY_SECTIONS = ["BLOWER DM 28", "BLOWER DM 40", "BLOWER IDM"]

# blower_dm_idm_master column prefix -> (pricelist item, display label, unit)
LEGACY_COMPONENTS = [
    ("angle65_50",    "M.S. Angle 65,50",            "M.S. Angle 65,50",           "kg"),
    ("channel",       "M.S. Channel",                "M.S. Channel",               "kg"),
    ("sheet8mm",      "M.S. Sheet 8mm",              "M.S. Sheet 8mm",             "kg"),
    ("sheet4mm",      "M.S. Sheet 4mm",              "M.S. Sheet 4mm",             "kg"),
    ("sheet2mm",      "M.S. Sheet 2mm",              "M.S. Sheet 2mm",             "kg"),
    ("flat",          "M.S. Flat",                   "M.S. Flat",                  "kg"),
    ("ms_round",      "M.S. Round",                  "M.S. Round",                 "kg"),
    ("ci_hub",        "C.I. Hub",                    "C.I. Hub",                   "kg"),
    ("coupling",      "Coupling",                    "Coupling",                   "kg"),
    ("plumber_block", "Plumber block with Bearing",  "Plumber Block with Bearing", "nos"),
    ("hardware",      "Hardware Bolt",               "Hardware Bolt",              "kg"),
]


def legacy_rates(conn: sqlite3.Connection) -> dict:
    names = tuple(dict.fromkeys(c[1] for c in LEGACY_COMPONENTS))
    q = ("SELECT item, price FROM component_price_master WHERE item IN (%s)"
         % ",".join("?" * len(names)))
    return {r[0]: (r[1] or 0.0) for r in conn.execute(q, names)}


def legacy_compute(row: dict, rates: dict) -> dict:
    items = []
    subtotal = 0.0
    for prefix, price_item, label, unit in LEGACY_COMPONENTS:
        qty = _f(row.get(prefix + "_qty"))
        if qty <= 0:
            continue
        rate = rates.get(price_item, 0.0)
        amount = qty * rate
        subtotal += amount
        items.append({"s_no": len(items) + 1, "item": label, "qty": round(qty, 2),
                      "unit": unit, "rate": round(rate, 2), "rate_ref": price_item,
                      "amount": round(amount, 2)})
    factory = subtotal * LEGACY_OVERHEAD
    return {"model": row.get("model"), "items": items, "subtotal": round(subtotal),
            "factory_cost": round(factory), "selling": round(factory * LEGACY_MARKUP)}


def legacy_models(conn: sqlite3.Connection) -> dict:
    """DM/IDM fabrication breakdown per model, grouped by section, HP-sorted."""
    rates = legacy_rates(conn)
    cur = conn.execute("SELECT * FROM blower_dm_idm_master")
    cols = [d[0] for d in cur.description]
    data = {}
    for r in cur.fetchall():
        row = dict(zip(cols, r))
        data.setdefault(row.get("section") or "—", []).append(legacy_compute(row, rates))

    def _hp(m):
        try:
            return float(str(m.get("model") or "").split("/")[0])
        except (TypeError, ValueError):
            return float("inf")
    for sec in data:
        data[sec].sort(key=_hp)
    return data
