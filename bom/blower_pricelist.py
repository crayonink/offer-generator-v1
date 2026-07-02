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


def blower_price(conn: sqlite3.Connection, model: str, with_motor: bool = False) -> float:
    """PERKIN price for a blower model. with_motor adds the motor at ×1.5.
    Single source used by every blower offer path."""
    r = conn.execute(
        "SELECT per_kg_amount, motor_price_abb FROM blower_pricelist_master "
        "WHERE model=? AND section IN ('MEDIUM PRESSURE','HIGH PRESSURE') LIMIT 1",
        (model,)).fetchone()
    if not r:
        return 0.0
    amount, motor = r[0], r[1]
    return price_with_motor(amount, motor) if with_motor else price_without_motor(amount)


def blower_models(conn: sqlite3.Connection) -> dict:
    """The two PERKIN tables for the Internal-Costing Blower tab.
    Returns {section: [ {model, hp, weight, amount, motor, price_wo, price_w} ]}."""
    rows = conn.execute(
        "SELECT section, model, hp, blower_weight, per_kg_amount, motor_price_abb, "
        "cfm, nm3_per_hr, pressure FROM blower_pricelist_master "
        "WHERE section IN ('MEDIUM PRESSURE','HIGH PRESSURE') "
        "ORDER BY section, CAST(hp AS REAL)").fetchall()
    data = {}
    for section, model, hp, weight, amount, motor, cfm, nm3, pressure in rows:
        data.setdefault(section, []).append({
            "model":     model,
            "hp":        _f(hp),
            "weight":    _f(weight),
            "amount":    round(_f(amount)),
            "motor":     round(_f(motor)),
            "cfm":       _f(cfm),
            "nm3_per_hr": _f(nm3),
            "price_wo":  round(price_without_motor(amount)),
            "price_w":   round(price_with_motor(amount, motor)),
        })
    return data
