"""Choosing the combustion air blowers from the catalogue.

A blower count comes from how much air one machine can move, not from how much
power the set draws. The BRF sheet types "Blower 100HP/40" x 3", and the app
used to arrive at the 3 by dividing the computed 208.46 shaft HP by a 100 HP
motor rating — which is not a capacity calculation at all. The two agreed by
coincidence on the 60 TPH job and would part company on any other.

This selects real machines from blower_pricelist_master instead: the smallest
model that covers the duty on its own, or, when nothing does, the largest model
in the class and as many of them as it takes.
"""
import os
import sqlite3

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.environ.get("VLPH_DB_PATH") or os.path.join(BASE_DIR, "vlph.db")

# The reheating furnace runs bigger machines than the standard ENCON ranges
# carry, so it has a section of its own — quoted PCBLZ centrifugal blowers.
BRF_CLASS = "BRF CENTRIFUGAL"
DEFAULT_CLASS = BRF_CLASS
PRESSURE_CLASSES = {"brf": BRF_CLASS, "40": "HIGH PRESSURE",
                    "28": "MEDIUM PRESSURE"}

# Quotation Q26-ET-0227, 27 March 2026, ENCON Thermal Engineers. Air quantity
# is CMH at NTP, which is Nm3/hr; CFM follows the same 1.7 the rest of the app
# converts on. The price is the whole set — blower, motor, coupling and
# anti-vibration pads — because that is what gets bought.
#
#   PCBLZ-105-130   68,000 CMH   620 mm WC   212.14 BHP   180 kW / 240 HP motor
#       blower 4,15,000 + motor 6,80,000 + coupling 14,000 + pads 12,500
#   PCBLZ-100-130   60,000 CMH   620 mm WC   184.69 BHP   160 kW / 215 HP motor
#       blower 3,77,000 + motor 4,80,000 + coupling 14,000 + pads 10,500
BRF_CATALOGUE = [
    # model, motor HP, fan BHP, Nm3/hr, pressure, blower only, set, motor
    ("PCBLZ-100-130", 215.0, 184.69, 60000.0, "620 mm WC", 377000.0, 881500.0, 480000.0),
    ("PCBLZ-105-130", 240.0, 212.14, 68000.0, "620 mm WC", 415000.0, 1121500.0, 680000.0),
]
CFM_PER_NM3HR = 1.7


def seed_brf_blowers(conn):
    """Put the quoted machines in blower_pricelist_master. Idempotent, and it
    leaves an edited row alone — a price corrected in the UI stays corrected."""
    # The fan's shaft power at its duty point. It is what a blower is actually
    # chosen on: air alone cannot be compared across machines quoted at
    # different static pressures.
    cols = {r[1] for r in conn.execute("PRAGMA table_info(blower_pricelist_master)")}
    if "fan_bhp" not in cols:
        conn.execute("ALTER TABLE blower_pricelist_master ADD COLUMN fan_bhp REAL")
    have = {r[0] for r in conn.execute(
        "SELECT model FROM blower_pricelist_master WHERE section = ?",
        (BRF_CLASS,))}
    n = 0
    for model, hp, bhp, nm3, press, wo_motor, w_motor, motor in BRF_CATALOGUE:
        if model in have:
            conn.execute("UPDATE blower_pricelist_master SET fan_bhp = ? "
                         "WHERE section = ? AND model = ? AND fan_bhp IS NULL",
                         (bhp, BRF_CLASS, model))
            continue
        conn.execute(
            "INSERT INTO blower_pricelist_master "
            "(section, model, hp, fan_bhp, cfm, nm3_per_hr, pressure, "
            " price_without_motor, price_with_motor, motor_price_abb) "
            "VALUES (?,?,?,?,?,?,?,?,?,?)",
            (BRF_CLASS, model, hp, bhp, round(nm3 / CFM_PER_NM3HR, 2), nm3, press,
             wo_motor, w_motor, motor))
        n += 1
    conn.commit()
    if n:
        conn.commit()
    return n


def _catalogue(pressure_class):
    """Models in one pressure class that have both a capacity and a price."""
    try:
        conn = sqlite3.connect(DB_PATH)
        rows = conn.execute(
            "SELECT model, hp, cfm, nm3_per_hr, pressure, price_with_motor, "
            "       COALESCE(fan_bhp, 0) "
            "FROM blower_pricelist_master "
            "WHERE section = ? AND cfm IS NOT NULL AND price_with_motor IS NOT NULL "
            "ORDER BY cfm", (pressure_class,)).fetchall()
        conn.close()
    except Exception:
        return []
    return [dict(model=r[0], hp=r[1], cfm=r[2], nm3_per_hr=r[3],
                 pressure=r[4], price=r[5], fan_bhp=r[6]) for r in rows]


# The duty figure the app computes is CFM x 40 / 3200, i.e. CFM / 80. A fan at
# 40" SP and 65% efficiency draws CFM / 103, so the workbook's constant already
# carries efficiency plus about 29% headroom: 16,676 CFM is 161.5 BHP at the
# fan and 208.46 by this formula. That makes it a motor-selection figure, and
# it belongs against the motor rating.
#
# Measuring it against the vendor's quoted fan BHP instead double-counts the
# margin — their own quotes size the motor 13-16% above the fan (184.69 -> 215,
# 212.14 -> 240), so the headroom is in the figure twice.
MOTOR_MARGIN = 1.0        # raise to demand more motor than the duty figure


def select_blowers(cfm_required, pressure_class=DEFAULT_CLASS, hp_required=0.0,
                   motor_margin=MOTOR_MARGIN):
    """The machine to use and how many.

    Chosen on motor rating against the computed duty power, then on air for
    ranges that quote no power at all.
    """
    models = _catalogue(pressure_class)
    if not models:
        return None

    need = hp_required * (motor_margin or 1.0)
    rated = [m for m in models if m.get("hp")]
    if need > 0 and rated:
        rated.sort(key=lambda m: m["hp"])
        for m in rated:
            if m["hp"] >= need:
                return _result(m, 1, cfm_required, pressure_class, True,
                               hp_required, "motor", need)
        # Nothing single covers it: the largest, and as many as it takes.
        big = rated[-1]
        count = int(-(-need // big["hp"]))
        return _result(big, count, cfm_required, pressure_class, False,
                       hp_required, "motor", need)

    if cfm_required <= 0:
        return None
    for m in models:
        if m["cfm"] >= cfm_required:
            return _result(m, 1, cfm_required, pressure_class, True,
                           hp_required, "air", need)
    big = models[-1]
    count = int(-(-cfm_required // big["cfm"]))          # ceil
    return _result(big, count, cfm_required, pressure_class, False,
                   hp_required, "air", need)


def _result(m, count, cfm_required, pressure_class, single,
            hp_required=0.0, basis="air", hp_demanded=0.0):
    provided = m["cfm"] * count
    return {
        "fan_bhp_each": m.get("fan_bhp") or 0.0,
        "fan_bhp_total": round((m.get("fan_bhp") or 0.0) * count, 2),
        "hp_required": round(hp_required, 2),
        "hp_demanded": round(hp_demanded, 2),
        "basis": basis,
        "model": m["model"], "hp_each": m["hp"], "cfm_each": m["cfm"],
        "nm3hr_each": m["nm3_per_hr"], "pressure": m["pressure"],
        "price_each": m["price"], "count": count,
        "total_price": round(m["price"] * count, 2),
        "cfm_required": round(cfm_required, 2),
        "cfm_provided": round(provided, 2),
        "spare_cfm": round(provided - cfm_required, 2),
        "pressure_class": pressure_class,
        "single_machine": single,
        "installed_hp": round(m["hp"] * count, 2),
    }
