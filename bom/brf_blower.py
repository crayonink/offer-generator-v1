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

# The sheet's blowers are 40" W.G., which is the high pressure range.
DEFAULT_CLASS = "HIGH PRESSURE"
PRESSURE_CLASSES = {"40": "HIGH PRESSURE", "28": "MEDIUM PRESSURE"}


def _catalogue(pressure_class):
    """Models in one pressure class that have both a capacity and a price."""
    try:
        conn = sqlite3.connect(DB_PATH)
        rows = conn.execute(
            "SELECT model, hp, cfm, nm3_per_hr, pressure, price_with_motor "
            "FROM blower_pricelist_master "
            "WHERE section = ? AND cfm IS NOT NULL AND price_with_motor IS NOT NULL "
            "ORDER BY cfm", (pressure_class,)).fetchall()
        conn.close()
    except Exception:
        return []
    return [dict(model=r[0], hp=r[1], cfm=r[2], nm3_per_hr=r[3],
                 pressure=r[4], price=r[5]) for r in rows]


def select_blowers(cfm_required, pressure_class=DEFAULT_CLASS):
    """The machine to use and how many, for a required air flow in CFM.

    Returns a dict, or None when the catalogue has nothing to offer.
    """
    models = _catalogue(pressure_class)
    if not models or cfm_required <= 0:
        return None

    # One machine if one will do — the smallest that covers it, so a small
    # furnace is not handed the largest blower in the range.
    for m in models:
        if m["cfm"] >= cfm_required:
            return _result(m, 1, cfm_required, pressure_class, single=True)

    # Otherwise the largest in the class, and as many as it takes.
    big = models[-1]
    count = int(-(-cfm_required // big["cfm"]))          # ceil
    return _result(big, count, cfm_required, pressure_class, single=False)


def _result(m, count, cfm_required, pressure_class, single):
    provided = m["cfm"] * count
    return {
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
