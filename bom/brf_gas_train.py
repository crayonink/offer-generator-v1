"""The assembled gas train, from the IAPL quotation of 27 July 2026.

Subject: "Enquiry of Assembled Gas Train Flow 2500 Nm3/hr", ex-works, best
discounted. Five of every component, which is one train per fired zone on the
60 TPH furnace — so the quotation's 9,59,611 is five sets, not one.

  Main gas train
    EVF12AV-008          Madas auto-reset solenoid, DN150, 500 mbar   1,78,688
  Pilot train, 1/2"
    BLV-02-CS-04-1-F-I   IAPL ball valve, DN015, CS body, SS304 ball        522
    PGC100-S-250M-PI-2   pressure gauge, 0-250 mbar, 100 mm dial            743
    RC02V0020-030        Madas pressure regulator, DN015, 2 bar         5,261.80
    EWF02V-008           Madas auto-reset solenoid, DN015, 500 mbar     6,707.40
                                                          per train   1,91,922.20

The quotation's unit prices for the last two are rounded — 26,309 for five is
5,261.80 each, not the 5,262 printed, and 33,537 is 6,707.40. Taken from the
line totals the five sets come to 9,59,611 exactly.

The workbook types one Gas Train at 11,03,550 against the whole furnace. This
is a different thing: a train per zone, priced from its parts.
"""
import os
import sqlite3

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.environ.get("VLPH_DB_PATH") or os.path.join(BASE_DIR, "vlph.db")

TABLE = "brf_gas_train_master"

# section, item, model, description, unit price
GAS_TRAIN_COMPONENTS = [
    ("Main", "Main gas train solenoid valve", "EVF12AV-008",
     "Madas, auto reset, fast open/close, normally closed, DN150, "
     "flanged ANSI, 500 mbar, 230 VAC, flow reg., VITON", 178688.0),
    ("Pilot", "Isolation valve", "BLV-02-CS-04-1-F-I",
     "IAPL ball valve, DN015, threaded, 1 pc, full bore, CS body, "
     "SS304 ball, PTFE seat", 522.0),
    ("Pilot", "Pressure gauge", "PGC100-S-250M-PI-2",
     "0-250 mbar, 100 mm dial, G1/2, capsule", 743.0),
    ("Pilot", "Pressure regulator", "RC02V0020-030",
     "Madas, DN015, threaded EN, 2 bar, P2 = 40 to 110 mbar, VITON diaphragm",
     5261.8),
    ("Pilot", "Solenoid valve", "EWF02V-008",
     "Madas, auto reset, fast open/close, normally closed, DN015, "
     "threaded EN, 500 mbar, 230 VAC, flow reg., VITON", 6707.4),
]

# What the quotation was raised against, so a furnace that outgrows it says so.
QUOTED_FLOW_NM3HR = 2500.0
QUOTED_SETS = 5
QUOTED_TOTAL = 959611.0


def seed_gas_train(conn):
    """Create and fill the component table. Idempotent, and it leaves an
    edited price alone."""
    conn.execute(
        f"CREATE TABLE IF NOT EXISTS {TABLE} ("
        " id INTEGER PRIMARY KEY AUTOINCREMENT,"
        " section TEXT, item TEXT, model TEXT, description TEXT,"
        " unit_price REAL, make TEXT, source TEXT)")
    have = {r[0] for r in conn.execute(f"SELECT model FROM {TABLE}")}
    n = 0
    for section, item, model, desc, price in GAS_TRAIN_COMPONENTS:
        if model in have:
            continue
        make = "IAPL" if model.startswith("BLV") else (
            "Madas" if model.startswith(("EVF", "EWF", "RC")) else "")
        conn.execute(
            f"INSERT INTO {TABLE} "
            "(section, item, model, description, unit_price, make, source) "
            "VALUES (?,?,?,?,?,?,?)",
            (section, item, model, desc, price, make, "IAPL 27-07-2026"))
        n += 1
    if n:
        conn.commit()
    return n


def _components():
    try:
        conn = sqlite3.connect(DB_PATH)
        rows = conn.execute(
            f"SELECT section, item, model, description, unit_price, make "
            f"FROM {TABLE} ORDER BY id").fetchall()
        conn.close()
    except Exception:
        return []
    return [dict(section=r[0], item=r[1], model=r[2], description=r[3],
                 unit_price=r[4], make=r[5]) for r in rows]


def price_gas_train(train_count=1, firing_rate_nm3hr=0.0):
    """One assembled train per fired zone, priced from its parts.

    Returns None when the components have not been seeded, so the caller can
    fall back to the workbook's typed figure.
    """
    parts = _components()
    if not parts:
        return None
    count = max(0, int(train_count or 0))
    rows = [[p["section"], p["item"], p["model"], p["make"],
             p["unit_price"], count, round(p["unit_price"] * count, 2)]
            for p in parts]
    per_train = round(sum(p["unit_price"] for p in parts), 2)
    return {
        "rows": rows,
        "per_train": per_train,
        "train_count": count,
        "total_price": round(per_train * count, 2),
        "quoted_flow_nm3hr": QUOTED_FLOW_NM3HR,
        # Each train carries a share of the firing rate; if that share is above
        # what the quotation was raised for, the train is undersized.
        "flow_per_train_nm3hr": round(firing_rate_nm3hr / count, 2) if count else 0.0,
        "over_quoted_flow": bool(count and firing_rate_nm3hr / count
                                 > QUOTED_FLOW_NM3HR),
    }
