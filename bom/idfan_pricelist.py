"""ID fan catalogue — the supplier price list (Ref. Q26-ETPL-0803).

Until now the ID fan borrowed the combustion blower's catalogue: it was sized
properly, then priced by looking up an ENCON blower frame of the same
horsepower. That was a stand-in. A blower frame is not an ID fan — the fan
runs hot flue gas at 500 mmWC, and the quoted price covers a different set of
parts — so the number was only ever the right order of magnitude, and it went
blank above 60 HP because the blower list stops there.

This is the real list. Nine models, PCBLY-30-135 to PCBLY-90-135, each rated
at 500 mmWC static and selected on air capacity. The supplier prices the fan
as a build-up:

    blower + base frame + pulley drive set + anti-vibration pads
      + inlet damper + inlet flexible connection
      + outlet damper + outlet flexible connection      = fan unit price
    fan unit price + motor                              = total price

Both totals are stored as the supplier gives them rather than recomputed, so
what we quote is what was quoted to us. Excluded per the list: taxes, packing,
freight, electrical controls and switchgear, installation and commissioning.
"""

import sqlite3

TABLE = "idfan_pricelist_master"

COLUMNS = [
    "air_cap_cmh", "static_pr_mmwc", "total_pr_mmwc", "model",
    "inlet_dia_mm", "impeller_dia_mm", "fan_rpm", "shaft_dia_mm", "bearing",
    "bhp_operating", "bhp_ambient", "motor_kw", "motor_hp", "motor_pole",
    "price_blower", "price_base_frame", "price_pulley_drive", "price_avm_pads",
    "price_inlet_damper", "price_inlet_flexi", "price_outlet_damper",
    "price_outlet_flexi", "price_fan_unit", "price_motor", "price_total",
]

# Ref. Q26-ETPL-0803, sheet "Table 1". BHP is given twice: at the operating
# temperature and at ambient. The ambient figure is roughly double, because
# cold gas is denser, and on every one of the nine models the quoted motor
# covers the OPERATING duty only — 20 HP against 14.59 operating and 28.53
# ambient, and so on down the list. So the motor on this list is the hot
# running motor. Starting one of these fans cold needs a VFD, a closed inlet
# damper, or a larger motor than the list quotes; see select_id_fan.
CATALOGUE = [
    (5000,  500, 520, "PCBLY-30-135", 300, 1350, 1588,  55, "22312K",  14.59,  28.53,  15,  20, 4, 145000, 10000, 12000, 2500,  6000,  3900,  6000,  3900, 189300,  58000, 247300),
    (9000,  500, 520, "PCBLY-40-135", 400, 1350, 1590,  60, "22313K",  24.14,  47.22,  22,  30, 4, 172000, 16000, 15000, 2500,  8000,  5200,  8000,  5200, 231900,  82100, 314000),
    (12000, 500, 515, "PCBLY-50-135", 500, 1350, 1600,  60, "22313K",  31.27,  61.15,  30,  40, 4, 199000, 22000, 18000, 4500, 10000,  6500, 10000,  6500, 276500, 114000, 390500),
    (16000, 500, 518, "PCBLY-55-135", 550, 1350, 1605,  65, "22315K",  40.79,  79.77,  37,  50, 4, 226000, 28000, 21000, 4500, 11000,  7150, 11000,  7150, 315800, 133200, 449000),
    (20000, 500, 520, "PCBLY-60-135", 600, 1350, 1610,  65, "22315K",  50.32,  98.41,  45,  60, 4, 253000, 34000, 24000, 5500, 12000,  7800, 12000,  7800, 356100, 160300, 516400),
    (24000, 500, 521, "PCBLY-65-135", 650, 1350, 1615,  75, "22317K",  59.87, 117.08,  55,  75, 4, 280000, 40000, 27000, 5500, 13000,  8450, 13000,  8450, 395400, 205300, 600700),
    (30000, 500, 524, "PCBLY-70-135", 700, 1350, 1620,  90, "22320K",  74.23, 145.17,  75, 100, 4, 307000, 46000, 30000, 5500, 14000,  9100, 14000,  9100, 434700, 311200, 745900),
    (38000, 500, 522, "PCBLY-80-135", 800, 1350, 1630, 100, "22322K",  92.53, 180.97,  90, 120, 4, 334000, 52000, 33000, 7500, 16000, 10400, 16000, 10400, 479300, 350400, 829700),
    (48000, 500, 522, "PCBLY-90-135", 900, 1350, 1645, 100, "22322K", 117.60, 230.00, 112, 150, 4, 361000, 58000, 36000, 7500, 18000, 11700, 18000, 11700, 521900, 430600, 952500),
]


def seed_idfan_catalog(conn: sqlite3.Connection) -> int:
    """Create the table and fill it if it is empty. Idempotent: a list already
    loaded — or edited from /pricelist — is left alone, so a redeploy does not
    undo a rate change."""
    cols_sql = ", ".join(
        f"{c} TEXT" if c in ("model", "bearing") else f"{c} REAL" for c in COLUMNS
    )
    conn.execute(f"CREATE TABLE IF NOT EXISTS {TABLE} ({cols_sql})")
    have = conn.execute(f"SELECT COUNT(*) FROM {TABLE}").fetchone()[0]
    if have:
        return 0
    conn.executemany(
        f"INSERT INTO {TABLE} ({', '.join(COLUMNS)}) "
        f"VALUES ({', '.join('?' * len(COLUMNS))})",
        CATALOGUE,
    )
    conn.commit()
    return len(CATALOGUE)


def idfan_models(conn: sqlite3.Connection = None) -> list:
    """Every model, smallest first.

    Falls back to the list as supplied when there is no connection to read, or
    when the table has not been seeded yet — callers size fans on paths that do
    not always carry a database handle, and a missing handle should not mean a
    missing fan.
    """
    rows = None
    if conn is not None:
        try:
            rows = conn.execute(
                f"SELECT {', '.join(COLUMNS)} FROM {TABLE} ORDER BY air_cap_cmh"
            ).fetchall()
        except sqlite3.Error:
            rows = None
    models = [dict(zip(COLUMNS, r)) for r in (rows or CATALOGUE)]
    _apply_pricelist(models, conn)
    return models


def _apply_pricelist(models, conn):
    """Let the pricelist rates win over the catalogue's.

    The catalogue is the engineering data — capacity, BHP, motor size — and
    that stays put. The money is a rate like any other, and a rate lives on
    /pricelist where the office can correct it. Without this the two would
    drift apart the moment anyone edited one: the pricelist would show the new
    figure and every BOM would keep quoting the old.

    Anything the pricelist does not carry keeps the catalogue's price, so a
    part-filled list cannot leave a fan with no money against it.
    """
    if conn is None:
        return
    try:
        fan = {r[0]: r[1] for r in conn.execute(
            "SELECT item, price FROM component_price_master "
            "WHERE category = ? AND price IS NOT NULL", (PM_FAN_CATEGORY,))}
        motor = {r[0]: r[1] for r in conn.execute(
            "SELECT item, price FROM component_price_master "
            "WHERE category = ? AND price IS NOT NULL", (PM_MOTOR_CATEGORY,))}
    except sqlite3.Error:
        return
    if not fan and not motor:
        return
    for m in models:
        f = fan.get(m["model"])
        if f is not None:
            m["price_fan_unit"] = float(f)
        mo = motor.get(f"ID Fan Motor {m['motor_hp']:g} HP")
        if mo is not None:
            m["price_motor"] = float(mo)
        m["price_total"] = round((m["price_fan_unit"] or 0)
                                 + (m["price_motor"] or 0), 2)


def select_id_fan(flow_m3hr: float, conn: sqlite3.Connection = None) -> dict | None:
    """The smallest fan that carries the flow.

    Selection is on air capacity, which is how the list is laid out and how the
    supplier quotes: every model is rated at the same 500 mmWC static, so the
    duty is settled by volume alone.

    The motor that comes with the model is the hot running motor — see the note
    on CATALOGUE. Callers that size for a cold start should compare their own
    motor kW against motor_kw here and say so when it is larger, rather than
    assume the quoted price carries a motor big enough to start cold.

    Returns None when the flow is past the largest model (48,000 CMH), so the
    caller can say so rather than quietly quoting the biggest fan on the list.
    """
    if not flow_m3hr or flow_m3hr <= 0:
        return None
    for row in idfan_models(conn):
        if row["air_cap_cmh"] >= flow_m3hr:
            return row
    return None


# ── The same list, on the pricelist page, keyed by HP ────────────────────────
# The catalogue above is selected on air capacity, which is how a fan is
# actually chosen. But the pricelist is read by people pricing a motor frame,
# and every other rotating item on it is filed by horsepower — "Blower Alone
# (40 inch)" and "Blower Motor" both are. So the fan and its motor go on in the
# same shape, split the same way, and can be corrected there like any other
# rate.
PM_FAN_CATEGORY = "ID Fan (500 mmWC)"
PM_MOTOR_CATEGORY = "ID Fan Motor"
PM_COMPANY = "ENCON"
PM_SOURCE = "Q26-ETPL-0803"


def seed_idfan_price_master(conn: sqlite3.Connection) -> int:
    """Put the ID fan and its motor on component_price_master, HP against price.

    Idempotent, and it leaves an edited rate alone: a row already present is
    not rewritten, so a redeploy cannot undo a correction made on /pricelist.
    """
    rows = idfan_models(conn)
    if not rows:
        return 0
    have = {
        r[0] for r in conn.execute(
            "SELECT item FROM component_price_master WHERE category IN (?, ?)",
            (PM_FAN_CATEGORY, PM_MOTOR_CATEGORY))
    }
    added = []
    for r in rows:
        hp = r["motor_hp"]
        spec = f"{hp:g} HP \u00b7 {r['air_cap_cmh']:,.0f} CMH \u00b7 500 mmWC"
        fan_item = r["model"]
        motor_item = f"ID Fan Motor {hp:g} HP"
        if fan_item not in have:
            added.append((fan_item, PM_FAN_CATEGORY, "nos", r["price_fan_unit"],
                          PM_COMPANY, PM_SOURCE, spec))
        if motor_item not in have:
            added.append((motor_item, PM_MOTOR_CATEGORY, "nos", r["price_motor"],
                          PM_COMPANY, PM_SOURCE, f"{hp:g} HP \u00b7 {r['motor_kw']:g} kW"))
            have.add(motor_item)          # models can share a motor size
    if not added:
        return 0
    conn.executemany(
        "INSERT INTO component_price_master "
        "(item, category, unit, price, company, type, specification) "
        "VALUES (?,?,?,?,?,?,?)", added)
    conn.commit()
    return len(added)


# ── Motor for a required shaft power ────────────────────────────────────────
# The fan is chosen on air volume; its motor is chosen on kilowatts. Those are
# two different questions, and the catalogue only answers the first — the motor
# it lists is the one the supplier matched to the hot running duty. A cold-start
# rating asks for more, and taking the catalogue's HP for it said 75 HP against
# a figure that works out at 113.
MOTOR_HP_LADDER = (20, 30, 40, 50, 60, 75, 100, 120, 150)
_HP_PER_KW = 1 / 0.746


def idfan_motor_for_kw(kw: float, conn: sqlite3.Connection = None) -> dict | None:
    """Smallest ID-fan motor that covers `kw`, with its price.

    Returns None past the largest motor on the list, so the caller can say so
    rather than quietly quoting the biggest one.
    """
    if not kw or kw <= 0:
        return None
    hp_needed = kw * _HP_PER_KW
    models = idfan_models(conn)
    by_hp = {}
    for m in models:
        by_hp.setdefault(m["motor_hp"], m)
    for hp in MOTOR_HP_LADDER:
        if hp + 1e-9 >= hp_needed and hp in by_hp:
            m = by_hp[hp]
            return {"motor_hp": hp, "motor_kw": m["motor_kw"],
                    "price_motor": m["price_motor"], "hp_required": hp_needed}
    return None
