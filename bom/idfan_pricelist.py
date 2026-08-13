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
    return [dict(zip(COLUMNS, r)) for r in (rows or CATALOGUE)]


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
