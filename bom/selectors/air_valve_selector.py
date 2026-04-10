import sqlite3


def select_motorized_control_valve(air_flow_nm3hr: float) -> dict:
    """
    Select Motorized Control Valve from DB
    """

    conn = sqlite3.connect("vlph.db")
    cursor = conn.cursor()

    cursor.execute("""
        SELECT nb, rated_flow, price
        FROM motorized_valve_master
        WHERE ? BETWEEN min_flow AND max_flow
        LIMIT 1
    """, (air_flow_nm3hr,))

    row = cursor.fetchone()
    conn.close()

    if not row:
        raise ValueError(
            f"No Motorized Control Valve found for {air_flow_nm3hr} Nm3/hr"
        )

    return {
        "nb": row[0],
        "flow_nm3hr": row[1],
        "price": row[2],
    }



def select_butterfly_valve(nb: int) -> dict:
    """
    Select Butterfly Valve from component_price_master.
    Picks CAIR shut-off butterfly entries:
      'SHUT OFF VALVE 040NB (Butterfly)' ... 'SHUT OFF VALVE 350NB (Butterfly)'
    Returns the smallest available NB >= requested.
    """
    conn = sqlite3.connect("vlph.db")
    cursor = conn.cursor()

    cursor.execute("""
        SELECT specification, price
        FROM component_price_master
        WHERE item LIKE 'SHUT OFF VALVE %NB (Butterfly)'
          AND CAST(SUBSTR(item, 16, 3) AS INTEGER) >= ?
        ORDER BY CAST(SUBSTR(item, 16, 3) AS INTEGER) ASC
        LIMIT 1
    """, (nb,))

    row = cursor.fetchone()
    conn.close()

    if not row:
        raise ValueError(f"No Butterfly Valve found for NB >= {nb}")

    spec = row[0] or ""           # e.g. '040NB (Butterfly)'
    nb_val = int(spec[:3]) if spec[:3].isdigit() else nb
    return {
        "nb":    nb_val,
        "price": float(row[1]),
        "make":  "CAIR",
    }