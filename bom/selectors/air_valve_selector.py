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
    Select Butterfly Valve from DB based on NB
    """

    conn = sqlite3.connect("vlph.db")
    cursor = conn.cursor()

    cursor.execute("""
        SELECT price
        FROM butterfly_valve_master
        WHERE nb = ?
        LIMIT 1
    """, (nb,))

    row = cursor.fetchone()
    conn.close()

    if not row:
        raise ValueError(f"No Butterfly Valve found for NB {nb}")

    return {
        "nb": nb,
        "price": row[0],
    }