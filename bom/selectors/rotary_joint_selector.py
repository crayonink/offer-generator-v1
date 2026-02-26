import sqlite3


def select_rotary_joint(air_flow_nm3hr: float) -> dict:
    """
    Select Rotary Joint from DB
    """

    conn = sqlite3.connect("vlph.db")
    cursor = conn.cursor()

    cursor.execute("""
        SELECT nb, price
        FROM rotary_joint_master
        WHERE ? BETWEEN min_flow AND max_flow
        LIMIT 1
    """, (air_flow_nm3hr,))

    row = cursor.fetchone()
    conn.close()

    if not row:
        raise ValueError(
            f"No Rotary Joint found for {air_flow_nm3hr} Nm3/hr"
        )

    return {
        "nb": row[0],
        "price": row[1],
    }