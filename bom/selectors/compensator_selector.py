import sqlite3


def select_compensator(air_flow_nm3hr: float) -> dict:
    """
    Select Compensator from DB
    """

    conn = sqlite3.connect("vlph.db")
    cursor = conn.cursor()

    cursor.execute("""
        SELECT nb, price
        FROM compensator_master
        WHERE ? BETWEEN min_flow AND max_flow
        LIMIT 1
    """, (air_flow_nm3hr,))

    row = cursor.fetchone()
    conn.close()

    if not row:
        raise ValueError(
            f"No Compensator found for {air_flow_nm3hr} Nm3/hr"
        )

    return {
        "nb": row[0],
        "price": row[1],
    }