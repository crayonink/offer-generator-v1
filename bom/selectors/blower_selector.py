import sqlite3


def select_blower(air_flow_nm3hr: float) -> dict:
    """
    Select Combustion Air Blower from DB
    """

    conn = sqlite3.connect("vlph.db")
    cursor = conn.cursor()

    cursor.execute("""
        SELECT hp, pressure_mm_wc, rated_flow, price
        FROM blower_master
        WHERE ? BETWEEN min_flow AND max_flow
        LIMIT 1
    """, (air_flow_nm3hr,))

    row = cursor.fetchone()
    conn.close()

    if not row:
        raise ValueError(
            f"No Blower found for {air_flow_nm3hr} Nm3/hr"
        )

    return {
        "hp": row[0],
        "pressure_mm_wc": row[1],
        "flow_nm3hr": row[2],
        "price": row[3],
    }