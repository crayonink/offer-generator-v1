import sqlite3


def select_encon_mg_burner(required_gas_flow_nm3hr: float) -> dict:
    """
    Select ENCON MG Burner from database.
    Rounds UP to nearest max_flow.
    """

    conn = sqlite3.connect("vlph.db")
    cursor = conn.cursor()

    cursor.execute("""
        SELECT model, max_flow, price
        FROM burner_master
        WHERE max_flow >= ?
        ORDER BY max_flow ASC
        LIMIT 1
    """, (required_gas_flow_nm3hr,))

    row = cursor.fetchone()
    conn.close()

    if not row:
        raise ValueError(
            f"No ENCON MG burner available for "
            f"{required_gas_flow_nm3hr:.1f} Nm3/hr"
        )

    return {
        "model": row[0],
        "max_flow_nm3hr": row[1],
        "price": row[2],
    }