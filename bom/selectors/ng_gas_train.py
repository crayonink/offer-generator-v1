import sqlite3


def select_ng_gas_train(required_flow_nm3hr: float, burner_model: str) -> dict:
    """
    Select NG Gas Train based on burner model.
    """

    conn = sqlite3.connect("vlph.db")
    cursor = conn.cursor()

    cursor.execute("""
        SELECT flow_nm3hr, price
        FROM gas_train_master
        WHERE burner_model = ?
        LIMIT 1
    """, (burner_model,))

    row = cursor.fetchone()
    conn.close()

    if not row:
        raise ValueError(f"No Gas Train found for burner {burner_model}")

    return {
        "flow_nm3hr": row[0],
        "price": row[1],
    }