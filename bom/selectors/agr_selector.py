import sqlite3


def select_agr(ng_flow_nm3hr: float) -> dict:
    """
    AGR selector using database (vlph.db)
    """

    conn = sqlite3.connect("vlph.db")
    cursor = conn.cursor()

    cursor.execute("""
        SELECT nb, price
        FROM agr_master
        WHERE ? BETWEEN min_flow AND max_flow
        LIMIT 1
    """, (ng_flow_nm3hr,))

    row = cursor.fetchone()
    conn.close()

    if not row:
        raise ValueError(f"No AGR found for flow {ng_flow_nm3hr}")

    return {
        "nb": row[0],
        "price": row[1],
    }