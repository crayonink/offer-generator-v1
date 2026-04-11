import sqlite3
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DB_PATH = os.path.join(BASE_DIR, "vlph.db")


def select_ng_gas_train(required_flow_nm3hr: float) -> dict:
    """
    Select NG Gas Train based on required flow.
    Uses flow range logic (min_flow <= flow <= max_flow).
    """

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.execute("""
        SELECT inlet_nb, outlet_nb, min_flow, max_flow, price_inr
        FROM gas_train_master
        WHERE ? BETWEEN min_flow AND max_flow
        LIMIT 1
    """, (required_flow_nm3hr,))

    row = cursor.fetchone()

    # Fall back: if flow exceeds the largest available gas train, pick the
    # largest one. Low-CV fuels (Mixed Gas, Blast Furnace Gas) produce very
    # high volumetric flows that can exceed the pricelist range — the engineer
    # can scale up manually if needed.
    if not row:
        cursor.execute("""
            SELECT inlet_nb, outlet_nb, min_flow, max_flow, price_inr
            FROM gas_train_master
            ORDER BY max_flow DESC
            LIMIT 1
        """)
        row = cursor.fetchone()

    conn.close()

    if not row:
        raise ValueError(
            f"No suitable NG Gas Train found for required flow {required_flow_nm3hr}"
        )

    return {
        "inlet_nb": row[0],
        "outlet_nb": row[1],
        "min_flow": row[2],
        "max_flow": row[3],
        "price": row[4],
    }


# -------------------------------------------------
# DEBUG BLOCK
# -------------------------------------------------
if __name__ == "__main__":

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM gas_train_master")
    rows = cursor.fetchall()

    cursor.execute("PRAGMA table_info(gas_train_master)")
    columns = [col[1] for col in cursor.fetchall()]
    print(" | ".join(columns))
    print("-" * 100)

    for row in rows:
        print(" | ".join(str(v) for v in row))

    print(f"\nTotal rows: {len(rows)}")

    conn.close()