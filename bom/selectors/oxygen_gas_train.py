# bom/selectors/oxygen_gas_train.py
"""
NIRMAL Oxygen Skid Selector & Master Table Manager.
Pricelist source: NIRMAL OXYGEN SKID PRICE LIST - P34040.xlsx

Price list:
1 | Skid -1 | 50-60 Nm3/hr   | 7.78L  (Rs. 7,78,000)
2 | Skid -2 | 60-150 Nm3/hr  | 8.81L  (Rs. 8,81,000)
3 | Skid -3 | 150-250 Nm3/hr | 9.57L  (Rs. 9,57,000)
4 | Skid -4 | 250-600 Nm3/hr | 10.91L (Rs. 10,91,000)
5 | Skid -5 | 600-1000 Nm3/hr| 13.38L (Rs. 13,38,000)
"""

import os
import sqlite3

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DB_PATH = os.environ.get("VLPH_DB_PATH") or os.path.join(BASE_DIR, "vlph.db")

TABLE_NAME = "nirmal_oxygen_skid_master"

NIRMAL_OXYGEN_SKID_DATA = [
    (1, "Skid -1", "50-60 Nm3/hr", 0.0, 60.0, 7.78, 778000.0, "NIRMAL"),
    (2, "Skid -2", "60-150 Nm3/hr", 60.1, 150.0, 8.81, 881000.0, "NIRMAL"),
    (3, "Skid -3", "150-250 Nm3/hr", 150.1, 250.0, 9.57, 957000.0, "NIRMAL"),
    (4, "Skid -4", "250-600 Nm3/hr", 250.1, 600.0, 10.91, 1091000.0, "NIRMAL"),
    (5, "Skid -5", "600-1000 Nm3/hr", 600.1, 1000.0, 13.38, 1338000.0, "NIRMAL"),
]


def seed_nirmal_oxygen_skids(conn=None):
    """Seed nirmal_oxygen_skid_master table and update component_price_master."""
    should_close = False
    if conn is None:
        conn = sqlite3.connect(DB_PATH)
        should_close = True

    cursor = conn.cursor()

    # Create table if not exists
    cursor.execute(f"""
        CREATE TABLE IF NOT EXISTS {TABLE_NAME} (
            sr_no INTEGER PRIMARY KEY,
            tag_no TEXT NOT NULL,
            flow_cap TEXT NOT NULL,
            min_flow REAL NOT NULL,
            max_flow REAL NOT NULL,
            price_lakhs REAL NOT NULL,
            price_inr REAL NOT NULL,
            make TEXT NOT NULL DEFAULT 'NIRMAL'
        )
    """)

    # Populate/update rows
    for sr_no, tag_no, flow_cap, min_flow, max_flow, price_lakhs, price_inr, make in NIRMAL_OXYGEN_SKID_DATA:
        cursor.execute(f"""
            INSERT INTO {TABLE_NAME} (sr_no, tag_no, flow_cap, min_flow, max_flow, price_lakhs, price_inr, make)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(sr_no) DO UPDATE SET
                tag_no = excluded.tag_no,
                flow_cap = excluded.flow_cap,
                min_flow = excluded.min_flow,
                max_flow = excluded.max_flow,
                price_lakhs = excluded.price_lakhs,
                price_inr = excluded.price_inr,
                make = excluded.make
        """, (sr_no, tag_no, flow_cap, min_flow, max_flow, price_lakhs, price_inr, make))

        # Also register in component_price_master for rate sheets
        item_name = f"NIRMAL Oxygen {tag_no} ({flow_cap})"
        existing = cursor.execute("SELECT item FROM component_price_master WHERE item=?", (item_name,)).fetchone()
        if existing:
            cursor.execute("""
                UPDATE component_price_master
                SET category = 'Oxygen Gas Train', price = ?, company = 'NIRMAL'
                WHERE item = ?
            """, (price_inr, item_name))
        else:
            cursor.execute("""
                INSERT INTO component_price_master (item, category, unit, price, company)
                VALUES (?, 'Oxygen Gas Train', 'No', ?, 'NIRMAL')
            """, (item_name, price_inr))

    conn.commit()

    if should_close:
        conn.close()


def select_oxygen_gas_train(required_flow_nm3hr: float) -> dict:
    """
    Select NIRMAL Oxygen Gas Train Skid based on required O2 flow (Nm3/hr).
    """
    seed_nirmal_oxygen_skids()

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    # Query matching range
    cursor.execute(f"""
        SELECT sr_no, tag_no, flow_cap, min_flow, max_flow, price_lakhs, price_inr, make
        FROM {TABLE_NAME}
        WHERE ? >= min_flow AND ? <= max_flow
        ORDER BY sr_no ASC
        LIMIT 1
    """, (required_flow_nm3hr, required_flow_nm3hr))

    row = cursor.fetchone()

    # Fallback for flow below min or above max
    if not row:
        if required_flow_nm3hr <= 60.0:
            cursor.execute(f"SELECT sr_no, tag_no, flow_cap, min_flow, max_flow, price_lakhs, price_inr, make FROM {TABLE_NAME} ORDER BY sr_no ASC LIMIT 1")
        else:
            cursor.execute(f"SELECT sr_no, tag_no, flow_cap, min_flow, max_flow, price_lakhs, price_inr, make FROM {TABLE_NAME} ORDER BY sr_no DESC LIMIT 1")
        row = cursor.fetchone()

    conn.close()

    if not row:
        raise ValueError(f"No suitable NIRMAL Oxygen Skid found for required flow {required_flow_nm3hr}")

    sr_no, tag_no, flow_cap, min_flow, max_flow, price_lakhs, price_inr, make = row

    # Check component_price_master for user price overrides
    price = price_inr
    item_name = f"NIRMAL Oxygen {tag_no} ({flow_cap})"
    try:
        conn_c = sqlite3.connect(DB_PATH)
        res = conn_c.execute("SELECT price FROM component_price_master WHERE item=? LIMIT 1", (item_name,)).fetchone()
        conn_c.close()
        if res and res[0] is not None:
            price = float(res[0])
    except Exception:
        pass

    return {
        "sr_no": sr_no,
        "tag_no": tag_no,
        "flow_cap": flow_cap,
        "min_flow": min_flow,
        "max_flow": max_flow,
        "price": price,
        "price_lakhs": price_lakhs,
        "make": make,
        "model": f"NIRMAL Oxygen {tag_no} ({flow_cap})",
        "inlet_nb": tag_no,
        "outlet_nb": flow_cap,
    }


if __name__ == "__main__":
    seed_nirmal_oxygen_skids()
    print("NIRMAL Oxygen Skid Table successfully seeded in DB!")
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute(f"SELECT * FROM {TABLE_NAME}").fetchall()
    for r in rows:
        print(r)
    conn.close()
