import sqlite3
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DB_PATH = os.path.join(BASE_DIR, "vlph.db")


def select_rotary_joint(nb: int) -> dict:
    """
    Select Rotary Joint based on NB
    """

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.execute("""
        SELECT nb, price
        FROM rotary_joint_master
        WHERE nb >= ?
        ORDER BY nb ASC
        LIMIT 1
    """, (nb,))

    row = cursor.fetchone()
    conn.close()

    if not row:
        raise ValueError(
            f"No Rotary Joint found for NB >= {nb}"
        )

    return {
        "nb": row[0],
        "price": row[1],
    }