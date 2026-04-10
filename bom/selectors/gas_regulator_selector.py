import sqlite3
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DB_PATH = os.path.join(BASE_DIR, "vlph.db")


def select_gas_regulator(nb: int, category: str = "Standard 5 Bar") -> dict:
    """
    Select MADAS gas pressure regulator from gas_regulator_master.

    nb       : nominal bore in mm. Picks exact match, else next size up.
    category : 'Standard 5 Bar' (default), 'Reinforced 5 Bar',
               'Pilot 5 Bar', 'OPSO/UPSO 5 Bar', 'Standard 2 Bar'.

    For oil-based fuels the pilot line uses the smallest available regulator
    (DN025 / 25 NB) — pass nb=20 or nb=25 and you'll get the DN025 row.
    """
    conn = sqlite3.connect(DB_PATH)
    row = conn.execute(
        """
        SELECT nb, part_code, p2_range, list_price, connection
        FROM gas_regulator_master
        WHERE category = ? AND nb >= ?
        ORDER BY nb ASC
        LIMIT 1
        """,
        (category, nb),
    ).fetchone()
    conn.close()

    if not row:
        raise ValueError(
            f"No MADAS gas regulator found for NB >= {nb} in category '{category}'"
        )

    return {
        "nb":         row[0],
        "part_code":  row[1],
        "p2_range":   row[2],
        "price":      float(row[3]),
        "connection": row[4],
        "category":   category,
    }
