import sqlite3
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DB_PATH = os.path.join(BASE_DIR, "vlph.db")


def select_orifice_plate(nb: int) -> dict:
    """
    Select orifice plate from orifice_plate_master based on pipe NB.
    Picks exact NB match, or next size up if exact not available.
    """
    conn = sqlite3.connect(DB_PATH)
    row = conn.execute(
        "SELECT nb, total_price FROM orifice_plate_master WHERE nb >= ? ORDER BY nb ASC LIMIT 1",
        (nb,),
    ).fetchone()
    conn.close()

    if not row:
        raise ValueError(f"No Orifice Plate found for NB >= {nb}")

    return {
        "nb": row[0],
        "price": row[1],
    }
