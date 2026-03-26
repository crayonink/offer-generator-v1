# bom/price_master.py

"""
Component price registry — reads from component_price_master table in vlph.db.
"""

import sqlite3
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.path.join(BASE_DIR, "vlph.db")


def get_price(item_name: str) -> float:
    """
    Fetch price for a component from the database.
    Raises ValueError if item is not found.
    """
    conn = sqlite3.connect(DB_PATH)
    try:
        row = conn.execute(
            "SELECT price FROM component_price_master WHERE item = ? LIMIT 1",
            (item_name,),
        ).fetchone()
    finally:
        conn.close()

    if row is None:
        raise ValueError(f"Missing price for item: '{item_name}'")

    return float(row[0])


def get_all_prices() -> dict:
    """Return all component prices as a dict {item: price}."""
    conn = sqlite3.connect(DB_PATH)
    try:
        rows = conn.execute(
            "SELECT item, price FROM component_price_master ORDER BY category, item"
        ).fetchall()
    finally:
        conn.close()
    return {item: price for item, price in rows}
