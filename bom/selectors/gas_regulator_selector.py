import sqlite3
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DB_PATH = os.path.join(BASE_DIR, "vlph.db")

GAS_REGULATOR_DISCOUNT = 0.45   # 45% off MADAS list price


def select_gas_regulator(nb: int, category: str = "Standard 5 Bar") -> dict:
    """
    Select MADAS gas pressure regulator from gas_regulator_master.

    nb       : nominal bore in mm. Picks exact match, else next size up.
    category : 'Standard 5 Bar' (default), 'Reinforced 5 Bar',
               'Pilot 5 Bar', 'OPSO/UPSO 5 Bar', 'Standard 2 Bar'.

    Returned `price` has the 45% MADAS discount already applied
    (list × 0.55). The original list price is also returned for reference.
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

    list_price = float(row[3])
    part_code  = row[1]
    # Fallback: 45% off the master list price.
    net_price  = list_price * (1 - GAS_REGULATOR_DISCOUNT)

    # Prefer the editable pricelist (component_price_master) price, matched by
    # part code, so pricelist edits flow straight into the BOM (the stored
    # pricelist price is already the net figure). Falls back to list × 0.55.
    try:
        _c = sqlite3.connect(DB_PATH)
        _r = _c.execute(
            "SELECT price FROM component_price_master "
            "WHERE category LIKE 'Gas Regulator%' AND item LIKE ? "
            "AND price IS NOT NULL LIMIT 1",
            ("%" + str(part_code),),
        ).fetchone()
        _c.close()
        if _r and _r[0] is not None:
            net_price = float(_r[0])
    except Exception:
        pass

    return {
        "nb":         row[0],
        "part_code":  part_code,
        "p2_range":   row[2],
        "list_price": list_price,
        "price":      round(net_price, 2),
        "connection": row[4],
        "category":   category,
    }
