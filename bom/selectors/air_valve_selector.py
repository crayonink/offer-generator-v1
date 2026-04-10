import sqlite3


def select_motorized_control_valve(air_flow_nm3hr: float) -> dict:
    """
    Select Motorized Control Valve from DB
    """

    conn = sqlite3.connect("vlph.db")
    cursor = conn.cursor()

    cursor.execute("""
        SELECT nb, rated_flow, price
        FROM motorized_valve_master
        WHERE ? BETWEEN min_flow AND max_flow
        LIMIT 1
    """, (air_flow_nm3hr,))

    row = cursor.fetchone()
    conn.close()

    if not row:
        raise ValueError(
            f"No Motorized Control Valve found for {air_flow_nm3hr} Nm3/hr"
        )

    return {
        "nb": row[0],
        "flow_nm3hr": row[1],
        "price": row[2],
    }



def select_butterfly_valve(nb: int, vendor: str = "cair") -> dict:
    """
    Select Shut-Off Valve based on vendor.

    vendor = 'cair'      -> CAIR butterfly (SHUT OFF VALVE XXXNB (Butterfly))
    vendor = 'dembla'    -> DEMBLA ball  (SHUT OFF VALVE XXXNB)
    vendor = 'lt_lever'  -> L&T 2IWE4SL Lever (lt_butterfly_valve_master)
    vendor = 'lt_gear'   -> L&T 2IWE4SG Gear  (lt_butterfly_valve_master)

    Returns the smallest available NB >= requested.
    """
    vendor_lower = vendor.lower()
    conn = sqlite3.connect("vlph.db")

    # ── L&T butterfly valves ───────────────────────────────────────────────
    if vendor_lower in ("lt_lever", "lt_gear"):
        model = "2IWE4SL" if vendor_lower == "lt_lever" else "2IWE4SG"
        row = conn.execute("""
            SELECT nb, price, operation
            FROM lt_butterfly_valve_master
            WHERE model = ? AND nb >= ?
            ORDER BY nb ASC
            LIMIT 1
        """, (model, nb)).fetchone()
        conn.close()
        if not row:
            raise ValueError(f"No L&T {model} butterfly valve found for NB >= {nb}")
        return {
            "nb":    int(row[0]),
            "price": float(row[1]),
            "make":  f"L&T {model} ({row[2]})",
        }

    # ── CAIR / DEMBLA shut-off valves from component_price_master ──────────
    if vendor_lower == "dembla":
        like_pattern = "SHUT OFF VALVE %NB"
        company      = "DEMBLA"
    else:
        like_pattern = "SHUT OFF VALVE %NB (Butterfly)"
        company      = "CAIR"

    cursor = conn.cursor()
    cursor.execute("""
        SELECT item, price
        FROM component_price_master
        WHERE item LIKE ?
          AND company = ?
          AND CAST(SUBSTR(item, 16, 3) AS INTEGER) >= ?
        ORDER BY CAST(SUBSTR(item, 16, 3) AS INTEGER) ASC
        LIMIT 1
    """, (like_pattern, company, nb))

    row = cursor.fetchone()
    conn.close()

    if not row:
        raise ValueError(f"No Shut-Off Valve found for NB >= {nb} ({company})")

    item = row[0]                 # e.g. 'SHUT OFF VALVE 040NB (Butterfly)'
    nb_str = item[15:18]
    nb_val = int(nb_str) if nb_str.isdigit() else nb
    return {
        "nb":    nb_val,
        "price": float(row[1]),
        "make":  company,
    }