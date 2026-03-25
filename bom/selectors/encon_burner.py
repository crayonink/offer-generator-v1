import sqlite3

def select_encon_mg_burner(required_gas_flow_nm3hr: float) -> dict:
    """
    Select ENCON Gas Burner based on gas firing rate.
    Converts Nm3/hr → equivalent oil LPH.
    """

    # 🔥 Convert Gas Nm3/hr to equivalent Oil LPH
    equivalent_lph = required_gas_flow_nm3hr * 10500 / 8600

    conn = sqlite3.connect("vlph.db")
    cursor = conn.cursor()

    # -------------------------------------------------
    # 1️⃣ Select burner model using LPH range
    # -------------------------------------------------
    cursor.execute("""
        SELECT model
        FROM burner_selection_master
        WHERE ? BETWEEN min_firing_lph AND max_firing_lph
        LIMIT 1
    """, (equivalent_lph,))

    row = cursor.fetchone()

    if not row:
        conn.close()
        raise ValueError(
            f"No ENCON Gas burner available for "
            f"{required_gas_flow_nm3hr:.1f} Nm3/hr "
            f"(≈ {equivalent_lph:.1f} LPH)"
        )

    model = row[0]

    # -------------------------------------------------
    # 2️⃣ Fetch price from commercial table
    # -------------------------------------------------
    cursor.execute("""
        SELECT burner_price_inr
        FROM burner_master
        WHERE model = ?
    """, (model,))

    price_row = cursor.fetchone()
    conn.close()

    if not price_row:
        raise ValueError(f"Price not found for burner model {model}")

    return {
        "model": model,
        "input_nm3hr": required_gas_flow_nm3hr,
        "equivalent_lph": round(equivalent_lph, 2),
        "price": price_row[0],
    }