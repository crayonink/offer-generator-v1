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
    # 2️⃣ Fetch price from burner_pricelist_master
    #    burner_selection_master stores "ENCON G-4A" but
    #    pricelist stores "ENCON 4A" — strip the "G-" prefix
    # -------------------------------------------------
    pricelist_name = model.replace("G-", "")  # "ENCON G-4A" → "ENCON 4A"

    cursor.execute("""
        SELECT price
        FROM burner_pricelist_master
        WHERE burner_size = ?
          AND component = 'BURNER ALONE'
          AND section LIKE '%GAS%'
        LIMIT 1
    """, (pricelist_name,))

    price_row = cursor.fetchone()
    conn.close()

    if not price_row:
        raise ValueError(f"Price not found for burner model {model} (looked up as '{pricelist_name}')")

    return {
        "model": model,
        "input_nm3hr": required_gas_flow_nm3hr,
        "equivalent_lph": round(equivalent_lph, 2),
        "price": price_row[0],
    }