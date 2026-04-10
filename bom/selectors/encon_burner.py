import sqlite3

# Gas-fuel sub-types → all map to ENCON Gas Burner section
GAS_SUBTYPES = {"gas", "ng", "rlng", "lpg", "cog", "bg"}

# Oil-fuel sub-types → all map to IIP-ENCON Film Burner section
OIL_SUBTYPES = {"oil", "ldo", "fo", "hsd", "sko"}

# Map fuel category → burner pricelist section keyword
SECTION_BY_CATEGORY = {
    "gas":  "GAS",         # ENCON Gas Burner
    "oil":  "FILM",        # IIP-ENCON Film Burner (any oil sub-type)
    "dual": "DUAL FUEL",   # Dual Fuel Burner
}


def _resolve_category(fuel_type: str) -> str:
    """Map a specific fuel sub-type (ldo, fo, ng, rlng, ...) to its category."""
    f = fuel_type.lower()
    if f in OIL_SUBTYPES:
        return "oil"
    if f in GAS_SUBTYPES:
        return "gas"
    if f == "dual":
        return "dual"
    return "gas"


def select_encon_mg_burner(required_gas_flow_nm3hr: float, fuel_cv: float = 10500, fuel_type: str = "gas") -> dict:
    """
    Select ENCON Burner based on firing rate.
    fuel_type: 'gas' (default), 'oil', or 'dual' — picks the appropriate
    section in burner_pricelist_master.
    """

    # Convert Gas Nm3/hr to equivalent Oil LPH using actual fuel CV
    equivalent_lph = required_gas_flow_nm3hr * fuel_cv / 8600

    conn = sqlite3.connect("vlph.db")
    cursor = conn.cursor()

    # -------------------------------------------------
    # 1. Select burner model using LPH range
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
            f"No ENCON burner available for "
            f"{required_gas_flow_nm3hr:.1f} Nm3/hr "
            f"(~ {equivalent_lph:.1f} LPH)"
        )

    model = row[0]

    # Fetch price from burner_pricelist_master based on fuel category
    # ldo/fo/hsd/sko → oil → FILM section
    # ng/lpg/cog/bg/rlng → gas → GAS section
    # dual → DUAL FUEL section
    category = _resolve_category(fuel_type)
    section_keyword = SECTION_BY_CATEGORY[category]

    cursor.execute("""
        SELECT price
        FROM burner_pricelist_master
        WHERE burner_size = ?
          AND component = 'BURNER SET'
          AND section LIKE ?
        LIMIT 1
    """, (model, f"%{section_keyword}%"))

    price_row = cursor.fetchone()
    conn.close()

    if not price_row:
        raise ValueError(
            f"Price not found for burner {model} "
            f"(fuel_type={fuel_type}, section keyword='{section_keyword}')"
        )

    return {
        "model": model,
        "input_nm3hr": required_gas_flow_nm3hr,
        "equivalent_lph": round(equivalent_lph, 2),
        "price": price_row[0],
        "fuel_type": fuel_type,
    }