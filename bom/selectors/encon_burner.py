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


def select_encon_mg_burner(required_gas_flow_nm3hr: float, fuel_cv: float = 10500, fuel_type: str = "gas", burner_pressure_wg: int = 24) -> dict:
    """
    Select ENCON Burner based on firing rate.

    fuel_type           : 'gas' (default), 'oil', or 'dual' — drives pricelist section.
    burner_pressure_wg  : 24 or 36 (inches w.g.) — drives firing rate range lookup.
    """

    # Convert flow to LPH for selection.
    # For oil fuels the burner-calc output is already in LPH (CV is in kcal/L),
    # so no conversion is needed. For gas fuels we convert Nm3/hr to oil-equivalent
    # LPH using a reference oil CV of 8600 kcal/L.
    category = _resolve_category(fuel_type)
    if category == "oil":
        equivalent_lph = required_gas_flow_nm3hr   # already LPH
    else:
        equivalent_lph = required_gas_flow_nm3hr * fuel_cv / 8600

    conn = sqlite3.connect("vlph.db")
    cursor = conn.cursor()

    # -------------------------------------------------
    # 1. Select burner model using LPH range at the given pressure.
    #    Pick the smallest model whose max range is >= required LPH so the
    #    burner is properly sized (not the first one that overlaps).
    # -------------------------------------------------
    cursor.execute("""
        SELECT model
        FROM burner_selection_master
        WHERE pressure_wg = ?
          AND ? BETWEEN min_firing_lph AND max_firing_lph
        ORDER BY max_firing_lph ASC
        LIMIT 1
    """, (burner_pressure_wg, equivalent_lph))

    row = cursor.fetchone()

    if not row:
        conn.close()
        raise ValueError(
            f"No ENCON burner available for "
            f"{required_gas_flow_nm3hr:.1f} Nm3/hr "
            f"(~ {equivalent_lph:.1f} LPH) at {burner_pressure_wg}\" w.g."
        )

    model = row[0]

    # Fetch price from burner_pricelist_master based on fuel category
    # ldo/fo/hsd/sko → oil → FILM section
    # ng/lpg/cog/bg/rlng → gas → GAS section
    # dual → DUAL FUEL section
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
        "burner_pressure_wg": burner_pressure_wg,
    }