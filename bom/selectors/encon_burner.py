import sqlite3

# Gas-fuel sub-types → all map to ENCON Gas Burner section
GAS_SUBTYPES = {"gas", "ng", "rlng", "lpg", "cog", "bg", "mg"}

# Oil-fuel sub-types → all map to IIP-ENCON Film Burner section
OIL_SUBTYPES = {"oil", "ldo", "fo", "lshs", "hsd", "sko", "hdo", "cfo"}

# Map fuel category → burner pricelist section keyword
SECTION_BY_CATEGORY = {
    "gas":  "GAS",         # ENCON Gas Burner
    "oil":  "FILM",        # IIP-ENCON Film Burner (any oil sub-type)
    "dual": "DUAL FUEL",   # Dual Fuel Burner
}


def _get_fuel_density(fuel_type: str) -> float:
    """Fetch fuel density from component_price_master.

    Units follow the DB convention: oil rows are kg/L, gas rows are kg/m3.
    Raises ValueError if the row is missing — we never want to silently
    fall back to a hardcoded default and corrupt downstream sizing math.
    """
    key = f"FUEL DENSITY {fuel_type.upper()}"
    conn = sqlite3.connect("vlph.db")
    row = conn.execute(
        "SELECT price FROM component_price_master WHERE item = ? LIMIT 1",
        (key,),
    ).fetchone()
    conn.close()
    if row is None:
        raise ValueError(f"No density row found for fuel '{fuel_type}' (key '{key}')")
    return float(row[0])


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

    # Convert flow to LPH for burner selection.
    # Oil fuels: burner-calc gives kg/hr → divide by own density to get l/hr.
    # Gas fuels: Nm3/hr × CV ÷ 10500 = kg/hr, then ÷ LDO density (from DB) = l/hr.
    category = _resolve_category(fuel_type)
    if category == "oil":
        density = _get_fuel_density(fuel_type)           # kg/L
        density_unit = "kg/ltr"
        equivalent_lph = required_gas_flow_nm3hr / density  # kg/hr → l/hr
    else:
        # Actual gas density (kg/m3 = kg/Nm3 at STP) — used for display and mass-flow calcs.
        density = _get_fuel_density(fuel_type)
        density_unit = "kg/m³"
        # Burner selection indexes by oil-equivalent LPH, so convert gas mass
        # flow to an oil equivalent using LDO density from the DB.
        equivalent_kghr = required_gas_flow_nm3hr * fuel_cv / 10500
        equivalent_lph = equivalent_kghr / _get_fuel_density("ldo")

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

    # Build a display-friendly model name that matches the ENCON pricelist
    # convention: oil stays as-is (ENCON 7A), gas becomes 'ENCON -G 7A',
    # dual becomes 'ENCON DUAL- 7A'.
    if category == "gas":
        display_model = model.replace("ENCON ", "ENCON -G ", 1)
    elif category == "dual":
        display_model = model.replace("ENCON ", "ENCON DUAL- ", 1)
    else:
        display_model = model  # oil — no prefix change

    return {
        "model": model,
        "display_model": display_model,
        "input_nm3hr": required_gas_flow_nm3hr,
        "equivalent_lph": round(equivalent_lph, 2),
        "fuel_density": density,
        "fuel_density_unit": density_unit,
        "price": price_row[0],
        "fuel_type": fuel_type,
        "burner_pressure_wg": burner_pressure_wg,
    }