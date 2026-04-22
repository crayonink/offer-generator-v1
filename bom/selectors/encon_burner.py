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

    # Burner selection now matches by heat (kcal/hr) — density-free for gas.
    # For oil, `required_gas_flow_nm3hr` is actually kg/hr (legacy arg name);
    # for gas it is Nm3/hr. Either way, flow × CV = kcal/hr.
    category = _resolve_category(fuel_type)
    required_heat_kcal_hr = required_gas_flow_nm3hr * fuel_cv

    if category == "oil":
        density = _get_fuel_density(fuel_type)           # kg/L, still needed for pump LPH
        density_unit = "kg/ltr"
        equivalent_lph = required_gas_flow_nm3hr / density  # kg/hr → l/hr (pump sizing)
    elif category == "dual":
        # Dual-fuel burner sized by total heat; density is fuel-specific and
        # doesn't belong on the combined burner record.
        density = 0
        density_unit = ""
        equivalent_lph = 0
    else:
        density = _get_fuel_density(fuel_type)           # kg/m3, display only
        density_unit = "kg/m³"
        equivalent_lph = 0  # not meaningful for gas

    conn = sqlite3.connect("vlph.db")
    cursor = conn.cursor()

    # -------------------------------------------------
    # 1. Select burner model using heat-rating range (kcal/hr).
    #    Pick the smallest model whose max range covers the required heat.
    # -------------------------------------------------
    cursor.execute("""
        SELECT model
        FROM burner_selection_master
        WHERE pressure_wg = ?
          AND ? BETWEEN min_firing_kcal_hr AND max_firing_kcal_hr
        ORDER BY max_firing_kcal_hr ASC
        LIMIT 1
    """, (burner_pressure_wg, required_heat_kcal_hr))

    row = cursor.fetchone()

    if not row:
        conn.close()
        raise ValueError(
            f"No ENCON burner available for "
            f"{required_heat_kcal_hr:.0f} kcal/hr at {burner_pressure_wg}\" w.g."
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