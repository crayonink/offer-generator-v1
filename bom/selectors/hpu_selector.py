import sqlite3
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DB_PATH = os.path.join(BASE_DIR, "vlph.db")


# HPU heater capacity (kW) → max flow rate it can deliver (Litres/hour)
# From ENCON Heating & Pumping Unit datasheet.
HPU_FLOW_BY_KW = [
    (3,   40),
    (6,   80),
    (9,   120),
    (12,  160),
    (16,  200),
    (20,  250),
    (24,  300),
    (30,  400),
    (36,  500),
    (48,  600),
]


# Variant → model prefix
VARIANT_PREFIX = {
    "Simplex":  "HPS",
    "Duplex 1": "HPD",
    "Duplex 2": "HPDD",
}

# Pumping-Unit variant → model prefix (used for LSHS / FO where the oil is
# pre-heated separately, so no heating element is needed in the HPU package).
PU_VARIANT_PREFIX = {
    "Simplex":  "PUS",
    "Duplex 1": "PUD",
    "Duplex 2": "PUDD",
}

# Oils that use a standalone Pumping Unit (heating handled elsewhere).
PUMPING_UNIT_ONLY_FUELS = {"ldo", "lshs", "fo"}


def _hpu_kw_for_lph(required_lph: float) -> int:
    """Pick the smallest HPU kW whose max flow rate >= required LPH."""
    for kw, max_flow in HPU_FLOW_BY_KW:
        if max_flow >= required_lph:
            return kw
    # Above largest available — return the biggest
    return HPU_FLOW_BY_KW[-1][0]


def select_hpu(required_lph: float, variant: str = "Duplex 1") -> dict:
    """
    Select Heating & Pumping Unit (HPU) sized to actual oil firing rate (LPH).

    required_lph : actual fuel flow rate the HPU must deliver
    variant      : 'Simplex' | 'Duplex 1' | 'Duplex 2'

    Returns dict with HPU model name + total cost (sum of all hpu_master line
    items for that unit_kw + variant combination).
    """
    if variant not in VARIANT_PREFIX:
        raise ValueError(f"Invalid HPU variant '{variant}'. Use one of {list(VARIANT_PREFIX)}")

    unit_kw = _hpu_kw_for_lph(required_lph)

    conn = sqlite3.connect(DB_PATH)
    row = conn.execute(
        "SELECT SUM(amount) FROM hpu_master WHERE unit_kw = ? AND variant = ?",
        (unit_kw, variant),
    ).fetchone()
    conn.close()

    total = float(row[0]) if row and row[0] is not None else 0
    if total == 0:
        raise ValueError(f"No HPU pricing rows for {unit_kw} kW / {variant}")

    model = f"{VARIANT_PREFIX[variant]}-{unit_kw}"
    return {
        "model":   model,
        "unit_kw": unit_kw,
        "variant": variant,
        "price":   total,
        "unit_type": "HPU",
        "label":    "Heating and Pumping Unit (HPU)",
    }


def select_pumping_unit(required_lph: float, variant: str = "Duplex 1") -> dict:
    """
    Select a standalone Pumping Unit (no heating element) for heavy oils like
    LSHS/FO where the fuel is pre-heated separately. Price comes from
    pumping_unit_price (sell_price, already includes margin).
    """
    if variant not in PU_VARIANT_PREFIX:
        raise ValueError(f"Invalid Pumping Unit variant '{variant}'. Use one of {list(PU_VARIANT_PREFIX)}")

    unit_kw = _hpu_kw_for_lph(required_lph)

    conn = sqlite3.connect(DB_PATH)
    row = conn.execute(
        "SELECT sell_price FROM pumping_unit_price WHERE unit_kw = ? AND variant = ?",
        (unit_kw, variant),
    ).fetchone()
    conn.close()

    if not row or row[0] is None:
        raise ValueError(f"No pumping-unit pricing for {unit_kw} kW / {variant}")

    price = float(row[0])
    model = f"{PU_VARIANT_PREFIX[variant]}-{unit_kw}"
    return {
        "model":   model,
        "unit_kw": unit_kw,
        "variant": variant,
        "price":   price,
        "unit_type": "PU",
        "label":    "Pumping Unit",
    }
