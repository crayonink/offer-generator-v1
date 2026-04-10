import sqlite3
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DB_PATH = os.path.join(BASE_DIR, "vlph.db")


# Burner size → matched HPU capacity (kW), per ENCON Film Burner spec
BURNER_TO_HPU_KW = {
    "ENCON 2A": 3,
    "ENCON 3A": 3,
    "ENCON 4A": 6,
    "ENCON 5A": 9,
    "ENCON 6A": 9,
    "ENCON 7A": 12,
}


# Variant → model prefix
VARIANT_PREFIX = {
    "Simplex":  "HPS",
    "Duplex 1": "HPD",
    "Duplex 2": "HPDD",
}


def select_hpu(burner_model: str, variant: str = "Duplex 1") -> dict:
    """
    Select Heating & Pumping Unit (HPU) for a given burner model + variant.

    Returns dict with model name and total cost (sum of all line items in
    hpu_master for that unit_kw + variant combination).
    """
    if variant not in VARIANT_PREFIX:
        raise ValueError(f"Invalid HPU variant '{variant}'. Use one of {list(VARIANT_PREFIX)}")

    unit_kw = BURNER_TO_HPU_KW.get(burner_model)
    if unit_kw is None:
        raise ValueError(f"No HPU mapping for burner '{burner_model}'")

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
    }
