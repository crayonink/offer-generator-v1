"""
HPU raw-materials / components pricelist.

The HPU BOM (hpu_master) stores a qty per line, but the RATE for every line is
now sourced from the pricelist (component_price_master) so that editing a rate
in the Price Master reprices every HPU that uses it — the same "linked rate"
model the burner tabs use.

Two kinds of line:
  * Raw material (unit = kg): mapped to one of ~10 canonical material types
    (M.S. Channel, H.R. Sheet 3mm, M.S. Plate 5mm, …). Many different sizes in
    the BOM collapse to a single per-kg rate.
  * Bought-out (Nos / Ltr / Mtr / g): each distinct item is its own catalogue
    SKU, priced per piece. The canonical rate is the HIGHER of the rates seen
    across variants (a per-variant "Simplex discount" collapses to one rate).

LABOUR CHARGE is not a material and stays a fixed line (its stored amount).

The catalogue rows live in component_price_master under HPU_* categories and are
inserted idempotently at start-up by seed_hpu_catalog(); see main.py.
"""

import re
import sqlite3

# Category names used for the HPU rows in component_price_master.
CAT_RAW    = "HPU Raw Material"
CAT_PIPE   = "HPU Pipe Fitting"
CAT_VALVE  = "HPU Valve & Filter"
CAT_ELEC   = "HPU Electrical & Rotating"
CAT_INST   = "HPU Instrumentation"
CAT_CONSUM = "HPU Consumable & Misc"
HPU_CATEGORIES = (CAT_RAW, CAT_PIPE, CAT_VALVE, CAT_ELEC, CAT_INST, CAT_CONSUM)

# Canonical raw materials: label -> per-kg rate (the max seen in the BOM).
RAW_MATERIALS = {
    "M.S. Channel":        51.4,
    "M.S. Pipe":           135.0,
    "M.S. Flat":           71.0,
    "M.S. Plate 3mm":      77.0,
    "M.S. Plate 5mm":      51.7,
    "M.S. Plate 16/20mm":  51.02,
    "H.R. Sheet 3mm":      77.0,
    "H.R. Sheet 5mm":      49.74,
    "Gasket":              120.0,
    "Nut / Bolt":          100.0,
}


def normalize(name: str) -> str:
    """Upper-cased, whitespace-collapsed key so 'HAMMER TON ' == 'HAMMER TON'."""
    return re.sub(r"\s+", " ", (name or "").strip().upper())


def is_labour(item: str) -> bool:
    return "LABOUR" in (item or "").upper()


def is_raw(unit: str) -> bool:
    return (unit or "").strip().lower() == "kg"


def raw_material_of(item: str):
    """Map a kg BOM line to its canonical raw-material label, or None."""
    u = normalize(item)
    if "CHANNEL" in u:
        return "M.S. Channel"
    if "FLAT" in u:
        return "M.S. Flat"
    if "PIPE" in u or "TUBE" in u:
        return "M.S. Pipe"
    if "GASKIT" in u or "GASKET" in u:
        return "Gasket"
    if "NUT" in u or "BOLT" in u:
        return "Nut / Bolt"
    if "HR SHEET" in u:
        return "H.R. Sheet 5mm" if re.search(r"X\s*5\s*MM", u) else "H.R. Sheet 3mm"
    if "PLATE" in u:
        if re.search(r"16\s*MM|16MM|20\s*MM|20MM|DIAX16|DIAX20", u):
            return "M.S. Plate 16/20mm"
        if re.search(r"X\s*5\s*MM|X5\s*MM|DIA\s*X\s*5|DIAX5|\b5\s*MM", u):
            return "M.S. Plate 5mm"
        if re.search(r"X\s*3\s*MM|DIAX3|\b3\s*MM", u):
            return "M.S. Plate 3mm"
        return "M.S. Plate 5mm"
    return None


def bought_out_group(item: str) -> str:
    """Category for a non-raw, non-labour BOM line."""
    u = normalize(item)
    if any(k in u for k in ("GAUGE", "THERMOSTAT", "INDICATOR", "TEMP.")):
        return CAT_INST
    if any(k in u for k in ("MOTOR", "OIL PUMP", "HEATER", "COUPLING")):
        return CAT_ELEC
    if any(k in u for k in ("FILTER", "GATE VALVE", "REGULATOR")):
        return CAT_VALVE
    if any(k in u for k in ("FLANGE", "ELBOW", "TEE", "BEND", "PLUG",
                            "SOCKET", "NIPPAL", "NIPPLE", "CROSS")):
        return CAT_PIPE
    return CAT_CONSUM


def build_catalog(conn: sqlite3.Connection):
    """Derive the full HPU catalogue from hpu_master.

    Returns a list of {item, unit, price, category} rows: the 10 canonical raw
    materials plus one row per distinct bought-out SKU (price = max rate seen).
    """
    conn.row_factory = sqlite3.Row
    rows = [dict(r) for r in conn.execute(
        "SELECT item, unit, rate FROM hpu_master")]

    catalog = []
    # Raw materials — fixed canonical rates.
    for label, rate in RAW_MATERIALS.items():
        catalog.append({"item": label, "unit": "kg",
                        "price": rate, "category": CAT_RAW})

    # Bought-out SKUs — max rate per normalized name.
    seen = {}
    for r in rows:
        if is_labour(r["item"]) or is_raw(r["unit"]):
            continue
        key = normalize(r["item"])
        rate = r["rate"] or 0.0
        if key not in seen or rate > seen[key]["price"]:
            seen[key] = {"item": key, "unit": (r["unit"] or "Nos").strip(),
                         "price": rate, "category": bought_out_group(r["item"])}
    catalog.extend(seen.values())
    return catalog


def load_rates(conn: sqlite3.Connection) -> dict:
    """Load {normalized item -> price} for all HPU catalogue rows."""
    q = ("SELECT item, price FROM component_price_master WHERE category IN (%s)"
         % ",".join("?" * len(HPU_CATEGORIES)))
    return {normalize(r[0]): (r[1] or 0.0)
            for r in conn.execute(q, HPU_CATEGORIES)}


def resolve_rate(item: str, unit: str, rates: dict):
    """Return (rate, source_label) for a BOM line, pulling from `rates`
    (a dict from load_rates). source_label is the catalogue SKU the rate came
    from (for the ◆ linked-rate tooltip); None for labour / unresolved."""
    if is_labour(item):
        return None, None
    if is_raw(unit):
        label = raw_material_of(item)
        if label is not None:
            return rates.get(normalize(label), RAW_MATERIALS.get(label, 0.0)), label
        # kg line that isn't a known raw material (e.g. a mislabelled unit) —
        # fall through to the bought-out catalogue by name.
    key = normalize(item)
    if key in rates:
        return rates[key], key
    return 0.0, None


def seed_hpu_catalog(conn: sqlite3.Connection) -> int:
    """Idempotently insert the HPU catalogue into component_price_master.
    Existing rows (matched by item name) are left untouched so live Price-Master
    edits persist. Returns the number of rows inserted."""
    catalog = build_catalog(conn)
    inserted = 0
    for row in catalog:
        exists = conn.execute(
            "SELECT 1 FROM component_price_master WHERE item = ? LIMIT 1",
            (row["item"],)).fetchone()
        if exists:
            continue
        conn.execute(
            "INSERT INTO component_price_master (item, category, unit, price, "
            "previous_price) VALUES (?,?,?,?,?)",
            (row["item"], row["category"], row["unit"], row["price"], row["price"]),
        )
        inserted += 1
    conn.commit()
    return inserted
