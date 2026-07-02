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
# Used as a fallback rate when the pricelist row is missing.
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

# Each canonical raw material is priced from ONE row in component_price_master
# (deduped — no HPU-specific twin). Most map to the pre-existing generic
# raw-material rows; two have no generic equivalent and are HPU-owned.
RAW_SOURCE = {
    "M.S. Channel":        "M.S. Channel",          # renamed from "M.S. Chanel"
    "M.S. Pipe":           "M.S. Pipe",
    "M.S. Flat":           "M.S. Flat",
    "M.S. Plate 5mm":      "M.S. Plate 5mm",
    "M.S. Plate 16/20mm":  "M.S. Plate 16mm*5mm",
    "H.R. Sheet 3mm":      "M.S. Sheet 3mm",
    "H.R. Sheet 5mm":      "M.S. Sheet 5mm",
    "Nut / Bolt":          "Hardware Bolt",
    "M.S. Plate 3mm":      "M.S. Plate 3mm",         # HPU-owned (no generic row)
    "Gasket":              "Gasket",                 # HPU-owned (no generic row)
}

# HPU-owned raw materials (no generic equivalent) — seeded under 'Raw Material'.
RAW_OWNED = {"M.S. Plate 3mm": 77.0, "Gasket": 120.0}

# One-time consolidation (higher rate wins) applied to the shared generic rows
# whose HPU BOM rate was higher than the pre-existing pricelist rate.
_RAW_BUMPS = {"M.S. Flat": 71.0, "M.S. Sheet 3mm": 77.0, "M.S. Plate 5mm": 51.7}

# HPU raw rows this module used to create that duplicate a generic row — removed
# during consolidation so each material has a single source.
_GENERIC_HOMED = ["M.S. Channel", "M.S. Plate 16/20mm", "H.R. Sheet 3mm",
                  "H.R. Sheet 5mm", "Nut / Bolt", "M.S. Pipe", "M.S. Flat",
                  "M.S. Plate 5mm"]

# HPU bought-out lines priced from a row OUTSIDE the HPU categories (clubbed
# into a shared vendor group). BOM item -> catalogue row it's priced from.
# The HPU pressure gauge is the small HGURU gauge, kept with the BAUMER/HGURU
# gauges in the Instrumentation group.
BOUGHT_SOURCE = {
    "PRESSURE GAUGE 0 -7 KG": "PRESSURE GAUGE SMALL (HGURU)",
}
# One-time reclass of those rows into their shared group (item/make/spec/cat).
_BOUGHT_RECLASS = [
    {"old": "PRESSURE GAUGE 0 -7 KG", "new": "PRESSURE GAUGE SMALL (HGURU)",
     "company": "HGURU", "spec": "Small", "category": "Instrumentation"},
]


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
    # Only the HPU-owned raw materials (no generic pricelist row) are seeded,
    # under 'Raw Material'. The rest are priced from the existing generic rows
    # (see RAW_SOURCE) so there are no duplicate raw-material rows.
    for label, rate in RAW_OWNED.items():
        catalog.append({"item": label, "unit": "kg",
                        "price": rate, "category": "Raw Material"})

    # Bought-out SKUs — max rate per normalized name. Items sourced from an
    # external group (BOUGHT_SOURCE) are not seeded as HPU rows.
    ext = {normalize(k) for k in BOUGHT_SOURCE}
    seen = {}
    for r in rows:
        if is_labour(r["item"]) or is_raw(r["unit"]):
            continue
        key = normalize(r["item"])
        if key in ext:
            continue
        rate = r["rate"] or 0.0
        if key not in seen or rate > seen[key]["price"]:
            seen[key] = {"item": key, "unit": (r["unit"] or "Nos").strip(),
                         "price": rate, "category": bought_out_group(r["item"])}
    catalog.extend(seen.values())
    return catalog


def load_rates(conn: sqlite3.Connection) -> dict:
    """Load {normalized item -> price} for the HPU catalogue rows PLUS the
    generic raw-material rows that HPU lines are priced from (RAW_SOURCE)."""
    names = tuple(dict.fromkeys(list(RAW_SOURCE.values()) + list(BOUGHT_SOURCE.values())))
    q = ("SELECT item, price FROM component_price_master "
         "WHERE category IN (%s) OR item IN (%s)"
         % (",".join("?" * len(HPU_CATEGORIES)), ",".join("?" * len(names))))
    return {normalize(r[0]): (r[1] or 0.0)
            for r in conn.execute(q, (*HPU_CATEGORIES, *names))}


def resolve_rate(item: str, unit: str, rates: dict):
    """Return (rate, source_label) for a BOM line, pulling from `rates`
    (a dict from load_rates). source_label is the pricelist row the rate came
    from (for the ◆ linked-rate tooltip); None for labour / unresolved."""
    if is_labour(item):
        return None, None
    if is_raw(unit):
        label = raw_material_of(item)
        if label is not None:
            src = RAW_SOURCE.get(label, label)
            rate = rates.get(normalize(src))
            if rate is None:
                rate = RAW_MATERIALS.get(label, 0.0)
            return rate, src
        # kg line that isn't a known raw material (e.g. a mislabelled unit) —
        # fall through to the bought-out catalogue by name.
    key = normalize(item)
    # Externally-sourced bought-out lines (clubbed into a shared vendor group).
    ext = {normalize(k): v for k, v in BOUGHT_SOURCE.items()}
    if key in ext:
        src = ext[key]
        return rates.get(normalize(src), 0.0), src
    if key in rates:
        return rates[key], key
    return 0.0, None


def consolidate_raw_materials(conn: sqlite3.Connection) -> int:
    """One-time de-duplication of HPU raw materials against the generic rows.

    Idempotent and self-guarding (via the presence of 'M.S. Chanel' or a
    left-over HPU raw twin). On the run that finds work to do it:
      1. deletes the HPU raw rows that duplicate a generic row,
      2. collapses 'M.S. Chanel' / 'M.S. Channel' into a single
         'M.S. Channel' row (correct spelling) at the higher rate,
      3. bumps the shared rows whose HPU rate was higher (higher wins),
      4. moves the HPU-owned raws (Plate 3mm, Gasket) into 'Raw Material'.
    Returns 1 if it acted, else 0.
    """
    cur = conn.cursor()
    ph = ",".join("?" * len(_GENERIC_HOMED))
    need = cur.execute(
        "SELECT 1 FROM component_price_master WHERE item='M.S. Chanel' "
        "OR (category='HPU Raw Material' AND item IN (%s)) LIMIT 1" % ph,
        _GENERIC_HOMED,
    ).fetchone()
    if not need:
        return 0

    # 1. drop HPU raw rows that duplicate a generic row.
    cur.execute(
        "DELETE FROM component_price_master WHERE category='HPU Raw Material' "
        "AND item IN (%s)" % ph, _GENERIC_HOMED)

    # 2. collapse Chanel/Channel into one correctly-spelled row at the max rate.
    prices = [r[0] for r in cur.execute(
        "SELECT price FROM component_price_master WHERE item IN "
        "('M.S. Chanel','M.S. Channel') AND price IS NOT NULL")]
    rate = max(prices + [RAW_MATERIALS["M.S. Channel"]])
    cur.execute("DELETE FROM component_price_master WHERE item IN ('M.S. Chanel','M.S. Channel')")
    cur.execute(
        "INSERT INTO component_price_master (item, category, unit, price, previous_price) "
        "VALUES ('M.S. Channel','Raw Material','kg',?,?)", (rate, rate))

    # 3. higher-rate-wins on the shared generic rows.
    for name, r in _RAW_BUMPS.items():
        cur.execute(
            "UPDATE component_price_master SET previous_price=price, price=? "
            "WHERE item=? AND price<?", (r, name, r))

    # 4. keep the HPU-owned raws but file them under 'Raw Material'.
    cur.execute(
        "UPDATE component_price_master SET category='Raw Material' "
        "WHERE category='HPU Raw Material' AND item IN ('M.S. Plate 3mm','Gasket')")

    conn.commit()
    return 1


# Cosmetic item renames applied to BOTH hpu_master and component_price_master
# (kept in sync so the resolver still matches by exact name). old -> new.
# Stored UPPER-CASE to match the other HPU bought-out rows (the viewer
# title-cases names for display, so this shows as "Temperature Gauge").
_HPU_RENAMES = {
    "TEMP. GAUGE 0 -150 *C": "TEMPERATURE GAUGE",
}


def apply_hpu_renames(conn: sqlite3.Connection) -> int:
    """Rename HPU items in both the BOM and the pricelist. Idempotent (the old
    name is gone after the first run). Returns rows touched."""
    n = 0
    for old, new in _HPU_RENAMES.items():
        for tbl in ("component_price_master", "hpu_master"):
            try:
                n += conn.execute(
                    f"UPDATE {tbl} SET item=? WHERE item=?", (new, old)).rowcount
            except sqlite3.OperationalError:
                pass
    conn.commit()
    return n


def reclassify_bought_sources(conn: sqlite3.Connection) -> int:
    """Move HPU bought-out rows that belong in a shared vendor group (e.g. the
    HPU pressure gauge -> small HGURU gauge in the Instrumentation group).
    Idempotent: once the old name is gone it's a no-op. Returns rows touched."""
    n = 0
    for m in _BOUGHT_RECLASS:
        old = conn.execute(
            "SELECT rowid FROM component_price_master WHERE item=? COLLATE NOCASE",
            (m["old"],)).fetchone()
        if not old:
            continue
        tgt = conn.execute(
            "SELECT rowid FROM component_price_master WHERE item=? COLLATE NOCASE",
            (m["new"],)).fetchone()
        if tgt and tgt[0] != old[0]:
            # target already present — drop the stray HPU row to avoid a dupe.
            conn.execute("DELETE FROM component_price_master WHERE rowid=?", (old[0],))
        else:
            conn.execute(
                "UPDATE component_price_master SET item=?, company=?, "
                "specification=?, category=? WHERE rowid=?",
                (m["new"], m["company"], m["spec"], m["category"], old[0]))
        n += 1
    conn.commit()
    return n


def seed_hpu_catalog(conn: sqlite3.Connection) -> int:
    """Idempotently insert the HPU catalogue into component_price_master.
    Existing rows (matched by item name) are left untouched so live Price-Master
    edits persist. Returns the number of rows inserted."""
    catalog = build_catalog(conn)
    inserted = 0
    for row in catalog:
        exists = conn.execute(
            "SELECT 1 FROM component_price_master WHERE item = ? COLLATE NOCASE LIMIT 1",
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
