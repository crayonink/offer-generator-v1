"""
bom/ladle_params.py

Reads ladle sizing parameters directly from vertical_master / horizontal_master
DB tables (populated when the pricebook is uploaded).

When the pricebook is uploaded:
  - MS STRUCTURE cost  = ms_kg × M.S.Plate_rate × 2.1  (live from component_price_master)
  - CERAMIC FIBER cost = rolls × price_per_roll    (reads exact amount from DB)
  - CONTROL PANEL cost = reads exact amount from DB
  - SWIRLING / PIPELINE / TROLLEY = reads exact amount from DB
  - HPU kW = parsed from "H & P UNIT 3KW..." description in DB

Falls back to hardcoded tables if DB has no data.
"""

import re
import sqlite3
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.path.join(BASE_DIR, "vlph.db")

MS_PLATE_MARKUP = 2.1   # Excel formula: ms_kg × plate_rate × 2.1


def _live_ms_plate_rate() -> float:
    """
    Fetch M.S. Plate rate from component_price_master.
    Looks for an item whose name contains 'M.S. Plate' and '16mm'.
    Falls back to MS_STRUCTURE_RATE / MS_PLATE_MARKUP if not found.
    """
    try:
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute(
            """SELECT rate FROM component_price_master
               WHERE LOWER(item_name) LIKE '%m.s. plate%16mm%'
               LIMIT 1"""
        )
        row = c.fetchone()
        conn.close()
        if row and row[0]:
            return float(row[0])
    except Exception:
        pass
    return MS_STRUCTURE_RATE / MS_PLATE_MARKUP  # fallback ≈ 72


# ─── fallback tables (used only when DB has no data) ──────────────────────────
_VLPH_FALLBACK = [
    # (max_tons, ms_kg, cf_rolls, hpu_kw, pipeline_swirling, panel)
    (10,  1096, 9,  3,  124929, 105000),
    (14,  1130, 8,  6,  139047, 105000),
    (15,  1215, 9,  6,  147000, 115500),
    (20,  1300, 10, 6,  158372, 115500),
    (30,  1354, 11, 12, 168000, 147000),
    (40,  1773, 13, 12, 178500, 147000),
    (60,  1800, 19, 20, 189000, 147000),
]

_HLPH_FALLBACK = [
    # (max_tons, ms_kg, cf_rolls, hpu_kw, pipeline, trolley, panel)
    (10,  1416, 13, 3,  115500, 129600, 126000),
    (30,  1683, 19, 12, 135000, 168000, 222600),
]

MS_STRUCTURE_RATE = 151.2  # Rs/kg fabricated — fallback constant


# ─── helpers ──────────────────────────────────────────────────────────────────

def _parse_tons(model: str):
    """
    Parse upper-bound ladle capacity from model name.
    "DESCRIPTION( VERTICAL 16-20 TON LPS)" → 20
    "DESCRIPTION( VERTICAL 10 TON LPS)"    → 10
    "10 TON HORIZONTAL LADLE PREHEATER"    → 10
    """
    # Range: "16-20 TON" or "50 TO 60 TON"
    m = re.search(r'(\d+)\s*(?:TO|-)\s*(\d+)\s+TON', model, re.I)
    if m:
        return int(m.group(2))
    # Single: "10 TON"
    m = re.search(r'(\d+)\s+TON', model, re.I)
    if m:
        return int(m.group(1))
    return None


def _parse_qty_num(qty_str):
    if not qty_str:
        return None
    m = re.search(r'(\d+(?:\.\d+)?)', str(qty_str))
    return float(m.group(1)) if m else None


def _parse_hpu_kw(text: str):
    """
    "C. H & P UNIT 3KW-1NO,(HPD-3)"       → 3
    "C. H & P UNIT 6 KW  1NO .(HPD-6)"    → 6
    "C. H & P UNIT 20 KW  1NO .(HPD-20)"  → 20
    """
    m = re.search(r'(\d+)\s*KW', text, re.I)
    return int(m.group(1)) if m else None


def _get_ladle_rows(ladle_tons: float, master_table: str):
    """
    Query vertical_master or horizontal_master for the best-matching model.
    Returns (max_tons, model_name, rows_dict) or (None, None, None).
    rows_dict = {particular: {"qty_str": ..., "amount": ...}}
    """
    try:
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute(f"SELECT DISTINCT model FROM {master_table}")
        models = [r[0] for r in c.fetchall()]
    except Exception:
        return None, None, None

    if not models:
        conn.close()
        return None, None, None

    candidates = []
    for model in models:
        t = _parse_tons(model)
        if t is not None:
            candidates.append((t, model))

    if not candidates:
        conn.close()
        return None, None, None

    candidates.sort()

    selected = None
    for t, model in candidates:
        if ladle_tons <= t:
            selected = (t, model)
            break
    if not selected:
        selected = candidates[-1]

    max_t, model_name = selected
    c.execute(
        f"SELECT particular, qty, amount FROM {master_table} WHERE model = ?",
        (model_name,),
    )
    rows = {}
    for particular, qty, amount in c.fetchall():
        rows[particular] = {"qty_str": qty, "amount": amount or 0.0}

    conn.close()
    return max_t, model_name, rows


# ─── VLPH ─────────────────────────────────────────────────────────────────────

def get_vlph_params(ladle_tons: float) -> dict:
    """
    Return sizing parameters for a VLPH of given ladle capacity.
    Reads from vertical_master DB; falls back to hardcoded table if empty.
    """
    max_t, model_name, rows = _get_ladle_rows(ladle_tons, "vertical_master")

    if rows is None:
        return _vlph_fallback(ladle_tons)

    # MS STRUCTURE — compute live: ms_kg × plate_rate × 2.1
    ms = rows.get("MS STRUCTURE", {})
    ms_kg       = int(_parse_qty_num(ms.get("qty_str")) or 0)
    plate_rate  = _live_ms_plate_rate()
    ms_rate     = round(plate_rate * MS_PLATE_MARKUP, 2)
    ms_cost     = round(ms_kg * ms_rate, 2) if ms_kg else ms.get("amount", 0.0)

    # CERAMIC FIBER
    cf = rows.get("CERAMIC FIBER", {})
    cf_rolls = int(_parse_qty_num(cf.get("qty_str")) or 0)

    # CONTROL PANEL
    panel_cost = rows.get("CONTROL PANEL", {}).get("amount", 0.0)

    # SWIRLING MECH + PIPELINE
    pipeline_cost = 0.0
    for key, val in rows.items():
        if "SWIRLING" in key.upper() or (
            "PIPELINE" in key.upper() and "TROLLEY" not in key.upper()
        ):
            pipeline_cost = val["amount"]
            break

    # HPU kW — parse from "H & P UNIT {kw}KW..." particular
    hpu_kw = None
    for key in rows:
        if "H & P UNIT" in key.upper() or "H&P UNIT" in key.upper() or "H &P UNIT" in key.upper():
            hpu_kw = _parse_hpu_kw(key)
            break
    if not hpu_kw:
        hpu_kw = _hpu_kw_fallback(max_t, is_hlph=False)

    return {
        "ladle_tons":             ladle_tons,
        "ms_structure_kg":        ms_kg,
        "ms_structure_rate":      ms_rate,
        "ms_structure_cost":      round(ms_cost, 2),
        "ceramic_rolls":          cf_rolls,
        "hpu_kw":                 hpu_kw,
        "pipeline_swirling_cost": round(pipeline_cost, 2),
        "control_panel_cost":     round(panel_cost, 2),
    }


# ─── HLPH ─────────────────────────────────────────────────────────────────────

def get_hlph_params(ladle_tons: float) -> dict:
    """
    Return sizing parameters for an HLPH of given ladle capacity.
    Reads from horizontal_master DB; falls back to hardcoded table if empty.
    """
    max_t, model_name, rows = _get_ladle_rows(ladle_tons, "horizontal_master")

    if rows is None:
        return _hlph_fallback(ladle_tons)

    # MS STRUCTURE — compute live: ms_kg × plate_rate × 2.1
    ms = rows.get("MS STRUCTURE", {})
    ms_kg       = int(_parse_qty_num(ms.get("qty_str")) or 0)
    plate_rate  = _live_ms_plate_rate()
    ms_rate     = round(plate_rate * MS_PLATE_MARKUP, 2)
    ms_cost     = round(ms_kg * ms_rate, 2) if ms_kg else ms.get("amount", 0.0)

    # CERAMIC FIBER
    cf = rows.get("CERAMIC FIBER", {})
    cf_rolls = int(_parse_qty_num(cf.get("qty_str")) or 0)

    # CONTROL PANEL
    panel_cost = rows.get("CONTROL PANEL", {}).get("amount", 0.0)

    # PIPELINE (no trolley)
    pipeline_cost = 0.0
    for key, val in rows.items():
        if "PIPELINE" in key.upper() and "TROLLEY" not in key.upper():
            pipeline_cost = val["amount"]
            break

    # TROLLEY DRIVE
    trolley_cost = 0.0
    for key, val in rows.items():
        if "TROLLEY" in key.upper():
            trolley_cost = val["amount"]
            break

    # HPU kW
    hpu_kw = None
    for key in rows:
        if "H & P UNIT" in key.upper() or "H&P UNIT" in key.upper() or "H &P UNIT" in key.upper():
            hpu_kw = _parse_hpu_kw(key)
            break
    if not hpu_kw:
        hpu_kw = _hpu_kw_fallback(max_t, is_hlph=True)

    return {
        "ladle_tons":         ladle_tons,
        "ms_structure_kg":    ms_kg,
        "ms_structure_rate":  ms_rate,
        "ms_structure_cost":  round(ms_cost, 2),
        "ceramic_rolls":      cf_rolls,
        "hpu_kw":             hpu_kw,
        "pipeline_cost":      round(pipeline_cost, 2),
        "trolley_drive_cost": round(trolley_cost, 2),
        "control_panel_cost": round(panel_cost, 2),
    }


# ─── fallback helpers ─────────────────────────────────────────────────────────

def _hpu_kw_fallback(max_t: int, is_hlph: bool) -> int:
    table = {10: 3, 14: 6, 15: 6, 20: 6, 30: 12, 40: 12, 60: 20}
    return table.get(max_t, 6)


def _vlph_fallback(ladle_tons: float) -> dict:
    plate_rate = _live_ms_plate_rate()
    ms_rate = round(plate_rate * MS_PLATE_MARKUP, 2)
    for max_t, ms_kg, cf_rolls, hpu_kw, pipeline, panel in _VLPH_FALLBACK:
        if ladle_tons <= max_t:
            return {
                "ladle_tons": ladle_tons,
                "ms_structure_kg": ms_kg,
                "ms_structure_rate": ms_rate,
                "ms_structure_cost": round(ms_kg * ms_rate, 2),
                "ceramic_rolls": cf_rolls,
                "hpu_kw": hpu_kw,
                "pipeline_swirling_cost": pipeline,
                "control_panel_cost": panel,
            }
    max_t, ms_kg, cf_rolls, hpu_kw, pipeline, panel = _VLPH_FALLBACK[-1]
    return {
        "ladle_tons": ladle_tons,
        "ms_structure_kg": ms_kg,
        "ms_structure_rate": ms_rate,
        "ms_structure_cost": round(ms_kg * ms_rate, 2),
        "ceramic_rolls": cf_rolls,
        "hpu_kw": hpu_kw,
        "pipeline_swirling_cost": pipeline,
        "control_panel_cost": panel,
    }


def _hlph_fallback(ladle_tons: float) -> dict:
    plate_rate = _live_ms_plate_rate()
    ms_rate = round(plate_rate * MS_PLATE_MARKUP, 2)
    for max_t, ms_kg, cf_rolls, hpu_kw, pipeline, trolley, panel in _HLPH_FALLBACK:
        if ladle_tons <= max_t:
            return {
                "ladle_tons": ladle_tons,
                "ms_structure_kg": ms_kg,
                "ms_structure_rate": ms_rate,
                "ms_structure_cost": round(ms_kg * ms_rate, 2),
                "ceramic_rolls": cf_rolls,
                "hpu_kw": hpu_kw,
                "pipeline_cost": pipeline,
                "trolley_drive_cost": trolley,
                "control_panel_cost": panel,
            }
    max_t, ms_kg, cf_rolls, hpu_kw, pipeline, trolley, panel = _HLPH_FALLBACK[-1]
    return {
        "ladle_tons": ladle_tons,
        "ms_structure_kg": ms_kg,
        "ms_structure_rate": ms_rate,
        "ms_structure_cost": round(ms_kg * ms_rate, 2),
        "ceramic_rolls": cf_rolls,
        "hpu_kw": hpu_kw,
        "pipeline_cost": pipeline,
        "trolley_drive_cost": trolley,
        "control_panel_cost": panel,
    }


# ─── smoke test ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("=== VLPH ===")
    for t in [10, 14, 15, 20, 30, 40, 60]:
        p = get_vlph_params(t)
        print(f"  {t}T  MS={p['ms_structure_cost']:>10,.0f}  "
              f"CF={p['ceramic_rolls']}rolls  HPU={p['hpu_kw']}kW  "
              f"Pipeline={p['pipeline_swirling_cost']:>9,.0f}  "
              f"Panel={p['control_panel_cost']:>8,.0f}")
    print()
    print("=== HLPH ===")
    for t in [10, 30]:
        p = get_hlph_params(t)
        print(f"  {t}T  MS={p['ms_structure_cost']:>10,.0f}  "
              f"CF={p['ceramic_rolls']}rolls  HPU={p['hpu_kw']}kW  "
              f"Pipeline={p['pipeline_cost']:>9,.0f}  "
              f"Trolley={p['trolley_drive_cost']:>9,.0f}  "
              f"Panel={p['control_panel_cost']:>8,.0f}")
