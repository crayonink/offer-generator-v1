"""Recuperator BOM builder — costing structure ported from the
Recuperator_Excel.xlsx 'Recuperator for Hardening' sheet.

Excel cost lines (F40..F64) and how they map here:

  F40  Price of All the pipes      = E28*F37 + E35*F38             -> SS304 hot + cold rows
  F48  Cost for MS Side hood 2 nos = flat Rs 50,000                -> 'MS Side Hood (Fabrication)' row
  F49  Cost of MS Combustion Air Inlet
        = (E41+E42+E43+E44+E45+E46+E47) * MS_rate                  -> 'MS Combustion Air Inlet Assembly' row
       (includes outer shell, ducts, holding plate, bottom box,
        machining flanges, AND the 1500 kg side hood weight)
  F50  MS Channel 150x75x10    = 10 m * 17 kg/m * MS_rate          -> MS Channel row
  F51  MS Angle 65x25          = 25 m * 8.8 kg/m * MS_rate         -> MS Angle row
  F52  MS Angle 75x10          = 9 m * 10 kg/m * MS_rate           -> MS Angle row
  F53  MS Angle 50x15          = 4.5 m * 15 kg/m * MS_rate         -> MS Angle row
  F55  Bending of Pipes        = (rows + cols) * 350               -> Pipe Bending row
  F56  Welding Rod             = rows * cols * 4 * 8               -> Welding Rods row
  F57  Hole Fabrication        = (rows * cols) * 2 * 100           -> Hole Fabrication row
  F58  Thermocouple with tt    = flat Rs 8000                      -> Thermocouple (MISC ITEMS) row

  F60  Sale  = F59 * 1.8 (conversion)
  F62  Designing  = F61 * 0.1
  F63  Negotiation = F61 * 0.1
  F64  Final = F61 + F62 + F63  (additive, not compound)

All formulas use rates that live in vlph.db.recup_rates so they can be
retuned from the /pricelist UI without a code push. Items not present
in recup_rates fall back to the Excel defaults.
"""
from __future__ import annotations

import math
import os
import sqlite3
from typing import Optional

import pandas as pd

from calculations.recup import RecupInputs, RecupResults, calculate_recup


_BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_DB_PATH = os.path.join(_BASE_DIR, "vlph.db")

# Stock-metal items used in the recuperator support frame. Each entry is
# (display_name, length_rate_key, kg_per_m_rate_key). Both numbers live in
# recup_rates so a planner can re-spec the support frame from /pricelist
# without touching code.
_STOCK_STEEL_SPECS = [
    ("MS Channel 150 x 75 x 10", "STOCK_CHANNEL_150x75x10_LEN_M", "STOCK_CHANNEL_150x75x10_KG_M"),
    ("MS Angle 65 x 25",         "STOCK_ANGLE_65x25_LEN_M",       "STOCK_ANGLE_65x25_KG_M"),
    ("MS Angle 75 x 10",         "STOCK_ANGLE_75x10_LEN_M",       "STOCK_ANGLE_75x10_KG_M"),
    ("MS Angle 50 x 15",         "STOCK_ANGLE_50x15_LEN_M",       "STOCK_ANGLE_50x15_KG_M"),
]


def _load_rates() -> dict:
    """Read every key/value out of recup_rates into a flat dict."""
    conn = sqlite3.connect(_DB_PATH)
    rates = {k: v for k, v in conn.execute(
        "SELECT key, value FROM recup_rates"
    ).fetchall()}
    conn.close()
    return rates


def build_recup_df(results: RecupResults, rates: Optional[dict] = None) -> pd.DataFrame:
    """Build the recuperator BOM DataFrame from sized results."""
    if rates is None:
        rates = _load_rates()

    def r(key: str, default: float = 0.0) -> float:
        return float(rates.get(key, default))

    ms_per_kg     = r('MS_PER_KG',         60.0)   # raw stock metal (channels/angles)
    ms_fab_per_kg = r('MS_FABRICATION_PER_KG', 70.0)  # fabricated MS (CAI assembly)
    flanges_kg    = r('FLANGES_KG',        100.0)
    # Tube material switch — SS304 ERW or MS ERW. Rate + label follow the
    # user's choice in RecupResults.tube_material; rates live in recup_rates
    # so they can be retuned from the /pricelist Recup Rates tab.
    tube_mat = (getattr(results, 'tube_material', 'SS') or 'SS').upper()
    if tube_mat == 'MS':
        tube_rate  = r('MS_TUBE_PER_KG',    70.0)
        tube_label = 'MS ERW Tube'
        tube_spec  = 'MS ERW'
    else:
        tube_rate  = r('SS304_TUBE_PER_KG', 250.0)
        tube_label = 'SS304 ERW Tube'
        tube_spec  = 'SS304 ERW'
    # Side hood weight is now a per-quote input on RecupResults — falls
    # back to the recup_rates default if results doesn't carry one.
    side_hood_kg  = float(getattr(results, 'side_hood_kg', 0)
                          or r('SIDE_HOOD_MS_KG', 1500.0))
    side_hood_fab = r('SIDE_HOOD_COST',    50000.0)
    thermo_cost   = r('THERMOCOUPLE_TT',   8000.0)

    # Excel uses literal constants for bending / welding / hole-fab rates.
    bending_per_unit = r('BENDING_PER_UNIT', 350.0)   # F55 multiplier
    rods_per_pipe    = int(r('RODS_PER_PIPE', 4))     # F56
    welding_per_rod  = r('WELDING_PER_ROD', 8.0)      # F56
    holes_per_pipe   = int(r('HOLES_PER_PIPE', 2))    # F57
    hole_fab_per_hole = r('HOLE_FAB_PER_HOLE', 100.0)  # F57

    rows_count = results.pipes_in_row
    cols_count = results.pipes_in_column
    n_total    = results.pipes_total
    n_per_bank = max(1, results.pipes_per_bank)

    rows: list[tuple] = []

    # ── BOUGHT OUT ITEMS ────────────────────────────────────────────────
    # F58: Thermocouple with TT — flat Rs 8000
    rows.append(("MISC ITEMS", "Thermocouple with TT", "R Type",
                 1, "TEMPSENS", thermo_cost, thermo_cost))

    # ── ENCON — Tubes (F40 broken into hot + cold) ─────────────────────
    hot_total  = round(results.weight_hot_bank_kg  * tube_rate, 2)
    cold_total = round(results.weight_cold_bank_kg * tube_rate, 2)
    rows.append((
        "ENCON ITEMS", f"{tube_label} — Hot Bank",
        f"{results.weight_hot_bank_kg:.2f} kg @ Rs.{tube_rate:.0f}/kg",
        n_per_bank, "ENCON",
        round(hot_total / n_per_bank, 2), hot_total,
    ))
    rows.append((
        "ENCON ITEMS", f"{tube_label} — Cold Bank",
        f"{results.weight_cold_bank_kg:.2f} kg @ Rs.{tube_rate:.0f}/kg",
        n_per_bank, "ENCON",
        round(cold_total / n_per_bank, 2), cold_total,
    ))

    # ── ENCON — MS Side Hood fabrication (F48, flat Rs 50,000) ─────────
    # The 1500 kg of side hood material is rolled into F49 below; F48 is
    # the additional forming / welding charge.
    rows.append((
        "ENCON ITEMS", "MS Side Hood (2 Nos) — Fabrication",
        "Forming + welding charge",
        2, "ENCON", round(side_hood_fab / 2, 2), side_hood_fab,
    ))

    # ── ENCON — MS Combustion Air Inlet Assembly (F49) ─────────────────
    # F49 = (E41+E42+E43+E44+E45+E46+E47) * MS_rate
    # E41..E45 come from the calc (geometry-derived); E46/E47 are
    # constants from recup_rates (flanges_kg, side_hood_kg).
    ms_total_kg = (
        results.ms_outer_shell_kg
        + results.ms_air_inlet_duct_kg
        + results.ms_hot_outlet_duct_kg
        + results.ms_pipe_holding_kg
        + results.ms_bottom_box_kg
        + flanges_kg
        + side_hood_kg
    )
    cai_cost = round(ms_total_kg * ms_fab_per_kg, 2)
    rows.append((
        "ENCON ITEMS", "MS Combustion Air Inlet Assembly",
        f"{ms_total_kg:.2f} kg @ Rs.{ms_fab_per_kg:.0f}/kg "
        f"(shell {results.ms_outer_shell_kg:.0f} + inlet {results.ms_air_inlet_duct_kg:.0f} "
        f"+ outlet {results.ms_hot_outlet_duct_kg:.0f} + holding {results.ms_pipe_holding_kg:.0f} "
        f"+ box {results.ms_bottom_box_kg:.0f} + flanges {flanges_kg:.0f} + hood {side_hood_kg:.0f})",
        1, "ENCON", cai_cost, cai_cost,
    ))

    # ── ENCON — Stock structural metal (F50..F53). Length + kg/m for each
    # item come from recup_rates so the spec is editable from /pricelist.
    for name, len_key, kgm_key in _STOCK_STEEL_SPECS:
        length_m = r(len_key)
        kg_per_m = r(kgm_key)
        if length_m <= 0 or kg_per_m <= 0:
            continue
        cost = round(length_m * kg_per_m * ms_per_kg, 2)
        rows.append((
            "ENCON ITEMS", name,
            f"{length_m:g} m × {kg_per_m:g} kg/m",
            1, "ENCON", cost, cost,
        ))

    # ── ENCON — Labour / fabrication (F55, F56, F57) ──────────────────
    bending = round((rows_count + cols_count) * bending_per_unit, 2)
    welding = round(n_total * rods_per_pipe * welding_per_rod, 2)
    holefab = round(n_total * holes_per_pipe * hole_fab_per_hole, 2)
    rows.append((
        "ENCON ITEMS", "Pipe Bending",
        f"{rows_count} + {cols_count} lines",
        1, "ENCON", bending, bending,
    ))
    rows.append((
        "ENCON ITEMS", "Welding Rods",
        f"{n_total * rods_per_pipe:,} rods",
        1, "ENCON", welding, welding,
    ))
    rows.append((
        "ENCON ITEMS", "Hole Fabrication",
        f"{n_total * holes_per_pipe:,} holes",
        1, "ENCON", holefab, holefab,
    ))

    df = pd.DataFrame(rows, columns=["MEDIA", "ITEM NAME", "REFERENCE", "QTY",
                                      "MAKE", "UNIT PRICE", "TOTAL"])

    # ── Summary rows ───────────────────────────────────────────────────
    bought_total = float(df.loc[df["MEDIA"] != "ENCON ITEMS", "TOTAL"].sum())
    encon_total  = float(df.loc[df["MEDIA"] == "ENCON ITEMS", "TOTAL"].sum())

    summary = pd.DataFrame([
        ["", "BOUGHT OUT ITEMS", "", "", "", "", round(bought_total, 2)],
        ["", "ENCON ITEMS",      "", "", "", "", round(encon_total,  2)],
        ["", "GRAND TOTAL",      "", "", "", "", round(bought_total + encon_total, 2)],
    ], columns=df.columns)
    return pd.concat([df, summary], ignore_index=True)


def recup_summary(results: RecupResults, rates: Optional[dict] = None) -> dict:
    """4-line cost summary using Excel's additive markup chain:
        Sale  = Grand * Conversion
        Final = Sale * (1 + Designing + Negotiation)
    (matches F60..F64 in the Excel — NOT compound). The Step-3 panel
    overrides any of these per quote."""
    if rates is None:
        rates = _load_rates()
    df = build_recup_df(results, rates)
    bought = float(df.loc[df["ITEM NAME"] == "BOUGHT OUT ITEMS", "TOTAL"].iloc[0])
    encon  = float(df.loc[df["ITEM NAME"] == "ENCON ITEMS",      "TOTAL"].iloc[0])
    grand  = float(df.loc[df["ITEM NAME"] == "GRAND TOTAL",      "TOTAL"].iloc[0])
    conv   = float(rates.get('MARKUP_CONVERSION', 1.8))
    desg   = float(rates.get('MARKUP_DESIGNING', 0.10))
    nego   = float(rates.get('MARKUP_NEGOTIATION', 0.10))
    sale   = grand * conv
    final  = sale * (1 + desg + nego)   # additive (Excel F64 = F61+F62+F63)
    return {
        'bought_out_total': round(bought, 2),
        'encon_total':      round(encon, 2),
        'grand_total':      round(grand, 2),
        'markup_conversion': conv,
        'sale_after_markup': round(sale, 2),
        'designing_pct':    desg * 100,
        'negotiation_pct':  nego * 100,
        'final_total':      round(math.ceil(final / 1000) * 1000, 0),
    }
