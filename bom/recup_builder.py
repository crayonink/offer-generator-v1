"""Recuperator BOM builder (USK Exports model).

Takes the sized RecupResults from calculations/recup.py and assembles a
priced bill of materials. Rates live in vlph.db.recup_rates so they can
be tuned without a code push.

Returns a pandas DataFrame in the same shape as the other builders so
the existing /api/generate-quote pipeline can re-use it:
    columns = [MEDIA, ITEM NAME, REFERENCE, QTY, MAKE, UNIT PRICE, TOTAL]
With BOUGHT OUT ITEMS, ENCON ITEMS, GRAND TOTAL summary rows.
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

    # MS structural weights — Excel uses fabrication estimates derived from
    # geometry. We surface them as fixed kg figures for now (matches the
    # 19.08.2025 sheet values for the 1500 kW example).
    ms_outer_shell_kg     = 81.32
    ms_air_inlet_duct_kg  = 110.21
    ms_hot_outlet_duct_kg = 121.44
    ms_pipe_holding_kg    = 525.75
    ms_bottom_box_kg      = 160.93
    flanges_kg            = r('FLANGES_KG', 100.0)
    side_hood_kg          = r('SIDE_HOOD_MS_KG', 1500.0)

    ms_per_kg     = r('MS_PER_KG', 60.0)
    ss304_per_kg  = r('SS304_TUBE_PER_KG', 220.0)
    welding_pp    = r('WELDING_PER_PIPE', 32.0)
    bending_flat  = r('BENDING_FLAT', 9100.0)
    hole_fab      = r('HOLE_FABRICATION', 33600.0)
    thermo        = r('THERMOCOUPLE_TT', 8000.0)

    rows: list[tuple] = []

    # ── BOUGHT OUT ITEMS ────────────────────────────────────────────────
    rows.append(("MISC ITEMS", "Thermocouple with TT", "R Type", 1, "TEMPSENS", thermo, thermo))

    # ── ENCON ITEMS — Tubes ─────────────────────────────────────────────
    hot_cost  = round(results.weight_hot_bank_kg  * ss304_per_kg)
    cold_cost = round(results.weight_cold_bank_kg * ss304_per_kg)
    rows.append((
        "ENCON ITEMS", "SS304 ERW Tube — Hot Bank",
        f"{results.weight_hot_bank_kg:.0f} kg @ Rs.{ss304_per_kg:.0f}/kg",
        results.pipes_total, "ENCON", round(hot_cost / max(1, results.pipes_total)), hot_cost,
    ))
    rows.append((
        "ENCON ITEMS", "SS304 ERW Tube — Cold Bank",
        f"{results.weight_cold_bank_kg:.0f} kg @ Rs.{ss304_per_kg:.0f}/kg",
        results.pipes_total, "ENCON", round(cold_cost / max(1, results.pipes_total)), cold_cost,
    ))

    # ── ENCON ITEMS — MS fabrication ────────────────────────────────────
    def _ms_row(name: str, kg: float):
        cost = round(kg * ms_per_kg)
        rows.append(("ENCON ITEMS", name, f"{kg:.1f} kg @ Rs.{ms_per_kg:.0f}/kg",
                     1, "ENCON", cost, cost))

    _ms_row("MS Outer Shell",         ms_outer_shell_kg)
    _ms_row("MS Combustion Air Inlet Duct", ms_air_inlet_duct_kg)
    _ms_row("MS Hot Air Outlet Duct", ms_hot_outlet_duct_kg)
    _ms_row("Pipe Holding Plate",     ms_pipe_holding_kg)
    _ms_row("MS Bottom Box",          ms_bottom_box_kg)
    _ms_row("Machining Flanges",      flanges_kg)

    # Side hoods – flat cost per spec
    side_hood_cost = round(r('SIDE_HOOD_COST', 50000.0))
    rows.append(("ENCON ITEMS", "MS Side Hood (2 Nos)",
                 f"{side_hood_kg:.0f} kg total", 2, "ENCON",
                 round(side_hood_cost / 2), side_hood_cost))

    # Combustion air inlet assembly – flat cost
    cai_cost = round(r('COMBUSTION_AIR_INLET', 155978.7))
    rows.append(("ENCON ITEMS", "MS Combustion Air Inlet Assembly",
                 "Fabricated", 1, "ENCON", cai_cost, cai_cost))

    # Structural channels / angles – flat cost each (per metre stock)
    for key, name in (
        ('MS_CHANNEL_150x75x10', 'MS Channel 150 x 75 x 10'),
        ('MS_ANGLE_65x25',       'MS Angle 65 x 25'),
        ('MS_ANGLE_75x10',       'MS Angle 75 x 10'),
        ('MS_ANGLE_50x15',       'MS Angle 50 x 15'),
    ):
        cost = round(r(key))
        rows.append(("ENCON ITEMS", name, "per metre stock", 1, "ENCON", cost, cost))

    # Labour
    welding_cost = round(welding_pp * results.pipes_total)
    rows.append(("ENCON ITEMS", "Welding Rods",
                 f"4 rods/pipe x {results.pipes_total} pipes @ Rs.{welding_pp:.0f}/pipe",
                 1, "ENCON", welding_cost, welding_cost))
    rows.append(("ENCON ITEMS", "Pipe Bending",
                 "Flat — all pipes", 1, "ENCON",
                 round(bending_flat), round(bending_flat)))
    rows.append(("ENCON ITEMS", "Hole Fabrication", "Labour", 1, "ENCON",
                 round(hole_fab), round(hole_fab)))

    df = pd.DataFrame(rows, columns=["MEDIA", "ITEM NAME", "REFERENCE", "QTY",
                                      "MAKE", "UNIT PRICE", "TOTAL"])

    # Summary rows
    bought_total = float(df.loc[df["MEDIA"] != "ENCON ITEMS", "TOTAL"].sum())
    encon_total  = float(df.loc[df["MEDIA"] == "ENCON ITEMS", "TOTAL"].sum())

    summary = pd.DataFrame([
        ["", "BOUGHT OUT ITEMS", "", "", "", "", bought_total],
        ["", "ENCON ITEMS",      "", "", "", "", encon_total],
        ["", "GRAND TOTAL",      "", "", "", "", bought_total + encon_total],
    ], columns=df.columns)
    return pd.concat([df, summary], ignore_index=True)


def recup_summary(results: RecupResults, rates: Optional[dict] = None) -> dict:
    """Return the 4-line cost summary the form needs (bought / encon /
    grand / final) — Final = grand × conversion × (1 + designing) × (1 + negotiation).
    Note: caller can override any of these on the Step-3 panel."""
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
    final  = sale * (1 + desg) * (1 + nego)
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
