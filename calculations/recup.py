"""Recuperator sizing + costing (USK Exports model).

Pure-logic module — no DB, no Flask. Mirrors the formulas in the
'Recuperator for Hardening' sheet from USK Exports_Recuperator_19.08.2025
so the web calculator produces the same numbers the engineer's
spreadsheet does.

Heat-transfer chain:
  1. Energy = connected_power_kw * 860         (kcal/hr)
  2. Fuel flow = Energy / CV                    (Nm³/hr)
  3. Combustion-air = fuel-specific × fuel flow (Nm³/hr — input or derived)
  4. Heat in preheated air = mass_air × Cp_air × dT_air (kcal/hr)
  5. Flue mass flow (mass_air ≈ same as combustion air mass + flue makeup
     — Excel uses ~1.2 × air for LPG combustion products; we let user
     supply mass_flue if they want)
  6. Final flue temp = Ti_flue - Q / (m_flue × Cp_flue)
  7. LMTD = ((Ti_flue - Tf_air) - (Tf_flue - Ti_air)) /
            ln((Ti_flue - Tf_air) / (Tf_flue - Ti_air))
  8. LMTD_corrected = LMTD / 1.2 (convective)
  9. Surface area = Q / (U × LMTD_corrected)
 10. Pipe count = Area / (π × dia × pipe_length)
 11. Round pipe count up to next row × column grid; each bank carries
     the same pipe count (hot bank + cold bank).
"""
from __future__ import annotations

import math
from dataclasses import dataclass


@dataclass
class RecupInputs:
    # Process parameters
    connected_power_kw: float
    fuel_cv_kcal_nm3:   float            # e.g. 21000 for LPG
    # Flue gas
    flue_flow_nm3hr:      float          # e.g. 1811 — derived in spreadsheet; let user override
    flue_mass_kghr:       float          # e.g. 2174 — derived in spreadsheet
    flue_temp_in_C:       float          # 800
    flue_temp_out_C:      float          # 421 (target — Excel solves for this)
    cp_flue_kcal_kgC:     float = 0.23
    # Combustion air to be preheated
    air_volume_nm3hr:     float = 1750.0
    air_temp_in_C:        float = 35.0
    air_temp_out_C:       float = 400.0
    cp_air_kcal_kgC:      float = 0.247
    # Recuperator geometry
    heat_transfer_coef:   float = 30.0   # kcal/m²-°C
    pipe_dia_mm:          float = 48.3
    pipe_thick_mm:        float = 2.77
    pipe_kg_per_m:        float = 3.16
    pipe_length_m_per_bank:float = 0.63
    bank_length_mm:       float = 696.1
    bank_width_mm:        float = 615.8
    bank_gap_mm:          float = 150.0
    # Forced row/column override (set 0 for auto-derive)
    pipes_in_row:         int = 0
    pipes_in_column:      int = 0


@dataclass
class RecupResults:
    # Energy
    energy_kcal_hr:        float
    fuel_flow_nm3hr:       float
    # Air heat
    air_mass_kg_hr:        float
    heat_required_kcal:    float
    # LMTD / area
    lmtd_C:                float
    lmtd_corrected_C:      float
    surface_area_m2:       float
    # Pipe count
    pipes_total_raw:       float          # math result before rounding
    pipes_total:           int            # rows x cols
    pipes_in_row:          int
    pipes_in_column:       int
    # Pipe weights (each bank gets the same count)
    weight_per_pipe_kg:    float
    weight_hot_bank_kg:    float
    weight_cold_bank_kg:   float
    weight_total_pipes_kg: float


def calculate_recup(inp: RecupInputs) -> RecupResults:
    if inp.connected_power_kw <= 0 or inp.fuel_cv_kcal_nm3 <= 0:
        raise ValueError("connected_power_kw and fuel_cv must be > 0")

    energy_kcal_hr = inp.connected_power_kw * 860.0
    fuel_flow_nm3hr = energy_kcal_hr / inp.fuel_cv_kcal_nm3

    # Heat needed to preheat the combustion air.
    # mass_air (kg/hr) = volume_nm3hr × 1.293 (kg/Nm³ at STP)
    air_mass_kg_hr = inp.air_volume_nm3hr * 1.293
    dT_air = inp.air_temp_out_C - inp.air_temp_in_C
    heat_required = air_mass_kg_hr * inp.cp_air_kcal_kgC * dT_air

    # LMTD (counter-flow): hot-in vs cold-out at one end, hot-out vs cold-in at the other.
    dT1 = inp.flue_temp_in_C  - inp.air_temp_out_C   # 800 - 400 = 400
    dT2 = inp.flue_temp_out_C - inp.air_temp_in_C    # 421 - 35  = 386
    if dT1 <= 0 or dT2 <= 0 or dT1 == dT2:
        # Degenerate / parallel flow fallback
        lmtd = (dT1 + dT2) / 2.0 if dT1 + dT2 > 0 else 1.0
    else:
        lmtd = (dT1 - dT2) / math.log(dT1 / dT2)
    # USK sheet divides by 1.2 for convective type.
    lmtd_corrected = lmtd / 1.2

    # Surface area required.
    surface_area_m2 = heat_required / (inp.heat_transfer_coef * lmtd_corrected)

    # Pipe count: surface area / (π × outer_dia × length_per_pipe)
    pipe_dia_m = inp.pipe_dia_mm / 1000.0
    pipe_total_length_per_pipe = inp.pipe_length_m_per_bank * 2  # hot + cold bank in series
    pipes_total_raw = surface_area_m2 / (math.pi * pipe_dia_m * pipe_total_length_per_pipe)

    # Row x column rounding: ceil to a 14×N grid (Excel uses 14 per row).
    if inp.pipes_in_row > 0 and inp.pipes_in_column > 0:
        rows_count = inp.pipes_in_row
        cols_count = inp.pipes_in_column
    else:
        rows_count = 14   # matches Excel example
        cols_count = max(1, math.ceil(pipes_total_raw / rows_count))
    pipes_total = rows_count * cols_count

    # Pipe weights
    weight_per_pipe = inp.pipe_kg_per_m * inp.pipe_length_m_per_bank
    weight_hot_bank = weight_per_pipe * pipes_total
    weight_cold_bank = weight_per_pipe * pipes_total
    weight_total_pipes = weight_hot_bank + weight_cold_bank

    return RecupResults(
        energy_kcal_hr=round(energy_kcal_hr, 2),
        fuel_flow_nm3hr=round(fuel_flow_nm3hr, 2),
        air_mass_kg_hr=round(air_mass_kg_hr, 2),
        heat_required_kcal=round(heat_required, 2),
        lmtd_C=round(lmtd, 2),
        lmtd_corrected_C=round(lmtd_corrected, 2),
        surface_area_m2=round(surface_area_m2, 2),
        pipes_total_raw=round(pipes_total_raw, 2),
        pipes_total=pipes_total,
        pipes_in_row=rows_count,
        pipes_in_column=cols_count,
        weight_per_pipe_kg=round(weight_per_pipe, 4),
        weight_hot_bank_kg=round(weight_hot_bank, 2),
        weight_cold_bank_kg=round(weight_cold_bank, 2),
        weight_total_pipes_kg=round(weight_total_pipes, 2),
    )
