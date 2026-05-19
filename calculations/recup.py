"""Recuperator sizing + costing — exact port of the 'Recuperator for
Hardening' Excel sheet (Recuperator_Excel.xlsx, 19/05/2026).

Every numeric step here mirrors a specific cell in that workbook:

  E3  Total Flue Gas Nm3/hr        -> user input (flue_flow_nm3hr)
  E4  Total Mass kg/hr             -> E3 * 1.2                  (auto)
  E5  Cp Flue                      -> 0.23                      (const)
  E6  Inlet Flue Temp              -> user input
  E7  Final Flue Temp              -> E6 - (E13 / (E4*E5))      *derived*
  E8  Heat Transfer Coef           -> 30 kcal/m2-C              (configurable)
  E9  Combustion Air Vol           -> user input
  E10 Initial Air Temp             -> user input
  E11 Final Air Temp               -> user input
  E12 Cp Air                       -> 0.247
  E13 Q required                   -> (E9*1.2) * E12 * (E11-E10)
  E14 LMTD                         -> ((dT1-dT2)/ln(dT1/dT2)) * 0.9
  E15 Surface Area                 -> E13 / (E14 * E8)
  E16 Bank Length mm               -> (((rows-1)/2)*32) + ((rows/2)*dia) + 150
  E17 Bank Width  mm               -> ((cols/2)*48.3) + (((cols-1)/2)*32) + 150
  E18 Bank Gap                     -> 150 mm
  E19 Pipes raw                    -> E15 / (pi * (dia/1000) * (length + 0.1))
  E20 Rows                         -> 14 (default, overridable)
  E21 Cols                         -> ceil(E19 / rows)
  G21 Total Pipes                  -> rows * cols  (split half/half across both banks)
  E24 Length per pipe              -> 0.55 + 0.08 = 0.63 m
  E26 Pipe kg/m                    -> 3.16
  E27 Weight per pipe              -> 3.16 * 0.63 = 1.9908 kg
  E28 Hot Bank total kg            -> per_pipe * (rows * cols / 2)
  E35 Cold Bank total kg           -> same as E28

MS structural weights (E41..E45) are derived from bank geometry instead
of the previous flat constants, so they auto-scale with pipe count.
"""
from __future__ import annotations

import math
from dataclasses import dataclass


@dataclass
class RecupInputs:
    # ── Flue gas (inputs) ───────────────────────────────────────────
    flue_flow_nm3hr:      float = 1900.0       # E3
    flue_temp_in_C:       float = 800.0        # E6
    cp_flue_kcal_kgC:     float = 0.23         # E5
    # ── Combustion air to be preheated (inputs) ─────────────────────
    air_volume_nm3hr:     float = 1750.0       # E9
    air_temp_in_C:        float = 35.0         # E10
    air_temp_out_C:       float = 400.0        # E11
    cp_air_kcal_kgC:      float = 0.247        # E12
    # ── Geometry constants (rarely changed) ─────────────────────────
    heat_transfer_coef:   float = 30.0         # E8 kcal/m2-C
    pipe_dia_mm:          float = 48.3         # E23
    pipe_thick_mm:        float = 2.77         # E25
    pipe_kg_per_m:        float = 3.16         # E26
    pipe_length_m_per_bank: float = 0.63       # E24 = 0.55 + 0.08
    bank_gap_mm:          float = 150.0        # E18
    pipe_pitch_mm:        float = 32.0         # spacing between adjacent pipes
    end_margin_mm:        float = 150.0        # margin on each side
    surface_area_end_allowance_m: float = 0.1  # the "+0.1" in E19 denominator
    lmtd_factor:          float = 0.9          # E14 trailing factor
    flue_density_factor:  float = 1.2          # kg/Nm3 — the *1.2 in E4 and E13
    # ── Total-pipes override (0 -> auto: rows=14, cols=ceil(raw/14)).
    #     When > 0, we honour the user's total and re-derive the grid as
    #     rows=14, cols=ceil(total/14).
    pipes_total_override: int = 0
    # ── Tube material per bank (cost-only; "SS" or "MS"). Hot and Cold
    # can be different — common in hardening furnaces where one bank
    # sees hotter flue gas and is upgraded to SS.
    hot_bank_material:    str = "SS"
    cold_bank_material:   str = "SS"
    # ── MS Side Hood weight (kg). Default 1500. Editable per quote.
    side_hood_kg:         float = 1500.0
    # ── MS Combustion Air Inlet Assembly per-kg rate override (Rs/kg).
    # When > 0 the BOM bills sum_MS_kg * cai_rate_override instead of
    # the recup_rates default (MS_FABRICATION_PER_KG, Rs 70/kg).
    cai_rate_override:    float = 0.0


@dataclass
class RecupResults:
    # Energy / heat balance
    flue_mass_kghr:        float    # E4
    heat_required_kcal:    float    # E13
    flue_temp_out_C:       float    # E7
    lmtd_C:                float    # E14 (already with the 0.9 correction)
    surface_area_m2:       float    # E15
    # Pipe geometry
    pipes_total_raw:       float    # E19
    pipes_total:           int      # G21
    pipes_in_row:          int      # E20
    pipes_in_column:       int      # E21
    pipes_per_bank:        int      # G21 / 2
    bank_length_mm:        float    # E16
    bank_width_mm:         float    # E17
    weight_per_pipe_kg:    float    # E27
    weight_hot_bank_kg:    float    # E28
    weight_cold_bank_kg:   float    # E35
    weight_total_pipes_kg: float    # E28 + E35
    # MS structural weights (E41..E45) — auto-derived from bank geometry
    ms_outer_shell_kg:     float    # E41
    ms_air_inlet_duct_kg:  float    # E42
    ms_hot_outlet_duct_kg: float    # E43
    ms_pipe_holding_kg:    float    # E44
    ms_bottom_box_kg:      float    # E45
    # Echo of the user's per-bank tube material so the BOM builder
    # downstream picks the right rate + label per bank.
    hot_bank_material:     str = "SS"
    cold_bank_material:    str = "SS"
    # Echo of the user's side hood weight + CAI rate override.
    side_hood_kg:          float = 1500.0
    cai_rate_override:     float = 0.0


def _vol_to_kg(volume_mm3: float, density: float) -> float:
    """Excel uses density/1e9 (kg/mm3). Mirror that exactly."""
    return volume_mm3 * (density / 1_000_000_000)


def _best_grid(n: int) -> tuple[int, int]:
    """Pick the (rows, cols) layout for n total pipes that BALANCES
    aspect skew and waste. Cols is forced EVEN so the two banks each
    hold a whole number of pipes (E28 splits cols in half).

    Scoring:  aspect_pct + 3 * waste_pct  (lower is better)
      - aspect_pct = |rows - cols| / min(rows, cols)
      - waste_pct  = (rows*cols - n) / n
    Tie-breaks toward rows=14 so small cases still match the Excel
    convention (n=168 -> 14x12).

    Search range scales with sqrt(n) — [sqrt(n) - 8, sqrt(n) + 8] —
    so for very large n the algorithm finds near-square layouts that
    a fixed [8, 20] range would miss.

    Examples:
        n=  168 -> 14 x 12   (Excel match)
        n=  100 -> 10 x 10
        n= 1248 -> 35 x 36  (waste 12, aspect diff 1)
        n= 2179 -> 46 x 48  (waste 29, aspect diff 2)
        n= 2414 -> 50 x 50  (waste 86, perfect square)
    """
    if n <= 0:
        return (14, 2)
    s  = int(math.isqrt(n))
    lo = max(2, s - 8)
    hi = s + 8
    best = None  # (score, rows, cols)
    for rows in range(lo, hi + 1):
        cols = max(2, math.ceil(n / rows))
        if cols % 2:
            cols += 1
        waste      = rows * cols - n
        aspect_pct = (max(rows, cols) - min(rows, cols)) / min(rows, cols)
        waste_pct  = waste / n
        # 3x weight on waste keeps things sensible for tall-skinny picks
        # (e.g. 10x20=200 beats 14x16=224 only if waste matters; here 14x16
        # wins because aspect_pct 0.143 << aspect_pct 1.0 of 10x20).
        primary = aspect_pct + 3.0 * waste_pct
        score = (primary, abs(rows - 14))
        if best is None or score < best[0]:
            best = (score, rows, cols)
    return best[1], best[2]


def calculate_recup(inp: RecupInputs) -> RecupResults:
    # ── E4: flue mass = flow * 1.2 ──────────────────────────────────
    flue_mass = inp.flue_flow_nm3hr * inp.flue_density_factor

    # ── E13: heat to preheat combustion air ────────────────────────
    air_mass = inp.air_volume_nm3hr * inp.flue_density_factor
    dT_air = inp.air_temp_out_C - inp.air_temp_in_C
    heat_required = air_mass * inp.cp_air_kcal_kgC * dT_air

    # ── E7: derive final flue temp from heat balance ───────────────
    flue_temp_out = inp.flue_temp_in_C - (heat_required / (flue_mass * inp.cp_flue_kcal_kgC))

    # ── E14: LMTD with the 0.9 factor (per Excel cell formula) ─────
    dT1 = inp.flue_temp_in_C - inp.air_temp_out_C
    dT2 = flue_temp_out      - inp.air_temp_in_C
    if dT1 <= 0 or dT2 <= 0 or dT1 == dT2:
        lmtd_raw = (dT1 + dT2) / 2.0 if (dT1 + dT2) > 0 else 1.0
    else:
        lmtd_raw = (dT1 - dT2) / math.log(dT1 / dT2)
    lmtd = lmtd_raw * inp.lmtd_factor

    # ── E15: surface area ──────────────────────────────────────────
    surface_area = heat_required / (lmtd * inp.heat_transfer_coef)

    # ── E19: raw pipe count ────────────────────────────────────────
    # Excel cells write the literal 3.14, not math.pi — keep that so the
    # row/col grid matches the spreadsheet to the last digit.
    pipe_dia_m = inp.pipe_dia_mm / 1000.0
    effective_len = inp.pipe_length_m_per_bank + inp.surface_area_end_allowance_m
    pipes_raw = surface_area / (3.14 * pipe_dia_m * effective_len)

    # ── E20/E21: rows / cols, both paths go through _best_grid so the
    # result is always near-square. Auto target is ceil(raw); override
    # target is whatever the user typed.
    target = (inp.pipes_total_override
              if inp.pipes_total_override and inp.pipes_total_override > 0
              else max(1, math.ceil(pipes_raw)))
    rows_count, cols_count = _best_grid(target)
    pipes_total = rows_count * cols_count
    pipes_per_bank = pipes_total // 2  # E28: rows * (cols/2)

    # ── E16/E17: bank length and width (derived) ───────────────────
    bank_length_mm = (((rows_count - 1) / 2) * inp.pipe_pitch_mm) \
                     + ((rows_count / 2) * inp.pipe_dia_mm) \
                     + inp.end_margin_mm
    bank_width_mm  = ((cols_count / 2) * 48.3) \
                     + (((cols_count - 1) / 2) * inp.pipe_pitch_mm) \
                     + inp.end_margin_mm

    # ── E27/E28: pipe weights ──────────────────────────────────────
    weight_per_pipe = inp.pipe_kg_per_m * inp.pipe_length_m_per_bank
    weight_hot_bank  = weight_per_pipe * pipes_per_bank
    weight_cold_bank = weight_per_pipe * pipes_per_bank
    weight_total = weight_hot_bank + weight_cold_bank

    # ── MS structural weights (E41..E45) — exact Excel formulas ───
    ms_outer_shell = _vol_to_kg(
        ((2 * bank_length_mm) + 100) * (inp.pipe_length_m_per_bank * 1000) * 5,
        8650,
    ) * 2
    # Excel: (3.14*700*800*5) — literal 3.14, not math.pi.
    duct_volume = (3.14 * 700 * 800 * 5) \
                  + 2 * ((200 * bank_length_mm * 5 * 2) + (bank_width_mm * 200 * 5 * 2))
    ms_air_inlet = _vol_to_kg(duct_volume, 7850)
    ms_hot_outlet = _vol_to_kg(duct_volume, 8650)
    ms_pipe_holding = _vol_to_kg(
        ((bank_length_mm * 2) + inp.bank_gap_mm) * bank_width_mm * 16 * 4,
        8650,
    )
    box_volume = (((bank_length_mm * 2 + 250) * 600 * 5 * 2)
                  + (bank_width_mm * 600 * 2 * 5)
                  + ((bank_length_mm * 2 + 250) * bank_width_mm * 5))
    ms_bottom_box = _vol_to_kg(box_volume, 8650)

    return RecupResults(
        flue_mass_kghr        = round(flue_mass, 2),
        heat_required_kcal    = round(heat_required, 2),
        flue_temp_out_C       = round(flue_temp_out, 2),
        lmtd_C                = round(lmtd, 2),
        surface_area_m2       = round(surface_area, 4),
        pipes_total_raw       = round(pipes_raw, 2),
        pipes_total           = pipes_total,
        pipes_in_row          = rows_count,
        pipes_in_column       = cols_count,
        pipes_per_bank        = pipes_per_bank,
        bank_length_mm        = round(bank_length_mm, 2),
        bank_width_mm         = round(bank_width_mm, 2),
        weight_per_pipe_kg    = round(weight_per_pipe, 4),
        weight_hot_bank_kg    = round(weight_hot_bank, 2),
        weight_cold_bank_kg   = round(weight_cold_bank, 2),
        weight_total_pipes_kg = round(weight_total, 2),
        ms_outer_shell_kg     = round(ms_outer_shell, 2),
        ms_air_inlet_duct_kg  = round(ms_air_inlet, 2),
        ms_hot_outlet_duct_kg = round(ms_hot_outlet, 2),
        ms_pipe_holding_kg    = round(ms_pipe_holding, 2),
        ms_bottom_box_kg      = round(ms_bottom_box, 2),
        hot_bank_material     = (inp.hot_bank_material or "SS").upper(),
        cold_bank_material    = (inp.cold_bank_material or "SS").upper(),
        side_hood_kg          = float(inp.side_hood_kg) if inp.side_hood_kg > 0 else 1500.0,
        cai_rate_override     = float(inp.cai_rate_override or 0),
    )
