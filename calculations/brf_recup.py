"""The BRF recuperator, ported from the workbook's 'Recuperator' sheet.

A cross-flow tube recuperator: flue gas gives up heat on its way out, the
combustion air picks it up on its way in. The chain is short — heat required,
LMTD, surface area, then the tube bank that provides that surface and what it
weighs and costs.

Two of its inputs are typed on the sheet rather than linked to the furnace,
and they do not agree with the furnace they sit beside:

    E4  total flue gas          13,200 Nm3/hr   furnace gives 31,050
    E10 combustion air          20,400 Nm3/hr   furnace gives 28,350

Both come from the furnace now. The flue gas is the firing rate plus the
combustion air — what goes in comes out — and the air to preheat is the air
the burners are getting. The defaults on this dataclass are still the sheet's
figures, so calling it bare reproduces the workbook; main.py passes the
computed pair in, and link_recup_to_furnace=False reverts to the typed ones.

Cells are named as the workbook writes them.
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field

SHEET_PI = 3.14          # the workbook writes 3.14, not pi
SS_DENSITY = 8650.0      # kg/m3, the tube, the shell, the plates and the box
MS_DENSITY = 7850.0      # only the combustion air inlet duct is mild steel


@dataclass
class BRFRecupInputs:
    # ── Duty ────────────────────────────────────────────────────────
    flue_gas_nm3hr:      float = 13200.0   # E4
    flue_mass_factor:    float = 1.2       # E5, Nm3 -> kg
    flue_specific_heat:  float = 0.23      # E6
    flue_temp_in_C:      float = 650.0     # E7
    htc_kcal_m2C:        float = 30.0      # E9
    air_nm3hr:           float = 20400.0   # E10
    air_mass_factor:     float = 1.2       # the 1.2 in E14
    air_temp_in_C:       float = 30.0      # E11
    air_temp_out_C:      float = 300.0     # E12
    air_specific_heat:   float = 0.247     # E13
    lmtd_factor:         float = 0.9       # the sheet's radiative correction
    # ── Bank ────────────────────────────────────────────────────────
    pipes_per_row:       int   = 28        # E21
    pipes_per_column:    int   = 27        # E22
    bank_gap_mm:         float = 250.0     # E19
    pipe_pitch_mm:       float = 32.0      # the 32 in E17/E18
    bank_margin_mm:      float = 150.0     # the 150 in E17/E18
    pipe_dia_mm:         float = 48.3      # E24
    pipe_length_m:       float = 2.8       # E25
    pipe_end_allow_m:    float = 0.1       # the +0.1 in E20
    hot_pipe_thick_mm:   float = 3.6       # E26
    cold_pipe_kg_per_m:  float = 4.0       # E33, a plain tube, taken flat
    # ── Rates ───────────────────────────────────────────────────────
    ss304_rate:          float = 300.0     # F39
    ms_rate:             float = 75.0      # F41
    ss_outer_shell_cost: float = 150000.0  # F48
    bend_rate_per_pipe:  float = 350.0     # F51
    weld_rods_per_pipe:  int   = 4         # F52
    weld_rod_rate:       float = 8.0       # F52
    hole_rate:           float = 100.0     # F53, twice per pipe
    thermocouple_cost:   float = 8000.0    # F54
    markup:              float = 1.8       # F56
    # Shell and duct geometry, the constants inside E43..E47
    shell_thick_mm:      float = 5.0
    duct_dia_mm:         float = 700.0
    duct_len_mm:         float = 800.0
    holding_plate_thick_mm: float = 12.0
    bottom_box_height_mm:   float = 600.0


@dataclass
class BRFRecupResults:
    flue_mass_kghr:      float   # E5
    flue_temp_out_C:     float   # E8
    heat_required_kcal:  float   # E14
    lmtd_C:              float   # E15
    surface_area_m2:     float   # E16
    bank_length_mm:      float   # E17
    bank_width_mm:       float   # E18
    pipes_required:      float   # E20
    pipes_provided:      int     # E21 x E22
    hot_pipe_kg_per_m:   float   # E27
    hot_pipe_kg:         float   # E28
    cold_pipe_kg:        float   # E34
    total_pipe_kg:       float   # E35
    pipe_cost:           float   # F42
    outer_shell_kg:      float   # E43
    air_inlet_duct_kg:   float   # E44
    hot_air_duct_kg:     float   # E45
    holding_plate_kg:    float   # E46
    bottom_box_kg:       float   # E47
    ms_parts_cost:       float   # F49
    material_cost:       float   # F50
    bending_cost:        float   # F51
    welding_cost:        float   # F52
    hole_cost:           float   # F53
    total_cost:          float   # F55
    selling_price:       float   # F57
    # What the furnace itself implies, for comparison with the typed inputs.
    furnace_air_nm3hr:   float = 0.0
    furnace_flue_nm3hr:  float = 0.0
    inputs_linked:       bool  = False
    # The duty inputs echoed back, so the sheet can print what it was given
    # rather than only what came out.
    flue_gas_nm3hr:      float = 0.0   # E4
    flue_temp_in_C:      float = 0.0   # E7
    air_nm3hr:           float = 0.0   # E10
    air_temp_in_C:       float = 0.0   # E11
    air_temp_out_C:      float = 0.0   # E12
    rows:                int   = 0     # E21
    cols:                int   = 0     # E22


def calculate_recuperator(inp=None, furnace_air_nm3hr=0.0,
                          furnace_flue_nm3hr=0.0) -> BRFRecupResults:
    r = inp or BRFRecupInputs()

    # ── Duty ────────────────────────────────────────────────────────
    flue_mass = r.flue_gas_nm3hr * r.flue_mass_factor                     # E5
    heat = (r.air_nm3hr * r.air_mass_factor) * r.air_specific_heat \
        * (r.air_temp_out_C - r.air_temp_in_C)                            # E14
    denom = flue_mass * r.flue_specific_heat
    flue_out = r.flue_temp_in_C - (heat / denom if denom else 0.0)         # E8

    # LMTD, counter-flow, with the sheet's 0.9 radiative correction.       E15
    d1 = r.flue_temp_in_C - r.air_temp_out_C
    d2 = flue_out - r.air_temp_in_C
    lmtd = ((d1 - d2) / math.log(d1 / d2) * r.lmtd_factor
            if d1 > 0 and d2 > 0 and d1 != d2 else 0.0)

    area = heat / (lmtd * r.htc_kcal_m2C) if lmtd else 0.0                # E16

    # ── Bank ────────────────────────────────────────────────────────
    rows, cols = r.pipes_per_row, r.pipes_per_column
    bank_len = (((rows - 1) / 2) * r.pipe_pitch_mm
                + (rows / 2) * r.pipe_dia_mm + r.bank_margin_mm)           # E17
    bank_wid = ((cols / 2) * r.pipe_dia_mm
                + ((cols - 1) / 2) * r.pipe_pitch_mm + r.bank_margin_mm)   # E18
    pipes_req = (area / (SHEET_PI * (r.pipe_dia_mm / 1000.0)
                         * (r.pipe_length_m + r.pipe_end_allow_m))
                 if area else 0.0)                                         # E20

    # ── Pipes ───────────────────────────────────────────────────────
    hot_kg_m = (SHEET_PI * r.pipe_dia_mm * (r.pipe_length_m * 1000.0)
                * r.hot_pipe_thick_mm * SS_DENSITY) / 1e9 / r.pipe_length_m  # E27
    hot_kg = hot_kg_m * r.pipe_length_m                                   # E28
    cold_kg = r.cold_pipe_kg_per_m * r.pipe_length_m                      # E34
    half = rows * (cols / 2)
    total_pipe_kg = hot_kg * half + cold_kg * half                        # E35
    pipe_cost = total_pipe_kg * r.ss304_rate                              # F42

    # ── Shell, ducts, plates, box ───────────────────────────────────
    t, L = r.shell_thick_mm, r.pipe_length_m * 1000.0
    # The shell wraps both banks, hence the trailing doubling.
    shell = ((2 * bank_len + 100) * L * t) * (SS_DENSITY / 1e9) * 2        # E43
    # The two ducts are the same shape and differ only in what they are made
    # of: the cold air inlet is mild steel, the hot air outlet stainless.
    duct_mm3 = (SHEET_PI * r.duct_dia_mm * r.duct_len_mm * t
                + 2 * ((200 * bank_len * t * 2) + (bank_wid * 200 * t * 2)))
    air_in = duct_mm3 * (MS_DENSITY / 1e9)                                 # E44
    hot_out = duct_mm3 * (SS_DENSITY / 1e9)                                # E45
    holding = ((bank_len * 2 + r.bank_gap_mm) * bank_wid
               * r.holding_plate_thick_mm * 2) * (SS_DENSITY / 1e9)        # E46
    box = ((bank_len * 2 + r.bank_gap_mm) * r.bottom_box_height_mm * t * 2
           + bank_wid * r.bottom_box_height_mm * 2 * t
           + (bank_len * 2 + r.bank_gap_mm) * bank_wid * t) \
        * (SS_DENSITY / 1e9)                                               # E47

    ms_cost = (air_in + shell + hot_out + holding + box) * r.ms_rate       # F49
    material = pipe_cost + r.ss_outer_shell_cost + ms_cost                 # F50

    bending = (rows + cols) * r.bend_rate_per_pipe                         # F51
    welding = rows * cols * r.weld_rods_per_pipe * r.weld_rod_rate         # F52
    holes = rows * cols * 2 * r.hole_rate                                  # F53
    total = material + bending + welding + holes + r.thermocouple_cost     # F55
    selling = math.ceil(r.markup * total / 10.0) * 10.0                    # F57

    def _4(x):
        return round(x, 4)

    return BRFRecupResults(
        flue_mass_kghr=_4(flue_mass), flue_temp_out_C=_4(flue_out),
        heat_required_kcal=_4(heat), lmtd_C=_4(lmtd), surface_area_m2=_4(area),
        bank_length_mm=_4(bank_len), bank_width_mm=_4(bank_wid),
        pipes_required=_4(pipes_req), pipes_provided=int(rows * cols),
        hot_pipe_kg_per_m=_4(hot_kg_m), hot_pipe_kg=_4(hot_kg),
        cold_pipe_kg=_4(cold_kg), total_pipe_kg=_4(total_pipe_kg),
        pipe_cost=round(pipe_cost, 2),
        outer_shell_kg=_4(shell), air_inlet_duct_kg=_4(air_in),
        hot_air_duct_kg=_4(hot_out), holding_plate_kg=_4(holding),
        bottom_box_kg=_4(box), ms_parts_cost=round(ms_cost, 2),
        material_cost=round(material, 2), bending_cost=round(bending, 2),
        welding_cost=round(welding, 2), hole_cost=round(holes, 2),
        total_cost=round(total, 2), selling_price=round(selling, 2),
        furnace_air_nm3hr=round(furnace_air_nm3hr, 2),
        furnace_flue_nm3hr=round(furnace_flue_nm3hr, 2),
        inputs_linked=False,
        flue_gas_nm3hr=r.flue_gas_nm3hr, flue_temp_in_C=r.flue_temp_in_C,
        air_nm3hr=r.air_nm3hr, air_temp_in_C=r.air_temp_in_C,
        air_temp_out_C=r.air_temp_out_C, rows=rows, cols=cols,
    )
