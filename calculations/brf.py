"""Billet reheating furnace — sizing chain, ported from the
'Furnace costing _60TPH_BRF 12mtr.xlsx' workbook (07/08/2026).

This covers the two sheets that decide what the furnace fires and what it
is piped with. Each step names the cell it comes from:

  calculation sheet — furnace duty
    D3  Furnace capacity          -> user input (t/hr)
    D4  Fuel consumption per ton  -> user input (SCM/ton)
    D5  Firing rate               -> D3 * D4
    D6  CV of the fuel            -> user input (kcal/Nm3)
    D7  Combustion air per Nm3    -> user input
    D8  Combustion air            -> D5 * D7
    F1  Number of zones           -> user input
    D9  Firing rate per zone      -> ROUNDUP(D5 / F1, 0)
    D11 Blower                    -> ((D8 / 1.7) * 40 / 3200)  HP

  Sizing Zone sheet — one row per zone, plus any standalone burner
    B   Burner rating (kW)        -> user input
    C   Number of burners         -> user input
    D   Fuel flow per burner      -> B * 860 / CV      (or typed over)
    E   Fuel flow for the zone    -> D * C
    F   Air flow at 30 C          -> B * C
    G   Air flow at preheat temp  -> F * (1/(P+1)) * ((T+273)/273)
    H   Air line area             -> G / air velocity / 3600
    I   Air line bore             -> ((4 * H / 3.14) ^ 0.5) * 1000
    J   Air line size             -> next standard NB at or above I
    K   Gas line area             -> E / gas velocity / 3600
    L   Gas line bore             -> ((4 * K / 3.14) ^ 0.5) * 1000
    M   Gas line size             -> next standard NB at or above L

The workbook writes the literal 3.14 rather than pi, and the bore comes out
about 0.05% wide as a result; that is kept, so a line sized here and a line
sized on the sheet land on the same NB every time.

One deliberate departure. The workbook's gas-line column reads
  K9 = J9 / $O$3 / 3600
which divides the SELECTED AIR LINE SIZE — 500, a bore in millimetres — by
the gas velocity, where the fuel flow for the zone (E) belongs. It is a
mis-reference rather than a method: the air column alongside it uses its own
flow (G), and every other area on the sheet is flow over velocity. Zone 1
lands on 80 NB from it where its 600 Nm3/hr of gas needs 100 NB. This module
sizes the gas line from the gas flow. See gas_line_nb_as_sheet on each zone
for what the workbook's own formula produces, so the two can be compared.
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field

from calculations.pipes import round_up_to_nb


# The workbook's own constants, named. All are overridable per quote.
KCAL_PER_KWH = 860.0            # the 860 in B * 860 / CV
CFM_PER_NM3HR = 1.7             # the 1.7 in the blower formula
BLOWER_HEAD_FACTOR = 40.0       # the 40
BLOWER_CONSTANT = 3200.0        # the 3200
SHEET_PI = 3.14                 # the workbook writes 3.14, not pi


@dataclass
class BRFZone:
    """One row of the Sizing Zone table — a fired zone, or a standalone
    burner sized alongside them."""
    name:          str
    burner_kw:     float
    burner_count:  int
    # False for a standalone burner sized alongside the zones. It takes a line
    # of its own in the table but is not one of the fired zones the firing rate
    # is divided between, so counting it would under-state every zone's share.
    is_zone:       bool = True
    # Fuel per burner. 0 means derive it from the rating and the calorific
    # value; the workbook types it over on the zones where the burner is run
    # richer than its rating implies (zones 1 and 2 carry 150 against a
    # derived 100).
    fuel_per_burner_nm3hr: float = 0.0


@dataclass
class BRFInputs:
    # ── Furnace duty (calculation sheet) ────────────────────────────
    furnace_capacity_tph:   float = 60.0     # D3
    fuel_per_ton_scm:       float = 45.0     # D4
    cv_kcal_nm3:            float = 8600.0   # D6
    combustion_air_per_nm3: float = 10.5     # D7
    # ── Line sizing (Sizing Zone sheet) ─────────────────────────────
    preheat_pressure_barg:  float = 0.05     # J2
    preheat_air_temp_C:     float = 300.0    # J3
    air_velocity_ms:        float = 12.0     # O2
    gas_velocity_ms:        float = 30.0     # O3
    # ── Zones ───────────────────────────────────────────────────────
    zones: list[BRFZone] = field(default_factory=list)
    # ── Blower ──────────────────────────────────────────────────────
    cfm_per_nm3hr:      float = CFM_PER_NM3HR
    blower_head_factor: float = BLOWER_HEAD_FACTOR
    blower_constant:    float = BLOWER_CONSTANT


@dataclass
class BRFZoneResult:
    name:                 str
    burner_kw:            float
    burner_count:         int
    fuel_per_burner_nm3hr: float   # D
    fuel_flow_nm3hr:      float    # E
    air_flow_nm3hr:       float    # F — at 30 C
    air_flow_hot_m3hr:    float    # G — at the preheat temperature
    air_line_area_m2:     float    # H
    air_line_bore_mm:     float    # I
    air_line_nb:          int      # J
    gas_line_area_m2:     float    # K
    gas_line_bore_mm:     float    # L
    gas_line_nb:          int      # M
    # What the workbook's own gas formula gives, for comparison — see the
    # module docstring. Not used for anything downstream.
    gas_line_bore_as_sheet_mm: float
    gas_line_nb_as_sheet:      int


@dataclass
class BRFResults:
    # Furnace duty
    firing_rate_nm3hr:      float   # D5
    combustion_air_nm3hr:   float   # D8
    firing_rate_per_zone_nm3hr: int # D9
    zone_count:             int     # F1
    blower_cfm:             float
    blower_hp:              float   # D11
    # Zones
    zones:                  list[BRFZoneResult]
    # Roll-ups across the zone table
    total_burners:          int
    total_fuel_nm3hr:       float
    total_air_nm3hr:        float
    # Mains — the air header from the blower and the gas header from the train,
    # sized on the whole furnace rather than a zone's share
    air_main_flow_m3hr:     float
    air_main_area_m2:       float
    air_main_bore_mm:       float
    air_main_nb:            int
    gas_main_flow_nm3hr:    float
    gas_main_area_m2:       float
    gas_main_bore_mm:       float
    gas_main_nb:            int


def _bore_mm(area_m2: float) -> float:
    """Inner bore from a flow area, the workbook's way: ((4A/3.14)^0.5)*1000."""
    if area_m2 <= 0:
        return 0.0
    return math.sqrt(4 * area_m2 / SHEET_PI) * 1000


def _nb_or_zero(bore_mm: float) -> int:
    """Next standard NB at or above the bore, or 0 when there is no pipe that
    big.

    The ladder in calculations/pipes.py stops at 600 NB, and it raises rather
    than returning something. That is right for a burner line, but a reheating
    furnace's air header is a duct, not a pipe: 28,350 Nm3/hr of preheated air
    at 12 m/s wants a 1,293 mm bore, and the whole calculation used to die on
    it. 0 means "past the ladder" — the bore is still reported, and whatever
    displays it can say so instead of showing a pipe size that does not exist.
    """
    try:
        return round_up_to_nb(bore_mm) if bore_mm > 0 else 0
    except ValueError:
        return 0


def calculate_brf(inp: BRFInputs) -> BRFResults:
    # A cleared box arrives as 0 and used to surface as "division by zero",
    # which says nothing about which box. Name it instead.
    for value, what in ((inp.cv_kcal_nm3, "calorific value"),
                        (inp.air_velocity_ms, "air velocity"),
                        (inp.gas_velocity_ms, "gas velocity")):
        if not value or value <= 0:
            raise ValueError(f"The {what} cannot be zero — nothing can be sized from it.")

    # ── D5: firing rate ────────────────────────────────────────────
    firing_rate = inp.furnace_capacity_tph * inp.fuel_per_ton_scm

    # ── D8: combustion air for that firing rate ────────────────────
    combustion_air = firing_rate * inp.combustion_air_per_nm3

    # ── D9: firing rate per zone, rounded up to a whole Nm3/hr ─────
    zone_count = sum(1 for z in inp.zones if z.is_zone and z.burner_count > 0)
    per_zone = int(math.ceil(firing_rate / zone_count)) if zone_count else 0

    # ── D11: blower, in horsepower ─────────────────────────────────
    # The air flow is put back into CFM before the head factor is applied —
    # the same 1.7 that converts a blower's CFM to Nm3/hr elsewhere.
    blower_cfm = combustion_air / inp.cfm_per_nm3hr if inp.cfm_per_nm3hr else 0.0
    blower_hp = blower_cfm * inp.blower_head_factor / inp.blower_constant

    # ── The zone table ─────────────────────────────────────────────
    # Air is expanded from 30 C to the preheat temperature and corrected for
    # the gauge pressure it is delivered at, because the line has to carry the
    # hot volume, not the normal volume.
    expansion = (1 / (inp.preheat_pressure_barg + 1)) \
                * ((inp.preheat_air_temp_C + 273) / 273)

    zones: list[BRFZoneResult] = []
    for z in inp.zones:
        per_burner = (z.fuel_per_burner_nm3hr
                      or (z.burner_kw * KCAL_PER_KWH / inp.cv_kcal_nm3))
        fuel_flow = per_burner * z.burner_count
        air_flow = z.burner_kw * z.burner_count
        air_hot = air_flow * expansion

        air_area = air_hot / inp.air_velocity_ms / 3600 if inp.air_velocity_ms else 0.0
        air_bore = _bore_mm(air_area)
        air_nb = _nb_or_zero(air_bore)

        gas_area = fuel_flow / inp.gas_velocity_ms / 3600 if inp.gas_velocity_ms else 0.0
        gas_bore = _bore_mm(gas_area)
        gas_nb = _nb_or_zero(gas_bore)

        # The workbook's own formula, kept only so the two can be compared.
        sheet_area = air_nb / inp.gas_velocity_ms / 3600 if inp.gas_velocity_ms else 0.0
        sheet_bore = _bore_mm(sheet_area)

        zones.append(BRFZoneResult(
            name                  = z.name,
            burner_kw             = z.burner_kw,
            burner_count          = z.burner_count,
            fuel_per_burner_nm3hr = round(per_burner, 4),
            fuel_flow_nm3hr       = round(fuel_flow, 2),
            air_flow_nm3hr        = round(air_flow, 2),
            air_flow_hot_m3hr     = round(air_hot, 2),
            air_line_area_m2      = round(air_area, 6),
            air_line_bore_mm      = round(air_bore, 2),
            air_line_nb           = air_nb,
            gas_line_area_m2      = round(gas_area, 6),
            gas_line_bore_mm      = round(gas_bore, 2),
            gas_line_nb           = gas_nb,
            gas_line_bore_as_sheet_mm = round(sheet_bore, 2),
            gas_line_nb_as_sheet      = _nb_or_zero(sheet_bore),
        ))

    # ── The mains ──────────────────────────────────────────────────
    # The 4 TPH sheet sizes these in a block of its own (B13:B19 for gas,
    # E13:E19 with J14:K19 for air); the 60 TPH sheet sizes only the branches.
    # Same method as a zone, on the whole furnace: the air header carries all
    # the combustion air at the preheat temperature, the gas header the whole
    # firing rate.
    #
    # That sheet expands the air by (T_preheat + 273.15) / (T_ambient + 273)
    # where the zone columns use 1/(P+1) x (T+273)/273. Ours uses the zone
    # expansion throughout, so a header and the branches off it are sized on
    # the same physics — on the 4 TPH job both land on 300 NB either way.
    # Its K18 also adds T_ambient/293 to the flow, about 1 Nm3/hr on 1890,
    # which looks like a slip and changes nothing; it is not carried here.
    air_main_flow = combustion_air * expansion
    air_main_area = air_main_flow / inp.air_velocity_ms / 3600 if inp.air_velocity_ms else 0.0
    air_main_bore = _bore_mm(air_main_area)
    gas_main_area = firing_rate / inp.gas_velocity_ms / 3600 if inp.gas_velocity_ms else 0.0
    gas_main_bore = _bore_mm(gas_main_area)

    return BRFResults(
        firing_rate_nm3hr          = round(firing_rate, 2),
        combustion_air_nm3hr       = round(combustion_air, 2),
        firing_rate_per_zone_nm3hr = per_zone,
        zone_count                 = zone_count,
        blower_cfm                 = round(blower_cfm, 2),
        blower_hp                  = round(blower_hp, 2),
        zones                      = zones,
        total_burners              = sum(z.burner_count for z in inp.zones),
        total_fuel_nm3hr           = round(sum(z.fuel_flow_nm3hr for z in zones), 2),
        total_air_nm3hr            = round(sum(z.air_flow_nm3hr for z in zones), 2),
        air_main_flow_m3hr         = round(air_main_flow, 2),
        air_main_area_m2           = round(air_main_area, 6),
        air_main_bore_mm           = round(air_main_bore, 2),
        air_main_nb                = _nb_or_zero(air_main_bore),
        gas_main_flow_nm3hr        = round(firing_rate, 2),
        gas_main_area_m2           = round(gas_main_area, 6),
        gas_main_bore_mm           = round(gas_main_bore, 2),
        gas_main_nb                = _nb_or_zero(gas_main_bore),
    )


def brf_60tph_12m() -> BRFInputs:
    """The uploaded 60 TPH / 12 m billet job, as the workbook has it —
    five fired zones plus the standalone burner sized beside them."""
    return BRFInputs(
        furnace_capacity_tph=60.0, fuel_per_ton_scm=45.0,
        cv_kcal_nm3=8600.0, combustion_air_per_nm3=10.5,
        preheat_air_temp_C=300.0, air_velocity_ms=12.0, gas_velocity_ms=30.0,
        zones=[
            BRFZone("Zone 1", 1000, 4, fuel_per_burner_nm3hr=150),
            BRFZone("Zone 2", 1000, 5, fuel_per_burner_nm3hr=150),
            BRFZone("Zone 3", 1000, 5),
            BRFZone("Zone 4", 1000, 5),
            BRFZone("Zone 5", 1000, 5),
            BRFZone("Burner", 1000, 1, is_zone=False),
        ],
    )


# ── Furnace geometry ────────────────────────────────────────────────────────
# Section 1 of the Furnace sheet: the billet, the hearth it needs, and the
# overall box that comes out once the refractory, sheet and channel are added
# round it. Cells named as the workbook writes them.
#
#   C7  Volume of billet     -> L * W * H, in metres
#   C9  Weight of a billet   -> volume * density
#   C13 Hearth area          -> capacity * 1000 / top-fired hearth load
#   C14 Effective length     -> ROUND(area / billet length, 0)
#   D17 Effective width      -> billet length + 0.3
#   D24 Overall width        -> effective width + right + left + sheet + channel
#   G25 Overall length       -> effective length + discharge + charge
#                               + both sheets + channel
#   C28/F28/H28 Zone lengths -> shares of the overall length

@dataclass
class BRFFurnaceInputs:
    # ── Billet ──────────────────────────────────────────────────────
    billet_length_mm:   float = 12000.0   # C4
    billet_width_mm:    float = 130.0     # C5
    billet_height_mm:   float = 130.0     # C6
    ms_density_kg_m3:   float = 7850.0    # C8
    # ── Duty ────────────────────────────────────────────────────────
    furnace_capacity_tph:      float = 60.0    # C10
    hearth_load_top_kg_hr_m2:  float = 300.0   # C11
    # Carried because the sheet asks for it, though the hearth area is worked
    # out from the top-fired figure alone.
    hearth_load_top_bottom_kg_hr_m2: float = 600.0   # C12
    # ── Around the width (mm) ───────────────────────────────────────
    right_refractory_mm: float = 510.0    # D20
    left_refractory_mm:  float = 510.0    # D21
    width_sheet_mm:      float = 16.0     # D22
    width_channel_mm:    float = 500.0    # D23
    # ── Around the length (mm) ──────────────────────────────────────
    discharge_refractory_mm: float = 2010.0  # G20
    charge_refractory_mm:    float = 1500.0  # G21
    sheet_charging_side_mm:  float = 6.0     # G22
    sheet_refractory_side_mm: float = 6.0    # G23
    length_channel_mm:       float = 300.0   # G24
    # ── Overrides (0 = derive) ──────────────────────────────────────
    # The sheet types the two effective dimensions into the overall-size block
    # rather than referring to the cells above. They agree here, but a furnace
    # whose box was set on the drawing can say so.
    effective_width_mm_override:  float = 0.0   # D19
    effective_length_mm_override: float = 0.0   # G19


@dataclass
class BRFFurnaceResults:
    # Echoed so the working can quote the figures it was computed from rather
    # than whatever the input boxes hold when it is painted.
    billet_length_mm:    float
    billet_width_mm:     float
    billet_height_mm:    float
    ms_density_kg_m3:    float
    billet_volume_m3:    float
    billet_weight_kg:    float
    hearth_area_m2:      float
    effective_length_m:  float
    effective_width_m:   float
    effective_length_mm: float
    effective_width_mm:  float
    overall_width_mm:    float
    overall_width_m:     float
    overall_length_mm:   float
    overall_length_m:    float
    zone_preheating_m:   float
    zone_heating_m:      float
    zone_soaking_m:      float
    zone_heating_soaking_m: float


def calculate_furnace(inp: BRFFurnaceInputs) -> BRFFurnaceResults:
    L_m = inp.billet_length_mm / 1000.0
    W_m = inp.billet_width_mm  / 1000.0
    H_m = inp.billet_height_mm / 1000.0

    volume = L_m * W_m * H_m                                   # C7
    weight = volume * inp.ms_density_kg_m3                     # C9
    area   = (inp.furnace_capacity_tph * 1000.0
              / inp.hearth_load_top_kg_hr_m2)                  # C13
    # ROUND, not ceiling: the sheet rounds 16.67 to 17 and would round 16.4
    # down to 16. Python's round() is banker's rounding, so do it explicitly.
    eff_len_m = math.floor(area / L_m + 0.5) if L_m else 0.0   # C14
    eff_wid_m = L_m + 0.3                                      # D17

    eff_len_mm = inp.effective_length_mm_override or eff_len_m * 1000.0   # G19
    eff_wid_mm = inp.effective_width_mm_override  or eff_wid_m * 1000.0   # D19

    overall_w_mm = (eff_wid_mm + inp.right_refractory_mm + inp.left_refractory_mm
                    + inp.width_sheet_mm + inp.width_channel_mm)          # D24
    overall_l_mm = (eff_len_mm + inp.discharge_refractory_mm
                    + inp.charge_refractory_mm + inp.sheet_charging_side_mm
                    + inp.sheet_refractory_side_mm + inp.length_channel_mm)  # G24/25
    overall_l_m = overall_l_mm / 1000.0

    preheating = overall_l_m * 0.2                                        # C28
    heating    = overall_l_m * 0.5 - 0.23 - 0.05                          # F28
    soaking    = (eff_len_mm + 1500 + 1500 + 230) / 1000.0 * 0.3          # H28

    return BRFFurnaceResults(
        billet_length_mm    = inp.billet_length_mm,
        billet_width_mm     = inp.billet_width_mm,
        billet_height_mm    = inp.billet_height_mm,
        ms_density_kg_m3    = inp.ms_density_kg_m3,
        billet_volume_m3    = round(volume, 6),
        billet_weight_kg    = round(weight, 2),
        hearth_area_m2      = round(area, 2),
        effective_length_m  = round(eff_len_m, 2),
        effective_width_m   = round(eff_wid_m, 3),
        effective_length_mm = round(eff_len_mm, 1),
        effective_width_mm  = round(eff_wid_mm, 1),
        overall_width_mm    = round(overall_w_mm, 1),
        overall_width_m     = round(overall_w_mm / 1000.0, 3),
        overall_length_mm   = round(overall_l_mm, 1),
        overall_length_m    = round(overall_l_m, 3),
        # Six places, not three. These feed the refractory take-off, where a
        # zone length is multiplied by a brick row every 100 mm — about 180x —
        # so rounding here to what the screen shows moved brick counts by a
        # whole brick and the roof weight by 2.6 kg. Display rounds; the engine
        # keeps the number.
        zone_preheating_m   = round(preheating, 6),
        zone_heating_m      = round(heating, 6),
        zone_soaking_m      = round(soaking, 6),
        zone_heating_soaking_m = round(heating + soaking, 6),
    )


# ── Refractory ──────────────────────────────────────────────────────────────
# Furnace sheet, section 2. The brick counts and weights that follow from the
# box worked out above: the roof it hangs from, the side walls either side, and
# the discharge-end wall. Cells are named as the workbook writes them.
#
# All of it is counted off the zone lengths and the effective width, so it
# moves with the billet and the capacity like the rest of the furnace.

@dataclass
class BRFRefractoryInputs:
    # Standard brick, (230 x 115 x 75) mm                            H35
    brick_l_mm: float = 230.0
    brick_w_mm: float = 115.0
    brick_h_mm: float = 75.0
    # Hanging brick spacing: one per 750 mm of (width + 460)         C35
    hanging_pitch_mm:  float = 750.0
    hanging_margin_mm: float = 460.0
    # Brick weights, high alumina                            C37/C39/C43/C46
    hanging_brick_60_kg: float = 39.36
    hanging_brick_40_kg: float = 36.85
    holding_brick_60_kg: float = 23.60
    holding_brick_40_kg: float = 22.10
    # Ceramic fibre, (7300 x 600 x 25) mm                    C48/C49/C51
    fibre_roll_length_mm: float = 7300.0
    fibre_roll_kg:        float = 15.0
    fibre_spare_rolls:    int   = 10
    # Castable                                               C53..C58
    castable_depth_mm:   float = 650.0
    castable_thick_mm:   float = 100.0
    castable_density:    float = 550.0
    castable_wastage:    float = 1.1
    castable_bag_kg:     float = 25.0
    # Wall heights, as the sheet builds them up              G37/G42/G47/G53
    side_wall_height_mm:          float = 1500 + 230 + 230 + 150   # 2110
    side_wall_internal_height_mm: float = 1500 + 230               # 1730
    preheat_wall_height_mm:       float = 400 + 230 + 230 + 150    # 1010
    preheat_wall_40_height_mm:    float = 400 + 115                # 515
    wall_thickness_mm:            float = 115.0
    # HAY, (600 x 900 x 50) mm                               L45/L46
    hay_l_mm: float = 600.0
    hay_w_mm: float = 900.0
    hay_t_mm: float = 50.0


@dataclass
class BRFRefractoryResults:
    brick_volume_mm3:           float   # H35
    # Roof
    hanging_bricks_per_width:   int     # D35
    hanging_bricks_soak_heat:   float   # C36
    hanging_bricks_preheat:     float   # C38
    hanging_brick_weight_kg:    float   # C40
    holding_bricks_per_width:   int     # C41
    holding_bricks_soak_heat:   float   # C42
    holding_brick_60_weight_kg: float   # C44
    holding_bricks_preheat:     float   # C45
    holding_brick_40_weight_kg: float   # C47
    fibre_rolls_raw:            float   # C50
    fibre_rolls:                int     # C51
    fibre_weight_kg:            float   # C52
    castable_volume_m3:         float   # C53
    castable_weight_kg:         float   # C54
    castable_total_kg:          float   # C56
    castable_bags:              float   # C58
    # Side walls, right and left
    side_wall_length_m:         float   # G36
    side_wall_cold_face_bricks: float   # G39
    side_wall_hot_face_bricks:  float   # G40
    side_wall_fire_bricks_60:   float   # G44
    preheat_wall_length_m:      float   # G46
    preheat_cold_face_bricks:   float   # G49
    preheat_hot_face_bricks:    float   # G51
    preheat_fire_bricks_40:     float   # G55
    tapered_wall_length_m:      float   # G58
    tapered_fire_bricks_40:     float   # G60
    # Discharge-end wall
    discharge_cold_face_bricks: float   # L37
    discharge_hot_face_bricks:  float   # L39
    discharge_fire_bricks_60:   float   # L44
    discharge_hay_pieces:       float   # L47


def calculate_refractory(fur, fin, inp=None) -> BRFRefractoryResults:
    """Brick counts and weights for the roof, the side walls and the discharge
    wall, off the furnace geometry.

    fur — BRFFurnaceResults, fin — BRFFurnaceInputs, inp — BRFRefractoryInputs.
    """
    r = inp or BRFRefractoryInputs()
    brick_vol = r.brick_l_mm * r.brick_w_mm * r.brick_h_mm             # H35
    t = r.wall_thickness_mm

    # ── Roof ───────────────────────────────────────────────────────
    hang_per_width = math.ceil(
        (fur.effective_width_mm + r.hanging_margin_mm) / r.hanging_pitch_mm)   # D35
    soak_heat = fur.zone_heating_m + fur.zone_soaking_m                        # G36
    hang_soak = (soak_heat + 0.23) * hang_per_width / 100 * 1000 + 50          # C36
    hang_pre = fur.zone_preheating_m * hang_per_width / 100 * 1000 + 25        # C38
    hang_wt = hang_soak * r.hanging_brick_60_kg + hang_pre * r.hanging_brick_40_kg  # C40

    hold_per_width = hang_per_width - 1                                        # C41
    hold_soak = hold_per_width * soak_heat * 1000 / 100 + 50                   # C42
    hold_wt60 = hold_soak * r.holding_brick_60_kg                              # C44
    hold_pre = fur.zone_preheating_m * 1000 / 100 * hold_per_width + 30        # C45
    hold_wt40 = hold_pre * r.holding_brick_40_kg                               # C47

    # The fibre runs the length of the furnace less the sheet and channel at
    # the ends, one run per holding-brick row plus the two edges.
    fibre_run_mm = (fur.overall_length_mm - (fin.sheet_charging_side_mm
                                             + fin.sheet_refractory_side_mm
                                             + fin.length_channel_mm))
    fibre_raw = fibre_run_mm / r.fibre_roll_length_mm * (hold_per_width + 2)   # C50
    fibre_rolls = math.ceil(fibre_raw) + r.fibre_spare_rolls                   # C51
    fibre_wt = fibre_rolls * r.fibre_roll_kg                                   # C52

    cast_vol = ((fur.effective_length_mm + fin.discharge_refractory_mm
                 + fin.charge_refractory_mm)
                * r.castable_depth_mm * r.castable_thick_mm / 1e9)             # C53
    cast_wt = cast_vol * r.castable_density                                    # C54
    cast_total = cast_wt * (hold_per_width + 2) * r.castable_wastage           # C56
    cast_bags = cast_total / r.castable_bag_kg                                 # C58

    # ── Side walls, right and left ─────────────────────────────────
    sw_vol = soak_heat * 1000 * r.side_wall_height_mm * t                      # G38
    sw_cold = sw_vol / brick_vol + 50                                          # G39
    sw_vol2 = soak_heat * 1000 * r.side_wall_internal_height_mm * t            # G43
    sw_fire = ((sw_vol2 / brick_vol) + 50) * 2 + 20                            # G44

    pre_len = fur.zone_preheating_m - 1.5                                      # G46
    pre_vol = pre_len * 1000 * r.preheat_wall_height_mm * t                    # G48
    pre_cold = pre_vol / brick_vol + 20                                        # G49
    pre_vol40 = pre_len * 1000 * r.preheat_wall_40_height_mm * t               # G54
    pre_fire40 = pre_vol40 / brick_vol * 2 + 20                                # G55

    tap_len = fur.zone_preheating_m - pre_len                                  # G58
    tap_vol = r.side_wall_height_mm * tap_len * 1000 * t                       # G59
    tap_fire = (tap_vol / brick_vol) * 2 + 50                                  # G60

    # ── Discharge-end wall ─────────────────────────────────────────
    dis_vol = fur.effective_width_mm * r.side_wall_height_mm * t               # L36
    dis_cold = dis_vol / brick_vol + 50                                        # L37
    dis_vol2 = fur.effective_width_mm * r.side_wall_internal_height_mm * t     # L43
    dis_fire = dis_vol2 / brick_vol * 2 + 50                                   # L44
    hay_vol = r.hay_l_mm * r.hay_w_mm * r.hay_t_mm                             # L46
    dis_hay = fur.effective_width_mm * r.side_wall_height_mm * r.hay_t_mm / hay_vol  # L47

    def _2(x):
        return round(x, 2)

    return BRFRefractoryResults(
        brick_volume_mm3=brick_vol,
        hanging_bricks_per_width=hang_per_width,
        hanging_bricks_soak_heat=_2(hang_soak),
        hanging_bricks_preheat=_2(hang_pre),
        hanging_brick_weight_kg=_2(hang_wt),
        holding_bricks_per_width=hold_per_width,
        holding_bricks_soak_heat=_2(hold_soak),
        holding_brick_60_weight_kg=_2(hold_wt60),
        holding_bricks_preheat=_2(hold_pre),
        holding_brick_40_weight_kg=_2(hold_wt40),
        fibre_rolls_raw=_2(fibre_raw),
        fibre_rolls=fibre_rolls,
        fibre_weight_kg=_2(fibre_wt),
        castable_volume_m3=round(cast_vol, 4),
        castable_weight_kg=_2(cast_wt),
        castable_total_kg=_2(cast_total),
        castable_bags=_2(cast_bags),
        side_wall_length_m=_2(soak_heat),
        side_wall_cold_face_bricks=_2(sw_cold),
        side_wall_hot_face_bricks=_2(sw_cold),
        side_wall_fire_bricks_60=_2(sw_fire),
        preheat_wall_length_m=_2(pre_len),
        preheat_cold_face_bricks=_2(pre_cold),
        preheat_hot_face_bricks=_2(pre_cold),
        preheat_fire_bricks_40=_2(pre_fire40),
        tapered_wall_length_m=_2(tap_len),
        tapered_fire_bricks_40=_2(tap_fire),
        discharge_cold_face_bricks=_2(dis_cold),
        discharge_hot_face_bricks=_2(dis_cold),
        discharge_fire_bricks_60=_2(dis_fire),
        discharge_hay_pieces=_2(dis_hay),
    )
