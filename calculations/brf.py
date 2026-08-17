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

from calculations.pipes import round_up_to_nb, select_oil_pipe_nb


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


# Oil-side constants, the same ones bom/regen_builder.py sizes its fans on, so
# an oil-fired furnace and an oil-fired regen system agree about the air.
OIL_AFR_KG_PER_KG = 15.0      # kg combustion air per kg of oil
RHO_AIR_KG_NM3    = 1.293     # kg/Nm3

GAS_FUELS = ("natural gas", "coke oven gas", "producer gas", "blast furnace gas")
OIL_FUELS = ("oil", "furnace oil", "fo", "hsd", "ldo", "hdo", "sko", "cfo", "lshs")


@dataclass
class BRFInputs:
    # ── Furnace duty (calculation sheet) ────────────────────────────
    furnace_capacity_tph:   float = 60.0     # D3
    fuel_per_ton_scm:       float = 45.0     # D4
    cv_kcal_nm3:            float = 8600.0   # D6
    combustion_air_per_nm3: float = 10.5     # D7
    # ── Fuel ────────────────────────────────────────────────────────
    # "Natural Gas", "Oil", or "Dual Fuel". A dual-fuel furnace fires either,
    # so both fuels are sized and each gets its own line; the air main is sized
    # on whichever of the two asks for more.
    fuel:                 str   = "Natural Gas"
    oil_per_ton_litre:    float = 40.0       # the oil equivalent of D4
    oil_cv_kcal_kg:       float = 10000.0    # furnace oil
    oil_density_kg_l:     float = 0.92
    oil_afr:              float = OIL_AFR_KG_PER_KG
    rho_air_kg_nm3:       float = RHO_AIR_KG_NM3
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
    # Oil side, on an oil or dual-fuel furnace. An oil line is sized by flow
    # band rather than velocity — it is a small bore carrying a liquid, not a
    # duct carrying a gas — so it comes from select_oil_pipe_nb.
    oil_flow_lph:         float = 0.0
    oil_line_nb:          int = 0


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
    # Fuel
    fuel:                   str = "Natural Gas"
    uses_gas:               bool = True
    uses_oil:               bool = False
    oil_firing_lph:         float = 0.0
    oil_firing_kghr:        float = 0.0
    oil_air_nm3hr:          float = 0.0
    gas_air_nm3hr:          float = 0.0
    oil_main_nb:            int = 0


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

    fuel_l = (inp.fuel or "Natural Gas").strip().lower()
    uses_oil = fuel_l in OIL_FUELS or fuel_l == "dual fuel"
    uses_gas = fuel_l not in OIL_FUELS or fuel_l == "dual fuel"

    # ── D5: firing rate ────────────────────────────────────────────
    firing_rate = (inp.furnace_capacity_tph * inp.fuel_per_ton_scm) if uses_gas else 0.0

    # ── The oil side, when there is one ────────────────────────────
    # Oil is metered by volume and burnt by mass, so the litres go to kilograms
    # before the air is worked out: air is a mass ratio, not a volume one.
    oil_lph = (inp.furnace_capacity_tph * inp.oil_per_ton_litre) if uses_oil else 0.0
    oil_kghr = oil_lph * inp.oil_density_kg_l
    oil_air = (oil_kghr * inp.oil_afr / inp.rho_air_kg_nm3) if inp.rho_air_kg_nm3 else 0.0

    # ── D8: combustion air ─────────────────────────────────────────
    # A dual-fuel furnace fires one or the other, never both, so the air side
    # is sized on whichever asks for more rather than on their sum.
    gas_air = firing_rate * inp.combustion_air_per_nm3
    combustion_air = max(gas_air, oil_air)

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
        # No gas line on an oil-fired furnace: the burners are the same burners
        # and take the same air, but nothing gas flows through them.
        per_burner = 0.0
        if uses_gas:
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

        # The same burners on oil: rating -> kcal/hr -> kg/hr -> litres/hr.
        oil_lph_zone = 0.0
        if uses_oil and inp.oil_cv_kcal_kg and inp.oil_density_kg_l:
            oil_lph_zone = (z.burner_kw * KCAL_PER_KWH / inp.oil_cv_kcal_kg
                            / inp.oil_density_kg_l) * z.burner_count

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
            oil_flow_lph              = round(oil_lph_zone, 2),
            oil_line_nb               = select_oil_pipe_nb(oil_lph_zone) if oil_lph_zone > 0 else 0,
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
        fuel                       = inp.fuel,
        uses_gas                   = uses_gas,
        uses_oil                   = uses_oil,
        oil_firing_lph             = round(oil_lph, 2),
        oil_firing_kghr            = round(oil_kghr, 2),
        oil_air_nm3hr              = round(oil_air, 2),
        gas_air_nm3hr              = round(gas_air, 2),
        oil_main_nb                = select_oil_pipe_nb(oil_lph) if oil_lph > 0 else 0,
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
    # Hearth: one brick on edge under the whole floor           L52..L71
    hearth_thickness_mm:  float = 115.0
    side_wall_cf_hf_mm:   float = 460.0   # K59, taken off the width for IS-8
    side_wall_hay_mm:     float = 100.0   # K60, likewise
    # Refractory block on the ground near the discharge door     L74..L79
    block_lane_pitch_mm:  float = 200.0
    block_soak_lanes:     float = 4.0
    block_length_m:       float = 0.6
    block_spare:          float = 2.0
    block_kg:             float = 55.0
    # Aluminium foil, (1000 x 600 x 0.5) mm                      C59..C64
    foil_l_m:      float = 1.0
    foil_w_m:      float = 0.6
    foil_t_m:      float = 0.0005
    foil_density:  float = 2700.0    # kg/m³
    foil_spare_m:  float = 10.0
    # Flue ducts, castable annulus around the pipe               C65..C80
    duct1_id_m: float = 1.9
    duct1_od_m: float = 2.2
    duct1_len_m: float = 14.0
    duct2_id_m: float = 1.9
    duct2_od_m: float = 2.1
    duct2_len_m: float = 7.5
    duct_castable_density: float = 2300.0   # LC40
    duct_castable_wastage: float = 1.1
    # Weight of one piece of each item, kg                       H68..H78
    kg_cold_face:  float = 1.8
    kg_hot_face:   float = 2.0
    kg_brick_60:   float = 5.0
    kg_brick_50:   float = 4.8
    kg_brick_40:   float = 4.3
    kg_is8:        float = 3.8
    kg_hysil:      float = 8.0
    # Mortar: one 50 kg bag per 200 bricks                       G75/G76
    bricks_per_mortar_bag: float = 200.0
    mortar_bag_kg:         float = 50.0
    firecreat_bags:        float = 10.0   # G78, for the gaps


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
    # The wall volumes and heights the counts are read off. Carried so the
    # sheet can be laid out the way the workbook lays it out, showing the
    # working rather than only the answer.
    side_wall_height_mm:        float = 0.0   # G37
    side_wall_volume_mm3:       float = 0.0   # G38
    side_wall_internal_height_mm: float = 0.0 # G42
    side_wall_volume_60_mm3:    float = 0.0   # G43
    preheat_wall_height_mm:     float = 0.0   # G47
    preheat_wall_volume_mm3:    float = 0.0   # G48
    preheat_wall_40_height_mm:  float = 0.0   # G53
    preheat_wall_40_volume_mm3: float = 0.0   # G54
    tapered_wall_volume_mm3:    float = 0.0   # G59
    discharge_volume_mm3:       float = 0.0   # L36
    discharge_volume_60_mm3:    float = 0.0   # L43
    hay_volume_mm3:             float = 0.0   # L46
    hanging_bricks_per_width_raw: float = 0.0 # C35
    fibre_roll_kg:              float = 0.0   # C49
    castable_bag_kg:            float = 0.0   # C57
    hanging_brick_60_kg:        float = 0.0
    hanging_brick_40_kg:        float = 0.0
    holding_brick_60_kg:        float = 0.0
    holding_brick_40_kg:        float = 0.0
    discharge_effective_width_mm: float = 0.0
    # ── Hearth ─────────────────────────────────────────  L53..L71
    hearth_overall_width_mm:    float = 0.0   # K58
    hearth_cf_volume_mm3:       float = 0.0   # L53
    hearth_cold_face_bricks:    float = 0.0   # L54
    hearth_hot_face_bricks:     float = 0.0   # L55
    hearth_is8_volume_mm3:      float = 0.0   # L61
    hearth_is8_bricks:          float = 0.0   # L62
    hearth_50_volume_mm3:       float = 0.0   # L64
    hearth_fire_bricks_50:      float = 0.0   # L65
    hearth_60_volume_mm3:       float = 0.0   # L67
    hearth_fire_bricks_60:      float = 0.0   # L68
    hearth_40_volume_mm3:       float = 0.0   # L70
    hearth_fire_bricks_40:      float = 0.0   # L71
    # ── Ground block near the discharge door ───────────  L75..L79
    block_lane_discharge:       float = 0.0   # L75
    block_lane_soaking:         float = 0.0   # L76
    block_pieces:               float = 0.0   # L78
    block_kg:                   float = 0.0   # L79
    # ── Aluminium foil ────────────────────────────────  C61..C64
    foil_piece_volume_m3:       float = 0.0   # C61
    foil_piece_kg:              float = 0.0   # C62
    foil_length_m:              float = 0.0   # C63
    foil_weight_kg:             float = 0.0   # C64
    # ── Flue duct castable ────────────────────────────  C69..C80
    duct1_volume_m3:            float = 0.0   # C69
    duct1_castable_kg:          float = 0.0   # C71
    duct1_bags:                 float = 0.0   # C72
    duct2_volume_m3:            float = 0.0   # C77
    duct2_castable_kg:          float = 0.0   # C79
    duct2_bags:                 float = 0.0   # C80
    # ── The take-off: every item, its count and its weight ──  F67..H79
    # [item, qty (nos), weight (kg)] in the workbook's own order.
    take_off:                   list  = field(default_factory=list)
    side_wall_hay_pieces:       float = 0.0   # G62
    total_refractory_kg:        float = 0.0
    total_refractory_tonne:     float = 0.0   # H79


def calculate_refractory(fur, fin, inp=None) -> BRFRefractoryResults:
    """Brick counts and weights for the roof, the side walls and the discharge
    wall, off the furnace geometry.

    fur — BRFFurnaceResults, fin — BRFFurnaceInputs, inp — BRFRefractoryInputs.
    """
    r = inp or BRFRefractoryInputs()
    brick_vol = r.brick_l_mm * r.brick_w_mm * r.brick_h_mm             # H35
    t = r.wall_thickness_mm
    # The length the refractory actually runs: the hot box plus the walled
    # ends, short of the sheet and channel that close the shell.  G19+G20+G21
    ref_len = (fur.effective_length_mm + fin.discharge_refractory_mm
               + fin.charge_refractory_mm)

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

    cast_vol = ref_len * r.castable_depth_mm * r.castable_thick_mm / 1e9       # C53
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
    # HAY behind the side walls, both of them, over the refractory length  G62
    sw_hay = ref_len * r.side_wall_height_mm * r.hay_t_mm / hay_vol

    # ── Hearth ─────────────────────────────────────────────────────
    # One brick course over the whole floor, then the same floor again in
    # IS-8 and in fire brick graded by zone. The 50 in every count is the
    # sheet's own spare allowance.
    ov_width = (fur.effective_width_mm + fin.right_refractory_mm
                + fin.left_refractory_mm)                                      # K58
    hth = r.hearth_thickness_mm
    he_cf_vol = ref_len * ov_width * hth                                       # L53
    he_cf = he_cf_vol / brick_vol + 50                                         # L54
    # IS-8 stops short of the side walls and the HAY behind them
    is8_width = ov_width - r.side_wall_cf_hf_mm - r.side_wall_hay_mm
    he_is8_vol = is8_width * ref_len * hth                                     # L61
    he_is8 = he_is8_vol / brick_vol + 50                                       # L62
    he_50_vol = is8_width * soak_heat * 1000 * hth                             # L64
    he_50 = he_50_vol / brick_vol + 50                                         # L65
    he_60_vol = fur.zone_soaking_m * 1000 * fur.effective_width_mm * hth       # L67
    he_60 = he_60_vol / brick_vol + 50                                         # L68
    he_40_vol = fur.zone_preheating_m * 1000 * hth * fur.effective_width_mm    # L70
    he_40 = he_40_vol / brick_vol + 50                                         # L71

    # ── Refractory block on the ground near the discharge door ─────
    blk_dis = ov_width / r.block_lane_pitch_mm                                 # L75
    blk_soak = (fur.zone_soaking_m * r.block_soak_lanes / r.block_length_m
                + r.block_spare)                                               # L76
    blk_qty = blk_dis + blk_soak                                               # L78

    # ── Aluminium foil, one run per holding-brick row ──────────────
    foil_piece_vol = r.foil_l_m * r.foil_w_m * r.foil_t_m                      # C61
    foil_piece_kg = foil_piece_vol * r.foil_density                            # C62
    foil_len = ref_len * hold_per_width / 1000 + r.foil_spare_m                # C63
    foil_wt = foil_len * foil_piece_kg                                         # C64

    # ── Flue ducts: castable in the annulus, before and after the recuperator
    def _duct(id_m, od_m, length_m):
        vol = 3.14 * (od_m ** 2 - id_m ** 2) / 4 * length_m                    # C69/C77
        wt = vol * r.duct_castable_density * r.duct_castable_wastage           # C71/C79
        return vol, wt, wt / r.castable_bag_kg                                 # C72/C80

    d1_vol, d1_wt, d1_bags = _duct(r.duct1_id_m, r.duct1_od_m, r.duct1_len_m)
    d2_vol, d2_wt, d2_bags = _duct(r.duct2_id_m, r.duct2_od_m, r.duct2_len_m)

    # ── The take-off ───────────────────────────────────────────────
    # Each item gathered across roof, both side walls, discharge wall and
    # hearth, then multiplied by the weight of one piece. The doubled terms
    # are the right and left walls, which are the same wall twice.
    cf_qty = sw_cold * 2 + dis_cold + pre_cold * 2 + he_cf                     # G68
    hf_qty = cf_qty                                                            # G69, same count
    b60_qty = sw_fire * 2 + dis_fire + he_60                                   # G70
    b50_qty = he_50                                                            # G71
    b40_qty = pre_fire40 * 2 + he_40 + tap_fire * 2                            # G72
    is8_qty = he_is8                                                           # G73
    hysil_qty = dis_hay + sw_hay * 2                                           # G74
    # Mortar follows the bricks it beds: one bag per 200.
    accosset_qty = (b60_qty + b50_qty + b40_qty + is8_qty) / r.bricks_per_mortar_bag
    fireclay_qty = (cf_qty + hf_qty) / r.bricks_per_mortar_bag                 # G76

    take_off = [
        ("CF",                          cf_qty,       cf_qty * r.kg_cold_face),
        ("HF",                          hf_qty,       hf_qty * r.kg_hot_face),
        ("60%",                         b60_qty,      b60_qty * r.kg_brick_60),
        ("50%",                         b50_qty,      b50_qty * r.kg_brick_50),
        ("40%",                         b40_qty,      b40_qty * r.kg_brick_40),
        ("IS-8",                        is8_qty,      is8_qty * r.kg_is8),
        ("HYSIL",                       hysil_qty,    hysil_qty * r.kg_hysil),
        ("ACCOSSET 50",                 accosset_qty, accosset_qty * r.mortar_bag_kg),
        ("FIRE CLAY for (CF & HF) paste", fireclay_qty, fireclay_qty * r.mortar_bag_kg),
        ("Block",                       blk_qty,      blk_qty * r.block_kg),
        ("FireCreat castable for gaps",  r.firecreat_bags,
                                        r.firecreat_bags * r.mortar_bag_kg),
    ]
    # The roof items are already weights, so they join the total directly
    # rather than as a count times a piece weight.
    total_kg = (sum(w for _, _, w in take_off)
                + hang_wt + hold_wt60 + hold_wt40      # C40, C44, C47
                + fibre_wt + cast_total                # C52, C56
                + foil_wt + d1_wt + d2_wt)             # C64, C71, C79

    take_off = [(name, round(q, 2), round(w, 2)) for name, q, w in take_off]

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
        side_wall_height_mm=r.side_wall_height_mm,
        side_wall_volume_mm3=sw_vol,
        side_wall_internal_height_mm=r.side_wall_internal_height_mm,
        side_wall_volume_60_mm3=sw_vol2,
        preheat_wall_height_mm=r.preheat_wall_height_mm,
        preheat_wall_volume_mm3=pre_vol,
        preheat_wall_40_height_mm=r.preheat_wall_40_height_mm,
        preheat_wall_40_volume_mm3=pre_vol40,
        tapered_wall_volume_mm3=tap_vol,
        discharge_volume_mm3=dis_vol,
        discharge_volume_60_mm3=dis_vol2,
        hay_volume_mm3=hay_vol,
        hanging_bricks_per_width_raw=round(
            (fur.effective_width_mm + r.hanging_margin_mm) / r.hanging_pitch_mm, 6),
        fibre_roll_kg=r.fibre_roll_kg,
        castable_bag_kg=r.castable_bag_kg,
        hanging_brick_60_kg=r.hanging_brick_60_kg,
        hanging_brick_40_kg=r.hanging_brick_40_kg,
        holding_brick_60_kg=r.holding_brick_60_kg,
        holding_brick_40_kg=r.holding_brick_40_kg,
        discharge_effective_width_mm=fur.effective_width_mm,
        hearth_overall_width_mm=ov_width,
        hearth_cf_volume_mm3=he_cf_vol,
        hearth_cold_face_bricks=_2(he_cf),
        hearth_hot_face_bricks=_2(he_cf),
        hearth_is8_volume_mm3=he_is8_vol,
        hearth_is8_bricks=_2(he_is8),
        hearth_50_volume_mm3=he_50_vol,
        hearth_fire_bricks_50=_2(he_50),
        hearth_60_volume_mm3=he_60_vol,
        hearth_fire_bricks_60=_2(he_60),
        hearth_40_volume_mm3=he_40_vol,
        hearth_fire_bricks_40=_2(he_40),
        block_lane_discharge=_2(blk_dis),
        block_lane_soaking=_2(blk_soak),
        block_pieces=_2(blk_qty),
        block_kg=r.block_kg,
        foil_piece_volume_m3=round(foil_piece_vol, 6),
        foil_piece_kg=round(foil_piece_kg, 4),
        foil_length_m=_2(foil_len),
        foil_weight_kg=_2(foil_wt),
        duct1_volume_m3=round(d1_vol, 3),
        duct1_castable_kg=_2(d1_wt),
        duct1_bags=_2(d1_bags),
        duct2_volume_m3=round(d2_vol, 3),
        duct2_castable_kg=_2(d2_wt),
        duct2_bags=_2(d2_bags),
        take_off=take_off,
        side_wall_hay_pieces=_2(sw_hay),
        total_refractory_kg=_2(total_kg),
        total_refractory_tonne=_2(total_kg / 1000),
    )


# ── Structure ───────────────────────────────────────────────────────────────
# Furnace sheet, section 3 (rows 81-99). The steel the refractory sits in: the
# C-channel band round the outside, the roof and bottom beams, and the plate on
# the walls, the two doors and the bottom tie rod.
#
# Two totals come out of this, and they are not the same number:
#
#   - the Furnace sheet's own L99, 66.59 t, the design weight. It weighs the
#     sections its own headings name — 500 x 180 at 87 kg/m — and carries 10%
#     for wastage;
#   - the Ref.+Str. take-off's L55, 72,172 kg, which is what the vendor
#     breakups quote against. It rounds every quantity up to a whole metre or
#     a whole piece and orders the roof and bottom beams as 350 x 180 and
#     350 x 140 at 53 kg/m — the sections actually stocked.
#
# Both are carried, because both are real: the first is what the furnace
# weighs, the second is what gets bought.

@dataclass
class BRFStructureInputs:
    ms_density: float = 7850.0            # L85, kg/m3
    wastage:    float = 1.1               # the 1.1 on the fabricated weights
    # C-channel, 250 x 80 x 7.2 thk                              C83..C90
    c_channel_kg_per_m:    float = 30.6                          # C84
    c_channel_single_mm:   float = 780 + 1500 + 460 + 115 + 145  # C85, 3.0 m
    c_channel_length_runs: float = 4.0                           # C86
    c_channel_width_runs:  float = 7.0                           # C87
    c_channel_cross_m:     float = (40 + 12) * 3                 # C88, 156 m
    # I-beam over the roof, 500 x 180                            G83..G88
    beam_roof_kg_per_m: float = 87.0      # G84
    beam_pitch_m:       float = 1.5       # G85, one beam per 1.5 m of length
    # I-beam top and bottom, 200 x 100                           G90..G99
    beam_tb_kg_per_m: float = 24.5        # G90
    # Plate joining the beams, 200 x 250 x 12                    L83..L87
    join_plate_l_m: float = 0.2
    join_plate_w_m: float = 0.25
    join_plate_t_m: float = 0.012
    join_plate_spare: float = 10.0        # L87
    # Plate on the side walls, 1500 x 6300 x 8                   L89..L94
    wall_plate_l_m: float = 1.5
    wall_plate_w_m: float = 6.3
    wall_plate_t_m: float = 0.008
    # Charging door                                              C92..C96
    charge_flat_t_m:  float = 0.008       # C92, the flat round the opening
    charge_flat_half: float = 0.5
    charge_door_t_m:  float = 0.016       # C95, the sliding plate
    charge_door_h_m:  float = 1.0
    # Plate on the discharge side                                L96/L97
    discharge_plate_t_m: float = 0.008
    # Plate for the bottom tie rod, 8036 x 250 x 8               C99
    tie_plate_w_m: float = 0.25
    tie_plate_t_m: float = 0.008
    # What the take-off orders the roof and bottom sections at, kg/m
    order_beam_350x180_kg_per_m: float = 53.0
    order_beam_350x140_kg_per_m: float = 53.0


@dataclass
class BRFStructureResults:
    # C-channel                                                  C85..C90
    c_channel_single_m:    float
    c_channel_length_m:    float
    c_channel_width_m:     float
    c_channel_cross_m:     float
    c_channel_total_m:     float   # C89
    c_channel_weight_kg:   float   # C90
    # Roof beam, 500 x 180                                       G85..G88
    beam_roof_count:       float   # G85
    beam_roof_width_m:     float   # G86
    beam_roof_total_m:     float   # G87
    beam_roof_weight_kg:   float   # G88
    # Top and bottom beam, 200 x 100                             G91..G99
    beam_top_count:        float   # G91
    beam_top_length_m:     float   # G92
    beam_top_total_m:      float   # G93
    beam_top_weight_kg:    float   # G94
    beam_bottom_count:     float   # G95
    beam_bottom_total_m:   float   # G96
    beam_bottom_weight_kg: float   # G97
    beam_tb_total_m:       float   # G98
    beam_tb_weight_kg:     float   # G99
    # Plate joining the beams                                    L84..L87
    join_plate_volume_m3:  float   # L84
    join_plate_kg:         float   # L86
    join_plate_count:      float   # L87
    join_plate_weight_kg:  float   # L87 x L86
    # Plate on the side walls                                    L90..L94
    wall_plate_volume_m3:  float   # L90
    wall_plate_kg:         float   # L91
    wall_plate_count:      float   # L92
    wall_plate_side_kg:    float   # L93
    wall_plate_weight_kg:  float   # L94, both walls
    # The doors and the bottom tie rod                      C92..C99, L96/L97
    charge_flat_volume_m3: float      # C92
    charge_flat_kg:        float      # C93
    charge_door_volume_m3: float      # C95
    charge_door_kg:        float      # C96
    discharge_plate_volume_m3: float  # L96
    discharge_plate_kg:    float      # L97
    tie_plate_kg:          float      # C99
    # The two totals
    design_weight_kg:      float   # L99 x 1000
    design_weight_tonne:   float   # L99
    take_off: list = field(default_factory=list)   # Ref.+Str. rows 45-53
    order_weight_kg:    float = 0.0    # Ref.+Str. L55
    order_weight_tonne: float = 0.0
    # Carried so the sheet can show the rate the weight was worked out at
    c_channel_kg_per_m: float = 0.0    # C84


def calculate_structure(fur, fin, ref, inp=None) -> BRFStructureResults:
    """The furnace steel, off the furnace box and the roof brick pitch.

    fur — BRFFurnaceResults, fin — BRFFurnaceInputs,
    ref — BRFRefractoryResults (for the roof brick pitch), inp — inputs.
    """
    s = inp or BRFStructureInputs()
    rho, w = s.ms_density, s.wastage
    len_m = fur.overall_length_mm / 1000.0      # G25
    wid_m = fur.overall_width_mm / 1000.0       # D25
    # The refractory length the channel runs along, same G19+G20+G21
    ref_len_m = (fur.effective_length_mm + fin.discharge_refractory_mm
                 + fin.charge_refractory_mm) / 1000.0

    # ── C-channel band round the outside ───────────────────────────
    c_single = s.c_channel_single_mm / 1000.0                        # C85
    c_len = ref_len_m * s.c_channel_length_runs                      # C86
    c_wid = wid_m * s.c_channel_width_runs                           # C87
    c_total = c_single + c_len + c_wid + s.c_channel_cross_m         # C89
    c_wt = c_total * s.c_channel_kg_per_m * w                        # C90

    # ── Roof beam: one per 1.5 m of length, spanning the width ─────
    b_roof_n = len_m / s.beam_pitch_m                                # G85
    b_roof_m = b_roof_n * wid_m                                      # G87
    b_roof_wt = b_roof_m * s.beam_roof_kg_per_m * w                  # G88

    # ── Top and bottom beam: one top beam per hanging-brick row ────
    b_top_n = ref.hanging_bricks_per_width                           # G91
    b_top_m = b_top_n * len_m                                        # G93
    b_top_wt = b_top_m * s.beam_tb_kg_per_m * w                      # G94
    b_bot_n = b_roof_n                                               # G95
    b_bot_m = len_m * b_bot_n                                        # G96
    b_bot_wt = b_bot_m * s.beam_tb_kg_per_m                          # G97
    b_tb_m = b_bot_m + b_top_m                                       # G98
    b_tb_wt = (b_bot_wt + b_top_wt) * w                              # G99

    # ── Plate joining the beams: two per beam per brick row ────────
    jp_vol = s.join_plate_l_m * s.join_plate_w_m * s.join_plate_t_m  # L84
    jp_kg = jp_vol * rho                                             # L86
    jp_n = b_roof_n * ref.hanging_bricks_per_width * 2 + s.join_plate_spare  # L87

    # ── Plate on the side walls ────────────────────────────────────
    wp_vol = s.wall_plate_l_m * s.wall_plate_w_m * s.wall_plate_t_m  # L90
    wp_kg = wp_vol * rho                                             # L91
    wp_n = (c_single * len_m * s.wall_plate_t_m) / wp_vol            # L92
    wp_side = wp_n * wp_kg                                           # L93
    wp_wt = wp_side * 2                                              # L94

    # ── Charging door: the flat round the opening, then the plate ──
    cf_vol = s.charge_flat_t_m * (c_single * 2 + wid_m) * s.charge_flat_half  # C92
    cf_kg = rho * cf_vol                                             # C93
    cd_vol = ((fur.effective_width_mm + fin.right_refractory_mm) * s.charge_door_h_m
              * s.charge_door_t_m / 1000.0)                          # C95
    cd_kg = cd_vol * rho * w                                         # C96

    # ── Plate on the discharge side, and the bottom tie rod ────────
    dp_vol = c_single * wid_m * s.discharge_plate_t_m                # L96
    dp_kg = dp_vol * rho                                             # L97
    tie_kg = wid_m * s.tie_plate_w_m * s.tie_plate_t_m * b_roof_n * rho * w  # C99

    design_kg = (dp_kg + wp_wt + jp_n * jp_kg + b_tb_wt + b_roof_wt
                 + tie_kg + cd_kg + c_wt)                            # L99

    # ── The take-off the vendors price against ─────────────────────
    # Ref.+Str. rows 45-53. Every quantity rounds up to a whole metre or a
    # whole piece, and the beams are ordered as sections that are stocked.
    #
    # The joining plate is billed at 21 kg apiece, which is the furnace length
    # rounded up: the workbook reads G92 where the plate's own 4.71 kg (L86)
    # belongs. It is kept as the sheet has it so the take-off ties out to the
    # 72,172 kg the breakups price, but it is flagged rather than left silent.
    up = math.ceil
    rows = [
        ("C-Channel",         "250 x 80 x 7.5 Thk.",   up(c_total),  "Metre",
                              up(s.c_channel_kg_per_m), "Kg/m"),
        ("I-Beam",            "350 x 180",             up(b_roof_m), "Metre",
                              s.order_beam_350x180_kg_per_m, "Kg/m"),
        ("I-Beam",            "200 x 100",             up(b_top_m),  "Metre",
                              up(s.beam_tb_kg_per_m), "Kg/m"),
        ("I-Beam",            "350 x 140",             up(b_bot_m),  "Metre",
                              s.order_beam_350x140_kg_per_m, "Kg/m"),
        ("Plate for Joining", "250 x 200 x 12 Thk.",   up(jp_n),     "Nos",
                              up(len_m), "Kg"),
        ("Plate for Walls",   "1500 x 6300 x 8 Thk.",  up(wp_n * 2), "Nos",
                              up(wp_kg), "Kg"),
        ("Plate for Door",    "7000 x 1000 x 16 Thk.", up(cd_kg),    "Kg", 1, "No"),
        ("Plate for Walls",   "3000 x 8000 x 8 Thk.",  up(dp_kg),    "Kg", 1, "No"),
        ("Plate for Bottom",  "8036 x 250 x 8 Thk.",   up(tie_kg),   "Kg", 1, "No"),
    ]
    take_off = [(item, size, qty, uom, unit, unit_uom, round(qty * unit, 2))
                for item, size, qty, uom, unit, unit_uom in rows]
    order_kg = sum(r[-1] for r in take_off)

    def _2(x):
        return round(x, 2)

    return BRFStructureResults(
        c_channel_single_m=_2(c_single),
        c_channel_length_m=_2(c_len),
        c_channel_width_m=_2(c_wid),
        c_channel_cross_m=_2(s.c_channel_cross_m),
        c_channel_total_m=_2(c_total),
        c_channel_weight_kg=_2(c_wt),
        beam_roof_count=_2(b_roof_n),
        beam_roof_width_m=_2(wid_m),
        beam_roof_total_m=_2(b_roof_m),
        beam_roof_weight_kg=_2(b_roof_wt),
        beam_top_count=b_top_n,
        beam_top_length_m=_2(len_m),
        beam_top_total_m=_2(b_top_m),
        beam_top_weight_kg=_2(b_top_wt),
        beam_bottom_count=_2(b_bot_n),
        beam_bottom_total_m=_2(b_bot_m),
        beam_bottom_weight_kg=_2(b_bot_wt),
        beam_tb_total_m=_2(b_tb_m),
        beam_tb_weight_kg=_2(b_tb_wt),
        join_plate_volume_m3=round(jp_vol, 6),
        join_plate_kg=_2(jp_kg),
        join_plate_count=_2(jp_n),
        join_plate_weight_kg=_2(jp_n * jp_kg),
        wall_plate_volume_m3=round(wp_vol, 4),
        wall_plate_kg=_2(wp_kg),
        wall_plate_count=_2(wp_n),
        wall_plate_side_kg=_2(wp_side),
        wall_plate_weight_kg=_2(wp_wt),
        charge_flat_volume_m3=round(cf_vol, 4),
        charge_flat_kg=_2(cf_kg),
        charge_door_volume_m3=round(cd_vol, 4),
        charge_door_kg=_2(cd_kg),
        discharge_plate_volume_m3=round(dp_vol, 4),
        discharge_plate_kg=_2(dp_kg),
        tie_plate_kg=_2(tie_kg),
        design_weight_kg=_2(design_kg),
        design_weight_tonne=_2(design_kg / 1000),
        take_off=take_off,
        order_weight_kg=_2(order_kg),
        order_weight_tonne=_2(order_kg / 1000),
        c_channel_kg_per_m=s.c_channel_kg_per_m,
    )
