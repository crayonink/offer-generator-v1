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


def _bore_mm(area_m2: float) -> float:
    """Inner bore from a flow area, the workbook's way: ((4A/3.14)^0.5)*1000."""
    if area_m2 <= 0:
        return 0.0
    return math.sqrt(4 * area_m2 / SHEET_PI) * 1000


def calculate_brf(inp: BRFInputs) -> BRFResults:
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
        air_nb = round_up_to_nb(air_bore) if air_bore > 0 else 0

        gas_area = fuel_flow / inp.gas_velocity_ms / 3600 if inp.gas_velocity_ms else 0.0
        gas_bore = _bore_mm(gas_area)
        gas_nb = round_up_to_nb(gas_bore) if gas_bore > 0 else 0

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
            gas_line_nb_as_sheet      = round_up_to_nb(sheet_bore) if sheet_bore > 0 else 0,
        ))

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
