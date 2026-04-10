"""
NG & Combustion Air pipe sizing calculations
Pure pipe logic
"""

from dataclasses import dataclass
import math


# -------------------------------------------------
# STANDARD PIPE NB SIZES (mm)
# -------------------------------------------------
STANDARD_PIPE_NB = [15, 20, 25, 32, 40, 50, 65, 80, 100, 125, 150, 200, 250, 300, 350, 400, 450, 500, 600]


def round_up_to_nb(diameter_mm: float) -> int:
    """Round calculated inner diameter up to the next standard pipe NB."""
    for nb in STANDARD_PIPE_NB:
        if diameter_mm <= nb:
            return nb
    raise ValueError(
        f"Required pipe size ({diameter_mm:.1f} mm) exceeds maximum available NB (350)"
    )


# -------------------------------------------------
# INPUT DATA STRUCTURE
# -------------------------------------------------
@dataclass
class PipeInputs:
    ng_flow_nm3hr: float
    air_flow_nm3hr: float
    ng_velocity_ms: float = 12.7   # Design velocity for all gas fuels (NG/LPG/COG/BG/RLNG/MG)
    air_velocity_ms: float = 15.0


# -------------------------------------------------
# OUTPUT DATA STRUCTURE
# -------------------------------------------------
@dataclass
class PipeResults:
    ng_pipe_inner_dia_mm: float
    ng_pipe_nb: int
    ng_actual_velocity_ms: float

    air_pipe_inner_dia_mm: float
    air_pipe_nb: int
    air_actual_velocity_ms: float


# -------------------------------------------------
# HELPER
# -------------------------------------------------
def _velocity(flow_m3s, diameter_mm):
    area = math.pi * (diameter_mm / 1000) ** 2 / 4
    return flow_m3s / area


# -------------------------------------------------
# MAIN CALCULATION FUNCTION
# -------------------------------------------------
def calculate_pipe_sizes(inputs: PipeInputs) -> PipeResults:

    # -----------------------------
    # VALIDATION
    # -----------------------------
    if inputs.ng_flow_nm3hr <= 0:
        raise ValueError("ng_flow_nm3hr must be > 0")

    if inputs.air_flow_nm3hr <= 0:
        raise ValueError("air_flow_nm3hr must be > 0")

    # -----------------------------
    # NG PIPE CALCULATION
    # -----------------------------
    ng_flow_m3s = inputs.ng_flow_nm3hr / 3600
    ng_area = ng_flow_m3s / inputs.ng_velocity_ms
    ng_dia_mm = math.sqrt((4 * ng_area) / math.pi) * 1000

    ng_nb = round_up_to_nb(ng_dia_mm)

    # Recalculate actual velocity after rounding
    ng_actual_velocity = _velocity(ng_flow_m3s, ng_nb)

    # -----------------------------
    # AIR PIPE CALCULATION
    # -----------------------------
    air_flow_m3s = inputs.air_flow_nm3hr / 3600
    air_area = air_flow_m3s / inputs.air_velocity_ms
    air_dia_mm = math.sqrt((4 * air_area) / math.pi) * 1000

    air_nb = round_up_to_nb(air_dia_mm)

    air_actual_velocity = _velocity(air_flow_m3s, air_nb)

    # -----------------------------
    # RESULT
    # -----------------------------
    return PipeResults(
        ng_pipe_inner_dia_mm=ng_dia_mm,
        ng_pipe_nb=ng_nb,
        ng_actual_velocity_ms=ng_actual_velocity,

        air_pipe_inner_dia_mm=air_dia_mm,
        air_pipe_nb=air_nb,
        air_actual_velocity_ms=air_actual_velocity,
    )