"""
NG & Combustion Air pipe sizing calculations
Pure pipe logic
"""

from dataclasses import dataclass
import math


# -------------------------------------------------
# AGR AVAILABLE NB (FROM SPECIFICATION FILE)
# AGR does NOT exist beyond 100 NB
# -------------------------------------------------
AVAILABLE_AGR_NB = [15, 20, 25, 32, 40, 50, 65, 80, 100]


def round_to_available_agr_nb(diameter_mm: float) -> int:
    """
    Round calculated inner diameter to next available AGR NB.
    Raises error if required size exceeds 100 NB.
    """
    for nb in AVAILABLE_AGR_NB:
        if diameter_mm <= nb:
            return nb

    raise ValueError(
        "Required pipe size exceeds maximum available AGR size (100 NB)"
    )


# -------------------------------------------------
# INPUT DATA STRUCTURE
# -------------------------------------------------
@dataclass
class PipeInputs:
    ng_flow_nm3hr: float
    air_flow_nm3hr: float
    ng_velocity_ms: float = 17.0   # Locked as per approved calculation sheet
    air_velocity_ms: float = 15.0


# -------------------------------------------------
# OUTPUT DATA STRUCTURE
# -------------------------------------------------
@dataclass
class PipeResults:
    ng_pipe_inner_dia_mm: float
    ng_pipe_nb: int
    air_pipe_inner_dia_mm: float


# -------------------------------------------------
# MAIN CALCULATION FUNCTION
# -------------------------------------------------
def calculate_pipe_sizes(inputs: PipeInputs) -> PipeResults:
    # -----------------------------
    # NG PIPE CALCULATION
    # -----------------------------
    ng_flow_m3s = inputs.ng_flow_nm3hr / 3600
    ng_area = ng_flow_m3s / inputs.ng_velocity_ms
    ng_dia_mm = math.sqrt((4 * ng_area) / math.pi) * 1000

    # Round to next available AGR NB
    ng_nb = round_to_available_agr_nb(ng_dia_mm)

    # -----------------------------
    # AIR PIPE CALCULATION
    # -----------------------------
    air_flow_m3s = inputs.air_flow_nm3hr / 3600
    air_area = air_flow_m3s / inputs.air_velocity_ms
    air_dia_mm = math.sqrt((4 * air_area) / math.pi) * 1000

    return PipeResults(
        ng_pipe_inner_dia_mm=ng_dia_mm,
        ng_pipe_nb=ng_nb,
        air_pipe_inner_dia_mm=air_dia_mm,
    )