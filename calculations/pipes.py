# calculations/pipes.py
"""
NG & Combustion Air pipe sizing calculations
Pure pipe logic
"""

from dataclasses import dataclass
import math


@dataclass
class PipeInputs:
    ng_flow_nm3hr: float
    air_flow_nm3hr: float
    ng_velocity_ms: float = 12.7
    air_velocity_ms: float = 15.0


@dataclass
class PipeResults:
    ng_pipe_inner_dia_mm: float
    air_pipe_inner_dia_mm: float


def calculate_pipe_sizes(inputs: PipeInputs) -> PipeResults:
    # NG pipe
    ng_flow_m3s = inputs.ng_flow_nm3hr / 3600
    ng_area = ng_flow_m3s / inputs.ng_velocity_ms
    ng_dia_mm = math.sqrt((4 * ng_area) / math.pi) * 1000

    # Air pipe
    air_flow_m3s = inputs.air_flow_nm3hr / 3600
    air_area = air_flow_m3s / inputs.air_velocity_ms
    air_dia_mm = math.sqrt((4 * air_area) / math.pi) * 1000

    return PipeResults(
        ng_pipe_inner_dia_mm=ng_dia_mm,
        air_pipe_inner_dia_mm=air_dia_mm,
    )
