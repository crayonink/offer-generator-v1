# calculations/process.py
"""
Process / firing rate calculations
FLOW DOMAIN ONLY (Nm3, Nm3/hr)
"""

from dataclasses import dataclass


@dataclass
class ProcessInputs:
    time_hr: float                     # hr
    fuel_consumption_nm3: float        # Nm3
    excess_factor: float = 1.10        # 10% extra
    air_fuel_ratio: float = 19.25       # Nm3 air / Nm3 NG


@dataclass
class ProcessResults:
    calculated_firing_rate_nm3hr: float
    final_firing_rate_nm3hr: float
    air_qty_nm3hr: float
    cfm: float


def calculate_process(inputs: ProcessInputs) -> ProcessResults:
    # Calculated firing rate
    calculated_firing_rate = (
        inputs.fuel_consumption_nm3 / inputs.time_hr
    )

    # Final firing rate (with excess)
    final_firing_rate = calculated_firing_rate * inputs.excess_factor

    # Air quantity
    air_qty = final_firing_rate * inputs.air_fuel_ratio

    # CFM conversion
    cfm = air_qty * 0.588577  # Nm3/hr → CFM

    return ProcessResults(
        calculated_firing_rate_nm3hr=calculated_firing_rate,
        final_firing_rate_nm3hr=final_firing_rate,
        air_qty_nm3hr=air_qty,
        cfm=cfm,
    )
