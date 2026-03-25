"""
calculations/process.py

Process / firing rate calculations
FLOW DOMAIN ONLY (Nm3, Nm3/hr)
"""

from dataclasses import dataclass


# Nm3/hr → CFM conversion factor
NM3HR_TO_CFM = 0.588577


@dataclass
class ProcessInputs:
    time_hr: float                      # Process time (hr)
    fuel_consumption_nm3: float         # Total fuel consumption (Nm3)
    excess_factor: float = 1.10         # Excess firing factor (10%)
    air_fuel_ratio: float = 19.25       # Nm3 air / Nm3 NG


@dataclass
class ProcessResults:
    calculated_firing_rate_nm3hr: float
    final_firing_rate_nm3hr: float
    air_qty_nm3hr: float
    cfm: float


def calculate_process(inputs: ProcessInputs) -> ProcessResults:
    """
    Calculates firing rate and air requirement for the process.

    Returns:
        ProcessResults with firing rate (Nm3/hr), air flow (Nm3/hr),
        and air flow converted to CFM.
    """

    if inputs.time_hr <= 0:
        raise ValueError("time_hr must be greater than zero")

    if inputs.fuel_consumption_nm3 < 0:
        raise ValueError("fuel_consumption_nm3 cannot be negative")

    # Calculated firing rate
    calculated_firing_rate = (
        inputs.fuel_consumption_nm3 / inputs.time_hr
    )

    # Final firing rate (with excess factor)
    final_firing_rate = calculated_firing_rate * inputs.excess_factor

    # Required combustion air
    air_qty = final_firing_rate * inputs.air_fuel_ratio

    # Convert Nm3/hr → CFM
    cfm = air_qty * NM3HR_TO_CFM

    return ProcessResults(
        calculated_firing_rate_nm3hr=calculated_firing_rate,
        final_firing_rate_nm3hr=final_firing_rate,
        air_qty_nm3hr=air_qty,
        cfm=cfm,
    )