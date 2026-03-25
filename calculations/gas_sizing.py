# calculations/gas_sizing.py

from dataclasses import dataclass


# -------------------------------------------------
# DATA CLASSES
# -------------------------------------------------

@dataclass
class GasSizingInputs:
    firing_rate_mw: float
    fuel_cv_kcal_nm3: float


@dataclass
class GasSizingResults:
    gas_flow_nm3hr: float


# -------------------------------------------------
# MAIN FUNCTION
# -------------------------------------------------

def calculate_gas_flow_nm3hr(inputs: GasSizingInputs) -> GasSizingResults:
    """
    Calculates natural gas flow from firing rate.

    Formula:
        Gas Flow (Nm3/hr) = MW * 860000 / CV

    Where:
        MW → thermal load
        CV → kcal/Nm3
    """

    if inputs.firing_rate_mw <= 0:
        raise ValueError("firing_rate_mw must be greater than zero")

    if inputs.fuel_cv_kcal_nm3 <= 0:
        raise ValueError("fuel_cv_kcal_nm3 must be greater than zero")

    gas_flow = (inputs.firing_rate_mw * 860000) / inputs.fuel_cv_kcal_nm3

    return GasSizingResults(
        gas_flow_nm3hr=gas_flow
    )