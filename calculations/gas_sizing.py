# calculations/gas_sizing.py

def calculate_gas_flow_nm3hr(
    *,
    firing_rate_mw: float,
    fuel_cv_kcal_nm3: float,
) -> float:
    """
    Calculates natural gas flow from firing rate.

    Formula:
        Gas Flow (Nm3/hr) = MW * 860000 / CV
    """
    return (firing_rate_mw * 860000) / fuel_cv_kcal_nm3
