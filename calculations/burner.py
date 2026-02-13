# burner calculations module
"""
Burner & heat load calculations
Pure logic module — no Streamlit, no pandas, no Excel
"""

from dataclasses import dataclass


@dataclass
class BurnerInputs:
    Ti: float                      # Initial temperature (°C)
    Tf: float                      # Final temperature (°C)
    refractory_weight: float       # Kg
    fuel_cv: float                 # Kcal/Nm3
    time_taken_hr: float           # Hours


@dataclass
class BurnerResults:
    avg_temp_rise: float
    firing_rate_kcal: float
    heat_load_kcal: float
    fuel_consumption_nm3: float
    calculated_firing_rate_nm3hr: float
    extra_firing_rate_nm3hr: float
    final_firing_rate_mw: float
    air_qty_nm3hr: float
    cfm: float
    blower_hp: float


def calculate_burner(inputs: BurnerInputs) -> BurnerResults:
    """
    Matches your Excel logic EXACTLY
    """

    # Average temperature to be raised
    avg_temp_rise = ((inputs.Tf + 200) / 2) - ((inputs.Ti + 100) / 2)

    # Firing rate
    firing_rate_kcal = inputs.refractory_weight * 0.25 * avg_temp_rise

    # Heat load
    heat_load_kcal = firing_rate_kcal / 0.52

    # Fuel consumption
    fuel_consumption_nm3 = heat_load_kcal / inputs.fuel_cv

    # Calculated firing rate (Nm3/hr)
    calculated_firing_rate_nm3hr = fuel_consumption_nm3 / inputs.time_taken_hr

    # 10% extra firing rate
    extra_firing_rate_nm3hr = calculated_firing_rate_nm3hr * 1.1

    # Final firing rate in MW
    final_firing_rate_mw = (
        extra_firing_rate_nm3hr * inputs.fuel_cv
    ) / (860 * 1000)

    # Air quantity
    air_qty_nm3hr = (
        inputs.fuel_cv * extra_firing_rate_nm3hr * 118
    ) / 100000

    # CFM
    cfm = air_qty_nm3hr / 1.7

    # Blower size (HP)
    blower_hp = cfm / 114

    return BurnerResults(
        avg_temp_rise=avg_temp_rise,
        firing_rate_kcal=firing_rate_kcal,
        heat_load_kcal=heat_load_kcal,
        fuel_consumption_nm3=fuel_consumption_nm3,
        calculated_firing_rate_nm3hr=calculated_firing_rate_nm3hr,
        extra_firing_rate_nm3hr=extra_firing_rate_nm3hr,
        final_firing_rate_mw=final_firing_rate_mw,
        air_qty_nm3hr=air_qty_nm3hr,
        cfm=cfm,
        blower_hp=blower_hp,
    )
