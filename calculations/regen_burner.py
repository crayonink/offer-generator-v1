"""
calculations/regen_burner.py
"""

from dataclasses import dataclass


# -------------------------------------------------
# SIZING DATA
# -------------------------------------------------

REGEN_SIZING_TABLE = {

    500: {
        "regen_ms_kg": 139.07, "regen_ss_kg": 11.06,
        "regen_refractory_kg": 422.45, "ceramic_balls_kg": 227.30,
        "burner_block_refractory_kg": 128.76, "burner_ms_kg": 167.53,
        "burner_refractory_kg": 200.39, "total_weight_kg": 1296.57,
        "regen_L": 0.7, "regen_H": 0.55, "regen_W": 0.2,
        "refractory_thk_m": 0.8,
        "burner_length": 0.75, "burner_dia": 0.45,
        "burner_block_inner_dia": 0.30, "burner_block_outer_dia": 0.45,
    },

    1000: {
        "regen_ms_kg": 139.08, "regen_ss_kg": 23.70,
        "regen_refractory_kg": 559.67, "ceramic_balls_kg": 349.07,
        "burner_block_refractory_kg": 158.96, "burner_ms_kg": 198.55,
        "burner_refractory_kg": 237.50, "total_weight_kg": 1666.53,
        "regen_L": 1.0, "regen_H": 0.55, "regen_W": 0.3,
        "refractory_thk_m": 0.8,
        "burner_length": 0.80, "burner_dia": 0.50,
        "burner_block_inner_dia": 0.35, "burner_block_outer_dia": 0.50,
    },

    1500: {
        "regen_ms_kg": 156.47, "regen_ss_kg": 26.07,
        "regen_refractory_kg": 662.48, "ceramic_balls_kg": 448.70,
        "burner_block_refractory_kg": 192.34, "burner_ms_kg": 232.05,
        "burner_refractory_kg": 277.58, "total_weight_kg": 1995.70,
        "regen_L": 1.1, "regen_H": 0.55, "regen_W": 0.3,
        "refractory_thk_m": 0.9,
        "burner_length": 0.85, "burner_dia": 0.55,
        "burner_block_inner_dia": 0.40, "burner_block_outer_dia": 0.55,
    },

    2000: {
        "regen_ms_kg": 284.47, "regen_ss_kg": 28.44,
        "regen_refractory_kg": 1065.43, "ceramic_balls_kg": 1028.09,
        "burner_block_refractory_kg": 276.98, "burner_ms_kg": 314.50,
        "burner_refractory_kg": 376.21, "total_weight_kg": 3374.13,
        "regen_L": 1.2, "regen_H": 0.75, "regen_W": 0.3,
        "refractory_thk_m": 1.2,
        "burner_length": 0.96, "burner_dia": 0.66,
        "burner_block_inner_dia": 0.50, "burner_block_outer_dia": 0.66,
    },

    2500: {
        "regen_ms_kg": 284.47, "regen_ss_kg": 37.92,
        "regen_refractory_kg": 1065.43, "ceramic_balls_kg": 1028.09,
        "burner_block_refractory_kg": 311.57, "burner_ms_kg": 347.46,
        "burner_refractory_kg": 415.63, "total_weight_kg": 3490.58,
        "regen_L": 1.2, "regen_H": 0.75, "regen_W": 0.4,
        "refractory_thk_m": 1.2,
        "burner_length": 1.00, "burner_dia": 0.70,
        "burner_block_inner_dia": 0.55, "burner_block_outer_dia": 0.70,
    },

    3000: {
        "regen_ms_kg": 349.28, "regen_ss_kg": 44.24,
        "regen_refractory_kg": 1375.01, "ceramic_balls_kg": 1556.60,
        "burner_block_refractory_kg": 357.67, "burner_ms_kg": 390.89,
        "burner_refractory_kg": 467.59, "total_weight_kg": 4541.28,
        "regen_L": 1.4, "regen_H": 0.85, "regen_W": 0.4,
        "refractory_thk_m": 1.3,
        "burner_length": 1.05, "burner_dia": 0.75,
        "burner_block_inner_dia": 0.60, "burner_block_outer_dia": 0.75,
    },

    4500: {
        "regen_ms_kg": 505.71, "regen_ss_kg": 55.30,
        "regen_refractory_kg": 1776.45, "ceramic_balls_kg": 2373.08,
        "burner_block_refractory_kg": 459.40, "burner_ms_kg": 485.20,
        "burner_refractory_kg": 580.40, "total_weight_kg": 6235.54,
        "regen_L": 1.4, "regen_H": 1.00, "regen_W": 0.5,
        "refractory_thk_m": 1.6,
        "burner_length": 1.15, "burner_dia": 0.85,
        "burner_block_inner_dia": 0.70, "burner_block_outer_dia": 0.85,
    },

    6000: {
        "regen_ms_kg": 660.56, "regen_ss_kg": 55.30,
        "regen_refractory_kg": 2152.18, "ceramic_balls_kg": 3193.34,
        "burner_block_refractory_kg": 635.85, "burner_ms_kg": 645.28,
        "burner_refractory_kg": 771.89, "total_weight_kg": 8114.41,
        "regen_L": 1.4, "regen_H": 1.10, "regen_W": 0.5,
        "refractory_thk_m": 1.9,
        "burner_length": 1.30, "burner_dia": 1.00,
        "burner_block_inner_dia": 0.75, "burner_block_outer_dia": 1.00,
    },
}

AVAILABLE_KW = sorted(REGEN_SIZING_TABLE.keys())


# -------------------------------------------------
# DATA CLASSES
# -------------------------------------------------

@dataclass
class RegenBurnerInputs:
    power_kw: int
    num_burners: int
    fuel_type: str
    fuel_cv: float


@dataclass
class RegenBurnerResults:
    power_kw: int
    num_burners: int
    total_system_kw: int
    fuel_type: str
    sizing: dict

    ng_flow_nm3hr: float
    air_flow_nm3hr: float


# -------------------------------------------------
# MAIN FUNCTION
# -------------------------------------------------

def calculate_regen_burner(inputs: RegenBurnerInputs):

    base = REGEN_SIZING_TABLE[inputs.power_kw]

    # Flow calculation
    ng_flow_per_burner = (inputs.power_kw * 860) / inputs.fuel_cv
    air_flow_per_burner = ng_flow_per_burner * 10

    ng_flow_nm3hr = ng_flow_per_burner * inputs.num_burners
    air_flow_nm3hr = air_flow_per_burner * inputs.num_burners

    return RegenBurnerResults(
        power_kw=inputs.power_kw,
        num_burners=inputs.num_burners,
        total_system_kw=inputs.power_kw * inputs.num_burners,
        fuel_type=inputs.fuel_type,
        sizing=base,
        ng_flow_nm3hr=ng_flow_nm3hr,
        air_flow_nm3hr=air_flow_nm3hr,
    )