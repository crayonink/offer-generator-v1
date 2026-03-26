"""
Regenerative Burner System — Process Calculation
Q(kJ) = mass × Cp × ΔT / efficiency
Power  = Q / (cycle_time × 3600)
Pairs  = ceil(required_kw / 1000)  [each pair = 2 × 500 KW burners]
"""

import math
from dataclasses import dataclass


@dataclass
class RegenInputs:
    material_weight_kg: float
    Ti: float               # Initial temperature °C
    Tf: float               # Final temperature °C
    Cp: float = 0.48        # kJ/kg·°C  (steel default)
    cycle_time_hr: float = 2.0
    efficiency: float = 0.65   # Regen systems recover ~35% of flue heat
    num_pairs_override: int = 0  # 0 = auto-calculate


@dataclass
class RegenResult:
    delta_T: float
    heat_required_kj: float
    heat_required_kcal: float
    required_kw: float
    num_pairs: int
    total_kw: int


def calculate_regen(inputs: RegenInputs) -> RegenResult:
    if inputs.Tf <= inputs.Ti:
        raise ValueError("Final temperature must be greater than initial temperature")
    if inputs.material_weight_kg <= 0:
        raise ValueError("Material weight must be > 0")
    if inputs.cycle_time_hr <= 0:
        raise ValueError("Cycle time must be > 0")

    delta_T = inputs.Tf - inputs.Ti
    heat_required_kj = inputs.material_weight_kg * inputs.Cp * delta_T / inputs.efficiency
    heat_required_kcal = heat_required_kj / 4.187
    required_kw = heat_required_kj / (inputs.cycle_time_hr * 3600)

    if inputs.num_pairs_override > 0:
        num_pairs = inputs.num_pairs_override
    else:
        num_pairs = max(1, math.ceil(required_kw / 1000))

    return RegenResult(
        delta_T=delta_T,
        heat_required_kj=heat_required_kj,
        heat_required_kcal=heat_required_kcal,
        required_kw=required_kw,
        num_pairs=num_pairs,
        total_kw=num_pairs * 1000,
    )
