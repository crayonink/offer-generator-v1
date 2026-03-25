import math


def select_air_duct(air_flow_nm3hr: float, velocity_ms: float = 15.0) -> dict:
    """
    Select combustion air duct size based on calculated diameter.
    Uses formula:
    d = sqrt(4Q / (π V 3600)) * 1000
    """

    # Calculate diameter in mm
    diameter_mm = math.sqrt(
        (air_flow_nm3hr * 4) / (math.pi * velocity_ms * 3600)
    ) * 1000

    # Standard available duct sizes
    STANDARD_NB = [200, 250, 300, 350]

    # Select next available NB >= calculated diameter
    nb = next((n for n in STANDARD_NB if n >= diameter_mm), STANDARD_NB[-1])

    return {
        "nb": nb,
        "calculated_diameter_mm": round(diameter_mm, 2),
        "velocity_ms": velocity_ms,
    }