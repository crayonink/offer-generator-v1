import math

def select_ng_pipe(ng_flow_nm3hr: float, velocity_ms: float = 17) -> dict:
    """
    NG pipe size calculation based on flow and velocity.
    Formula: d = sqrt(Q * 4 / (π * V * 3600)) * 1000  (mm)
    Then round up to nearest standard NB.
    """

    # Standard NB sizes in ascending order
    STANDARD_NB = [15, 20, 25, 32, 40, 50, 65, 80, 100, 125, 150, 200, 250, 300]

    # Calculate inner diameter in mm
    area = ng_flow_nm3hr / (velocity_ms * 3600)       # m²
    diameter_mm = math.sqrt(area * 4 / math.pi) * 1000  # mm

    # Round up to next standard NB
    size_nb = next((nb for nb in STANDARD_NB if nb >= diameter_mm), STANDARD_NB[-1])

    return {
        "size_nb":      size_nb,
        "diameter_mm":  round(diameter_mm, 2),
        "flow_nm3hr":   ng_flow_nm3hr,
        "velocity_ms":  velocity_ms,
    }


# Test
if __name__ == "__main__":
    result = select_ng_pipe(435)
    print(result)
    # {'size_nb': 100, 'diameter_mm': 95.15, 'flow_nm3hr': 435, 'velocity_ms': 17}