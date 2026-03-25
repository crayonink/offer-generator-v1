# bom/price_master.py

"""
Central price registry for all BOM items.

RULES:
- Every item used in BOM must exist here
- No silent fallbacks → missing price should raise error
"""


# -------------------------------------------------
# MASTER PRICE DICTIONARY
# -------------------------------------------------

PRICE_MASTER = {

    # -------------------------------------------------
    # CORE ELECTRONICS / IGNITION
    # -------------------------------------------------
    "PILOT BURNER": 12000,
    "REGEN PILOT BURNER": 12000,

    "IGNITION TRANSFORMER": 5500,
    "IGNITION TRANSFORMER (REGEN)": 5500,

    "SEQUENCE CONTROLLER": 10000,
    "SEQUENCE CONTROLLER (REGEN)": 10000,

    "UV SENSOR WITH AIR JACKET": 13000,
    "UV SENSOR WITH AIR JACKET (REGEN)": 13000,


    # -------------------------------------------------
    # COMBUSTION AIR LINE
    # -------------------------------------------------
    "COMPENSATOR": 24000,
    "MOTORIZED CONTROL VALVE": 85000,
    "BUTTERFLY VALVE": 12000,


    # -------------------------------------------------
    # NG LINE
    # -------------------------------------------------
    "BALL VALVE": 1600,
    "BALL VALVE (Pilot Burner)": 1600,
    "BALL VALVE (UV LINE)": 1500,

    "FLEXIBLE HOSE PIPE": 1200,
    "FLEXIBLE HOSE (Pilot Burner)": 1200,
    "FLEXIBLE HOSE (UV LINE)": 1000,

    "SOLENOID VALVE": 6000,
    "PRESSURE REGULATING VALVE": 8000,

    "NG GAS TRAIN": 75000,
    "AGR": 18000,


    # -------------------------------------------------
    # INSTRUMENTATION
    # -------------------------------------------------
    "PRESSURE GAUGE WITH TNV": 4000,
    "PRESSURE GAUGE WITH NV": 4000,
    "PRESSURE SWITCH LOW": 6500,
    "PRESSURE SWITCH HIGH + LOW": 12000,

    "THERMOCOUPLE": 32000,
    "COMPENSATING LEAD": 5000,
    "LIMIT SWITCHES": 2300,
    "TEMPERATURE TRANSMITTER": 13000,

    "CONTROL PANEL": 150000,


    # -------------------------------------------------
    # REVERSING SYSTEM
    # -------------------------------------------------
    "REVERSING VALVE (4-WAY)": 95000,
    "REVERSING VALVE ACTUATOR": 25000,
    "HOT GAS BY-PASS VALVE": 18000,
    "SEQUENCER / PLC FOR REVERSING": 120000,


    # -------------------------------------------------
    # BLOWER
    # -------------------------------------------------
    "REGEN BLOWER": 180000,


    # -------------------------------------------------
    # MISC / CONTROL
    # -------------------------------------------------
    "HYDRAULIC POWER PACK & CYLINDER": 310000,
    "CABLE FOR IGNITION TRANSFORMER": 200,
    "P.PID": 8000,
    "RATIO CONTROLLER": 55000,
}


# -------------------------------------------------
# SAFE ACCESS FUNCTION (IMPORTANT)
# -------------------------------------------------

def get_price(item_name: str) -> float:
    """
    Strict price fetcher.

    Raises error if item is missing instead of silently returning 0.
    """

    price = PRICE_MASTER.get(item_name)

    if price is None:
        raise ValueError(f"❌ Missing price for item: '{item_name}'")

    return price