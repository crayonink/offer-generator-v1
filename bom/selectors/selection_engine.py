from enum import Enum

# -------------------------------
# SELECTORS
# -------------------------------
from bom.selectors.encon_burner import select_encon_mg_burner
from bom.selectors.ng_gas_train import select_ng_gas_train
from bom.selectors.agr_selector import select_agr

from bom.selectors.blower_selector import (
    select_blower,
    calculate_blower_hp,
)

from bom.selectors.air_valve_selector import (
    select_motorized_control_valve,
    select_butterfly_valve,
)

from bom.selectors.air_duct_selector import select_air_duct
from bom.selectors.rotary_joint_selector import select_rotary_joint

# -------------------------------
# PIPE CALCULATIONS
# -------------------------------
from calculations.pipes import PipeInputs, calculate_pipe_sizes


# =========================================================
# SYSTEM TYPE
# =========================================================
class SystemType(Enum):
    VLPH = "vlph"
    REGEN = "regen"


# =========================================================
# FLOW RESOLUTION (CORE DIFFERENCE)
# =========================================================
def resolve_flows(system_type, capacity_kw, ng_flow, air_flow):
    """
    Determines how flow is distributed.

    VLPH  → single burner (no split)
    REGEN → multiple burners (split flow)
    """

    if system_type == SystemType.VLPH:
        return {
            "num_burners": 1,
            "per_burner_ng_flow": ng_flow,
            "per_burner_air_flow": air_flow,
        }

    elif system_type == SystemType.REGEN:
        num_burners = max(1, int(capacity_kw // 500))

        return {
            "num_burners": num_burners,
            "per_burner_ng_flow": ng_flow / num_burners,
            "per_burner_air_flow": air_flow / num_burners,
        }

    else:
        raise ValueError(f"Unsupported system type: {system_type}")


# =========================================================
# MAIN SELECTION ENGINE
# =========================================================
def select_equipment(
    *,
    system_type: SystemType,
    capacity_kw: float,
    ng_flow_nm3hr: float,
    air_flow_nm3hr: float,
) -> dict:
    """
    Central equipment selector.

    RULES:
    - Pipe sizing is the source of truth
    - Selection uses correct flow (split or full)
    - No quantity logic here (BOM handles that)
    """

    # -------------------------------------------------
    # VALIDATION
    # -------------------------------------------------
    if ng_flow_nm3hr <= 0:
        raise ValueError("ng_flow_nm3hr must be > 0")

    if air_flow_nm3hr <= 0:
        raise ValueError("air_flow_nm3hr must be > 0")

    if capacity_kw <= 0:
        raise ValueError("capacity_kw must be > 0")

    # -------------------------------------------------
    # 1️⃣ FLOW RESOLUTION (KEY STEP)
    # -------------------------------------------------
    flow_data = resolve_flows(
        system_type,
        capacity_kw,
        ng_flow_nm3hr,
        air_flow_nm3hr,
    )

    num_burners = flow_data["num_burners"]
    per_ng_flow = flow_data["per_burner_ng_flow"]
    per_air_flow = flow_data["per_burner_air_flow"]

    # -------------------------------------------------
    # 2️⃣ PIPE CALCULATION (SYSTEM LEVEL)
    # -------------------------------------------------
    pipe_results = calculate_pipe_sizes(
        PipeInputs(
            ng_flow_nm3hr=ng_flow_nm3hr,
            air_flow_nm3hr=air_flow_nm3hr,
        )
    )

    ng_nb = pipe_results.ng_pipe_nb
    air_nb = pipe_results.air_pipe_nb

    # -------------------------------------------------
    # 3️⃣ BURNER SELECTION (PER BURNER)
    # -------------------------------------------------
    try:
        burner = select_encon_mg_burner(per_ng_flow)
    except Exception:
        burner = {
            "model": "TEST-BURNER",
            "max_flow_nm3hr": per_ng_flow,
            "price": 0,
        }

    # -------------------------------------------------
    # 4️⃣ NG SIDE EQUIPMENT (SYSTEM LEVEL)
    # -------------------------------------------------
    ng_gas_train = select_ng_gas_train(ng_flow_nm3hr)

    agr = select_agr(
        nb=ng_nb,
        connection="Flanged" if ng_nb >= 65 else "Threaded",
        ratio="1:1",
        compact="No",
    )

    # -------------------------------------------------
    # 5️⃣ AIR SIDE EQUIPMENT (SYSTEM LEVEL)
    # -------------------------------------------------
    air_duct = select_air_duct(air_flow_nm3hr)

    motorized_control_valve = select_motorized_control_valve(
        air_flow_nm3hr
    )

    butterfly_valve = select_butterfly_valve(air_nb)

    rotary_joint = select_rotary_joint(air_nb)

    # -------------------------------------------------
    # 6️⃣ BLOWER (SYSTEM LEVEL — FIXED LOGIC)
    # -------------------------------------------------
    required_hp = calculate_blower_hp(air_flow_nm3hr)
    blower = select_blower(required_hp)

    # -------------------------------------------------
    # RETURN CLEAN STRUCTURE
    # -------------------------------------------------
    return {
        "system_type": system_type.value,
        "capacity_kw": capacity_kw,
        "num_burners": num_burners,

        # Pipe (source of truth)
        "pipe": pipe_results,

        # Burner (per burner logic)
        "burner": {
            "data": burner,
            "basis": "per_burner",
        },

        # NG side
        "ng_gas_train": {
            "data": ng_gas_train,
            "basis": "system",
        },
        "agr": {
            "data": agr,
            "basis": "per_burner",  # important for Regen scaling
        },

        # Air side
        "blower": {
            "data": blower,
            "basis": "system",
        },
        "air_duct": {
            "data": air_duct,
            "basis": "system",
        },
        "motorized_control_valve": {
            "data": motorized_control_valve,
            "basis": "system",
        },
        "butterfly_valve": {
            "data": butterfly_valve,
            "basis": "system",
        },
        "rotary_joint": {
            "data": rotary_joint,
            "basis": "system",
        },
    }