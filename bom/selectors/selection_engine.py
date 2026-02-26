from bom.selectors.encon_burner import select_encon_mg_burner
from bom.selectors.ng_gas_train import select_ng_gas_train
from bom.selectors.agr_selector import select_agr
from bom.selectors.blower_selector import select_blower

from bom.selectors.air_valve_selector import (
    select_motorized_control_valve,
    select_butterfly_valve,
)

from bom.selectors.air_duct_selector import select_air_duct
from bom.selectors.ng_pipe_selector import select_ng_pipe
from bom.selectors.rotary_joint_selector import select_rotary_joint
from bom.selectors.compensator_selector import select_compensator


def select_equipment(*, ng_flow_nm3hr: float, air_flow_nm3hr: float) -> dict:
    """
    Central equipment selector.
    STRICT LEGACY REPLICATION MODE.
    DB-backed dynamic logic.
    """

    # -------------------------------------------------
    #  1️⃣ Burner FIRST (anchor component)
    # -------------------------------------------------
    burner = select_encon_mg_burner(ng_flow_nm3hr)

    # -------------------------------------------------
    #  2️⃣ NG side follows burner
    # -------------------------------------------------
    ng_gas_train = select_ng_gas_train(
        ng_flow_nm3hr,
        burner["model"],
    )

    agr = select_agr(ng_flow_nm3hr)

    # -------------------------------------------------
    #  3️⃣ Air side logic
    # -------------------------------------------------
    air_duct = select_air_duct(air_flow_nm3hr)

    motorized_control_valve = select_motorized_control_valve(
        air_flow_nm3hr
    )

    butterfly_valve = select_butterfly_valve(
        air_duct["nb"]
    )

    rotary_joint = select_rotary_joint(
        air_flow_nm3hr
    )

    compensator = select_compensator(
        air_flow_nm3hr
    )

    blower = select_blower(
        air_flow_nm3hr
    )

    # -------------------------------------------------
    # RETURN STRICT SCHEMA
    # -------------------------------------------------
    return {

        # Burner Package
        "encon_burner": burner,
        "ng_gas_train": ng_gas_train,
        "agr": agr,

        #  Air Package
        "blower": blower,
        "air_duct": air_duct,
        "motorized_control_valve": motorized_control_valve,
        "butterfly_valve": butterfly_valve,
        "rotary_joint": rotary_joint,
        "compensator": compensator,

        # Misc
        "ng_pipe": select_ng_pipe(ng_flow_nm3hr),
    }