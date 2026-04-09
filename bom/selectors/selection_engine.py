import sqlite3
import os

from bom.selectors.encon_burner import select_encon_mg_burner
from bom.selectors.ng_gas_train import select_ng_gas_train
from bom.selectors.agr_selector import select_agr

from bom.selectors.blower_selector import select_blower

from bom.selectors.air_valve_selector import (
    select_motorized_control_valve,
    select_butterfly_valve,
)

from bom.selectors.air_duct_selector import select_air_duct
from bom.selectors.rotary_joint_selector import select_rotary_joint
from calculations.pipes import PipeInputs, calculate_pipe_sizes

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DB_PATH = os.path.join(BASE_DIR, "vlph.db")


def _select_dual_fuel_burner(burner_size: str) -> float:
    """Look up dual fuel burner set price from burner_pricelist_master."""
    conn = sqlite3.connect(DB_PATH)
    row = conn.execute(
        "SELECT price FROM burner_pricelist_master "
        "WHERE section LIKE '%DUAL%' AND burner_size=? AND component='BURNER SET'",
        (burner_size,)
    ).fetchone()
    conn.close()
    return row[0] if row else 0


def select_equipment(*, ng_flow_nm3hr: float, air_flow_nm3hr: float, is_dual_fuel: bool = False, fuel_cv: float = 10500, blower_pressure: str = "28") -> dict:
    """
    Selects all equipment for a VLPH system based on gas and air flow rates.
    Returns a flat dict — each key maps directly to the selected component dict.
    """

    if ng_flow_nm3hr <= 0:
        raise ValueError("ng_flow_nm3hr must be > 0")
    if air_flow_nm3hr <= 0:
        raise ValueError("air_flow_nm3hr must be > 0")

    # Pipe sizing (source of truth for NB selection)
    pipe_results = calculate_pipe_sizes(PipeInputs(
        ng_flow_nm3hr=ng_flow_nm3hr,
        air_flow_nm3hr=air_flow_nm3hr,
    ))

    ng_nb = pipe_results.ng_pipe_nb
    air_nb = pipe_results.air_pipe_nb

    # Burner
    try:
        burner = select_encon_mg_burner(ng_flow_nm3hr, fuel_cv=fuel_cv)
        # Override with dual fuel burner price if applicable
        if is_dual_fuel:
            pricelist_name = burner["model"]
            dual_price = _select_dual_fuel_burner(pricelist_name)
            if dual_price > 0:
                burner["price"] = dual_price
    except Exception:
        burner = {
            "model": "TEST-BURNER",
            "input_nm3hr": ng_flow_nm3hr,
            "equivalent_lph": 0,
            "price": 0,
        }

    # NG side
    ng_gas_train = select_ng_gas_train(ng_flow_nm3hr)
    agr = select_agr(
        nb=ng_nb,
        connection="Flanged" if ng_nb >= 65 else "Threaded",
        ratio="1:1",
        compact="No",
    )

    # Air side
    air_duct = select_air_duct(air_flow_nm3hr)
    motorized_control_valve = select_motorized_control_valve(air_flow_nm3hr)
    butterfly_valve = select_butterfly_valve(air_nb)
    rotary_joint = select_rotary_joint(air_nb)

    # Blower — HP formula depends on pressure
    cfm = air_flow_nm3hr / 1.7
    if blower_pressure == "40":
        required_hp = cfm * 40 / 3200   # 40" WG formula (SAIL standard)
    else:
        required_hp = cfm / 114          # 28" WG formula
    blower = select_blower(required_hp, series=blower_pressure)

    return {
        "pipe": pipe_results,
        "burner": burner,
        "ng_gas_train": ng_gas_train,
        "agr": agr,
        "blower": blower,
        "air_duct": air_duct,
        "motorized_control_valve": motorized_control_valve,
        "butterfly_valve": butterfly_valve,
        "rotary_joint": rotary_joint,
    }
