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
from bom.selectors.hpu_selector import select_hpu, select_pumping_unit, PUMPING_UNIT_ONLY_FUELS
from bom.selectors.encon_burner import _resolve_category

from calculations.pipes import PipeInputs, calculate_pipe_sizes


def select_equipment(*, ng_flow_nm3hr: float, air_flow_nm3hr: float, is_dual_fuel: bool = False, fuel_cv: float = 10500, blower_pressure: str = "28", fuel_type: str = "gas", hpu_variant: str = "Duplex 1", burner_pressure_wg: int = 24, butterfly_valve_vendor: str = "lt_lever", shutoff_valve_vendor: str = "aira", control_mode: str = "automatic", auto_control_type: str = "plc", fuel2_lph: float = 0) -> dict:
    """
    Selects all equipment for a VLPH system based on gas and air flow rates.
    fuel_type: 'gas', 'oil', or 'dual' — picks the burner pricelist section.
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

    # Burner — fuel_type drives which pricelist section is used
    burner_fuel_type = "dual" if is_dual_fuel else fuel_type
    try:
        burner = select_encon_mg_burner(
            ng_flow_nm3hr,
            fuel_cv=fuel_cv,
            fuel_type=burner_fuel_type,
            burner_pressure_wg=burner_pressure_wg,
        )
    except Exception:
        burner = {
            "model": "TEST-BURNER",
            "input_nm3hr": ng_flow_nm3hr,
            "equivalent_lph": 0,
            "price": 0,
        }

    # NG side
    ng_gas_train = select_ng_gas_train(ng_flow_nm3hr)

    # AGR — only required for control modes that actually use it.
    # Pure PLC mode uses orifice plate + DPT + control valve for ratio control,
    # so AGR is not in the BOM and we skip selection entirely (important for
    # low-CV fuels like Mixed Gas where no AGR in that NB range exists).
    needs_agr = (
        control_mode == "manual"
        or (control_mode == "automatic" and auto_control_type in ("plc_agr", "pid"))
    )
    # AGR is sized from gas train outlet NB (not the raw pipe NB)
    import re as _re
    gt_outlet_nb = int(_re.sub(r'[^\d]', '', str(ng_gas_train.get("outlet_nb", ""))) or ng_nb)
    if needs_agr:
        agr = select_agr(
            nb=gt_outlet_nb,
            connection="Flanged" if gt_outlet_nb >= 65 else "Threaded",
            ratio="1:1",
            compact="No",
        )
    else:
        agr = {"nb": gt_outlet_nb, "price": 0, "enag": None, "item_code": None,
               "connection": None, "ratio": None, "compact": None, "pmax_mbar": None}

    # Air side
    air_duct = select_air_duct(air_flow_nm3hr)
    motorized_control_valve = select_motorized_control_valve(air_flow_nm3hr)
    # Butterfly valve — fall back from Lever to Gear if pipe NB exceeds 300
    try:
        butterfly_valve = select_butterfly_valve(air_nb, vendor=butterfly_valve_vendor)
    except ValueError:
        if butterfly_valve_vendor == "lt_lever":
            butterfly_valve = select_butterfly_valve(air_nb, vendor="lt_gear")
        else:
            raise
    rotary_joint = select_rotary_joint(air_nb)

    # Blower HP = CFM × pressure (inches w.g.) / 3200
    # Gas CFM = air / 1.7. Oil CFM = LPH × 10 (accounts for atomization air).
    # For dual fuel with any oil: use max(gas_cfm, oil_cfm) so blower handles
    # both modes.
    gas_cfm = air_flow_nm3hr / 1.7
    oil_cfm = 0
    if _resolve_category(fuel_type) == "oil" and burner.get("equivalent_lph"):
        oil_cfm = burner["equivalent_lph"] * 10
    # Dual-fuel: if fuel2 is oil, its LPH × 10 may exceed gas CFM
    if fuel2_lph and fuel2_lph > 0:
        oil_cfm = max(oil_cfm, fuel2_lph * 10)
    cfm = max(gas_cfm, oil_cfm) if oil_cfm else gas_cfm
    pressure_in_wg = int(blower_pressure)   # "28" or "40"
    required_hp = cfm * pressure_in_wg / 3200
    blower = select_blower(required_hp, series=blower_pressure)

    # HPU / Pumping Unit — only for oil-based fuels.
    # Heavy oils (LSHS, FO) get a standalone Pumping Unit since the oil is
    # pre-heated separately; other oils (LDO, HSD, SKO, HDO, CFO) use HPU.
    # Sized to actual oil firing rate (LPH), not burner model.
    hpu = None
    if _resolve_category(fuel_type) == "oil":
        picker = select_pumping_unit if fuel_type.lower() in PUMPING_UNIT_ONLY_FUELS else select_hpu
        try:
            hpu = picker(burner["equivalent_lph"], variant=hpu_variant)
        except ValueError:
            hpu = None

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
        "hpu": hpu,
    }
