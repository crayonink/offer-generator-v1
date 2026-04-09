import pandas as pd

from bom.static_items import static_items
from bom.price_master import get_price
from bom.ladle_params import get_vlph_params


BOUGHT_OUT_EXCLUDE_ITEMS = {
    "RATIO CONTROLLER",
}

FUEL_NAMES = {
    "ng": "NG", "lpg": "LPG", "cog": "COG",
    "bg": "BG", "rlng": "RLNG", "ldo": "LDO",
}


def _row(media: str, item: str, ref: str, qty, unit_price_override=None):
    qty = qty if qty else 1
    if unit_price_override is not None:
        unit_price = unit_price_override
    else:
        try:
            unit_price = get_price(item)
        except ValueError:
            unit_price = 0

    if unit_price == 0:
        print(f"WARNING: No price found for '{item}'")

    return (media, item, ref, qty, unit_price, unit_price * qty)


OIL_FUELS = {"ldo"}
GAS_FUELS = {"ng", "rlng", "lpg", "cog", "bg"}


def _fuel_line_rows(label: str, fuel_type: str, equipment: dict,
                    control_mode: str = "automatic", auto_control_type: str = "plc"):
    """Generate fuel line BOM rows for a single fuel."""
    media = f"{label} LINE"
    rows = []

    # Gas train (gas fuels only)
    if fuel_type in GAS_FUELS:
        gas_train_name = f'GAS TRAIN {equipment["ng_gas_train"]["max_flow"]:.0f} NM3/Hr'
        rows.append(_row(
            media, gas_train_name,
            f'{equipment["ng_gas_train"]["inlet_nb"]} x '
            f'{equipment["ng_gas_train"]["outlet_nb"]} NB',
            1, unit_price_override=equipment["ng_gas_train"]["price"],
        ))

    # Oil line size is always 20 NB
    oil_nb = 20

    # Control-type-specific instrumentation
    if control_mode == "automatic":
        if auto_control_type == "plc":
            # PLC: gas → orifice plate + DPT + control valve, oil → flowmeter + control valve
            if fuel_type in GAS_FUELS:
                rows += [
                    _row(media, "ORIFICE PLATE", f'{equipment["agr"]["nb"]} NB', 1),
                    _row(media, "DPT", "Output 4-20 mA", 1),
                    _row(media, "PNEUMATIC CONTROL VALVE (Gas)", f'{equipment["agr"]["nb"]} NB', 1),
                ]
            elif fuel_type in OIL_FUELS:
                rows += [
                    _row(media, "FLOWMETER", f'{oil_nb} NB', 1),
                    _row(media, "MOTORIZED CONTROL VALVE 25NB (Globe)", "CAIR", 1),
                    _row(media, "SOLENOID VALVE", f'{oil_nb} NB', 1),
                    _row(media, "PRESSURE SWITCH LOW", '', 1),
                ]
        elif auto_control_type in ("plc_agr", "pid"):
            # PLC+AGR / PID: gas → AGR only, oil → AOR only
            if fuel_type in GAS_FUELS:
                rows.append(_row(
                    media, "AGR",
                    f'{equipment["agr"]["nb"]} NB',
                    1, unit_price_override=equipment["agr"]["price"],
                ))
            elif fuel_type in OIL_FUELS:
                rows.append(_row(media, "AIR OIL REGULATOR", f'{oil_nb} NB', 1))

    # AGR for non-PLC+AGR/PID modes (gas fuels)
    if fuel_type in GAS_FUELS and not (control_mode == "automatic" and auto_control_type in ("plc_agr", "pid")):
        rows.append(_row(
            media, "AGR",
            f'{equipment["agr"]["nb"]} NB',
            1, unit_price_override=equipment["agr"]["price"],
        ))

    return rows


def build_vlph_120t_df(
    equipment: dict,
    ladle_tons: float = 10.0,
    fuel1_type: str = "ng",
    fuel2_type: str = "none",
    equipment2: dict = None,
    control_mode: str = "automatic",
    auto_control_type: str = "agr",
) -> pd.DataFrame:
    """
    Builds VLPH BOM DataFrame.
    For dual fuel, equipment2 contains the second fuel's gas line equipment.
    """

    f1_label = FUEL_NAMES.get(fuel1_type, fuel1_type.upper())
    f2_label = FUEL_NAMES.get(fuel2_type, fuel2_type.upper()) if fuel2_type != "none" else None
    is_dual = fuel2_type != "none" and equipment2 is not None

    rows = []

    # Get ladle params (for ceramic rolls count, etc.)
    params = get_vlph_params(ladle_tons)

    # ── COMBUSTION AIR LINE ─────────────────────────────────────────────────
    is_plc = control_mode == "automatic" and auto_control_type == "plc"
    is_plc_agr = control_mode == "automatic" and auto_control_type == "plc_agr"
    is_pid = control_mode == "automatic" and auto_control_type == "pid"

    rows += [
        _row("COMB AIR", "COMPENSATOR", "", 1),
        _row("COMB AIR", "PRESSURE GAUGE WITH TNV", '0-2000 mm WC, Dial 4"', 1),
        _row("COMB AIR", "PRESSURE SWITCH LOW", '0-150 mBAR', 1),
    ]
    # PLC: air gets orifice plate + DPT + control valve
    if is_plc:
        rows += [
            _row("COMB AIR", "ORIFICE PLATE (Air)", f'{equipment["air_duct"]["nb"]} NB', 1),
            _row("COMB AIR", "DPT (Air)", "Output 4-20 mA", 1),
        ]
    # PLC, PLC+AGR, PID: air gets control valve
    if is_plc or is_plc_agr or is_pid:
        rows.append(_row(
            "COMB AIR", "PNEUMATIC CONTROL VALVE",
            f'{equipment["motorized_control_valve"]["nb"]} NB, '
            f'FLOW - {equipment["motorized_control_valve"]["flow_nm3hr"]} Nm3/hr',
            1,
            unit_price_override=equipment["motorized_control_valve"]["price"],
        ))
    rows += [
        _row(
            "COMB AIR", "BUTTERFLY VALVE",
            f'{equipment["butterfly_valve"]["nb"]} NB',
            1,
            unit_price_override=equipment["butterfly_valve"]["price"],
        ),
        _row(
            "COMB AIR", "ROTARY JOINT",
            f'{equipment["rotary_joint"]["nb"]} NB',
            1,
            unit_price_override=equipment["rotary_joint"]["price"],
        ),
        _row("COMB AIR", "BALL VALVE (Pilot Burner)", "20 NB", 1),
        _row("COMB AIR", "BALL VALVE (UV LINE)", "15 NB", 1),
        _row("COMB AIR", "FLEXIBLE HOSE (Pilot Burner)", "20 NB, 1500 mm", 1),
        _row("COMB AIR", "FLEXIBLE HOSE (UV LINE)", "15 NB, 1500 mm", 1),
    ]

    # ── FUEL 1 LINE ────────────────────────────────────────────────────────
    rows += _fuel_line_rows(f1_label, fuel1_type, equipment, control_mode, auto_control_type)

    # ── FUEL 2 LINE (dual fuel only) ──────────────────────────────────────
    if is_dual:
        rows += _fuel_line_rows(f2_label, fuel2_type, equipment2, control_mode, auto_control_type)

    # ── NG PILOT LINE ──────────────────────────────────────────────────────
    rows += [
        _row("NG PILOT LINE", "BALL VALVE", "20 NB", 1),
        _row("NG PILOT LINE", "PRESSURE GAUGE WITH NV", '0-1600 mm WC, Dial 4"', 1),
        _row("NG PILOT LINE", "BALL VALVE", "15 NB", 1),
        _row("NG PILOT LINE", "SOLENOID VALVE", "15 NB", 1),
        _row("NG PILOT LINE", "PRESSURE REGULATING VALVE",
             f'{equipment["agr"]["nb"]} NB', 1),
        _row("NG PILOT LINE", "FLEXIBLE HOSE PIPE", "15 NB - 1500 mm LONG", 1),
    ]


    # ── ENCON ITEMS ────────────────────────────────────────────────────────
    burner_desc = "ENCON DUAL FUEL Burner" if is_dual else equipment["burner"]["model"]
    rows += [
        _row(
            "ENCON ITEMS", burner_desc,
            f'GAS FLOW: {equipment["burner"]["input_nm3hr"]} Nm3/hr',
            1,
            unit_price_override=equipment["burner"]["price"],
        ),
        _row("ENCON ITEMS", "BEARING (24026)", "", 2),
        _row("ENCON ITEMS", "PLUMMER BLOCK",
             f'{params.get("plummer_block_kg", 300)} kg @ Rs.170/kg',
             1, unit_price_override=params.get("plummer_block_kg", 300) * 170),
        _row("ENCON ITEMS", "SHAFT",
             f'{params.get("shaft_kg", 350)} kg @ Rs.120/kg',
             1, unit_price_override=params.get("shaft_kg", 350) * 120),
        _row("ENCON ITEMS", "FABRICATION/ STRUCTURE",
             f'{params.get("fabrication_kg", 1900)} kg @ Rs.110/kg',
             1, unit_price_override=params.get("fabrication_kg", 1900) * 110),
        _row("ENCON ITEMS", "AIR-GAS PIPELINE",
             f'{params.get("pipeline_kg", 1000)} kg @ Rs.125/kg',
             1, unit_price_override=params.get("pipeline_kg", 1000) * 125),
        _row(
            "ENCON ITEMS", "COMBUSTION AIR BLOWER",
            f'{equipment["blower"]["hp"]} HP, '
            f'{equipment["blower"]["pressure"]} WC, '
            f'{equipment["blower"]["airflow_nm3hr"]} Nm3/hr — {equipment["blower"]["model"]}',
            1,
            unit_price_override=equipment["blower"]["price_premium"],
        ),
        _row("ENCON ITEMS", "IGNITION TRANSFORMER", "", 1),
        _row("ENCON ITEMS", "SEQUENCE CONTROLLER", "", 1),
        _row("ENCON ITEMS", "UV SENSOR WITH AIR JACKET", "", 1),
        _row("ENCON ITEMS", "PILOT BURNER", "", 1),
        _row("ENCON ITEMS", "CERAMIC FIBRE",
             f'{params["ceramic_rolls"]} Rolls',
             params["ceramic_rolls"]),
    ]

    # ── MISC ITEMS ─────────────────────────────────────────────────────────
    STATIC_SKIP = {"CONTROL PANEL"}  # CONTROL PANEL added separately below
    # PLC and PLC+AGR replace P.PID and RATIO CONTROLLER
    if is_plc or is_plc_agr:
        STATIC_SKIP.update({"P.PID", "RATIO CONTROLLER"})
    # PID: no temperature transmitter
    if is_pid:
        STATIC_SKIP.add("TEMPERATURE TRANSMITTER")
    for media, item, ref, qty in static_items():
        if item not in STATIC_SKIP:
            rows.append(_row(media, item, ref, qty))

    rows.append(_row("MISC ITEMS", "CONTROL PANEL", "1 Set", 1))
    rows.append(_row("MISC ITEMS", "INSTRUMENTS BALL VALVE", "15 NB", 3))
    # PLC WITH HMI only for PLC and PLC+AGR
    if is_plc or is_plc_agr:
        rows.append(_row("MISC ITEMS", "PLC WITH HMI", "", 1))

    df = pd.DataFrame(
        rows,
        columns=["MEDIA", "ITEM NAME", "REFERENCE", "QTY", "UNIT PRICE", "TOTAL"],
    )

    # Summary rows (SAIL style: Bought Out + In-house/ENCON)
    bought_out_total = df.loc[
        (df["MEDIA"] != "ENCON ITEMS")
        & (~df["ITEM NAME"].isin(BOUGHT_OUT_EXCLUDE_ITEMS)),
        "TOTAL",
    ].sum()

    encon_total = df.loc[df["MEDIA"] == "ENCON ITEMS", "TOTAL"].sum()

    grand_total = bought_out_total + encon_total

    df = pd.concat(
        [
            df,
            pd.DataFrame(
                [
                    ("", "BOUGHT OUT ITEMS",     "", "", "", bought_out_total),
                    ("", "ENCON ITEMS",          "", "", "", encon_total),
                    ("", "GRAND TOTAL",          "", "", "", grand_total),
                ],
                columns=df.columns,
            ),
        ],
        ignore_index=True,
    )

    return df
