# bom/hlph_builder.py
"""
BOM builder for Horizontal Ladle Pre-Heater (HLPH).

Differences from VLPH:
  - No ROTARY JOINT (horizontal — no swirling mechanism)
  - Has TROLLEY DRIVE MECHANISM instead
"""

import pandas as pd

from bom.static_items import static_items
from bom.price_master import get_price


BOUGHT_OUT_EXCLUDE_ITEMS = {
    "RATIO CONTROLLER",
}

LEGACY_ITEM_SEQUENCE = [
    "COMPENSATOR",
    "PRESSURE GAUGE WITH TNV",
    "PRESSURE SWITCH LOW",
    "MOTORIZED CONTROL VALVE",
    "BUTTERFLY VALVE",
    "BALL VALVE (Pilot Burner)",
    "BALL VALVE (UV LINE)",
    "FLEXIBLE HOSE (Pilot Burner)",
    "FLEXIBLE HOSE (UV LINE)",
    "BALL VALVE",
    "PRESSURE GAUGE WITH NV",
    "PRESSURE SWITCH HIGH + LOW",
    "SOLENOID VALVE",
    "PRESSURE REGULATING VALVE",
    "FLEXIBLE HOSE PIPE",
    "AGR",
    "THERMOCOUPLE",
    "COMPENSATING LEAD",
    "LIMIT SWITCHES",
    "CONTROL PANEL",
    "HYDRAULIC POWER PACK & CYLINDER",
    "CABLE FOR IGNITION TRANSFORMER",
    "TEMPERATURE TRANSMITTER",
    "P.PID",
    "RATIO CONTROLLER",
]


def _row(media: str, item: str, ref: str, qty: int, unit_price_override=None):
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


def build_hlph_df(equipment: dict) -> pd.DataFrame:
    """
    Builds HLPH BOM DataFrame from already-selected equipment dict.
    equipment must come from bom.selectors.selection_engine.select_equipment()
    """

    rows = []

    # COMBUSTION AIR LINE — same as VLPH but no rotary joint
    rows += [
        _row("COMB AIR", "COMPENSATOR", f'{equipment["air_duct"]["nb"]} NB F150#', 1),
        _row("COMB AIR", "PRESSURE GAUGE WITH TNV", '0-2000 mm WC, Dial 4"', 1),
        _row("COMB AIR", "PRESSURE SWITCH LOW", '0-150 mBAR', 1),
        _row(
            "COMB AIR", "MOTORIZED CONTROL VALVE",
            f'{equipment["motorized_control_valve"]["nb"]} NB, '
            f'FLOW - {equipment["motorized_control_valve"]["flow_nm3hr"]} Nm3/hr',
            1,
            unit_price_override=equipment["motorized_control_valve"]["price"],
        ),
        _row(
            "COMB AIR", "BUTTERFLY VALVE",
            f'{equipment["butterfly_valve"]["nb"]} NB',
            1,
            unit_price_override=equipment["butterfly_valve"]["price"],
        ),
        # No ROTARY JOINT for HLPH
    ]

    # NG PILOT LINE
    rows += [
        _row("NG PILOT LINE", "BALL VALVE", "20 NB", 2),
        _row("NG PILOT LINE", "BALL VALVE (Pilot Burner)", "20 NB", 1),
        _row("NG PILOT LINE", "BALL VALVE (UV LINE)", "15 NB", 1),
        _row("NG PILOT LINE", "PRESSURE GAUGE WITH NV", '0-1600 mm WC, Dial 4"', 1),
        _row("NG PILOT LINE", "PRESSURE SWITCH HIGH + LOW", "", 2),
        _row("NG PILOT LINE", "SOLENOID VALVE", "15 NB", 1),
        _row(
            "NG PILOT LINE", "PRESSURE REGULATING VALVE",
            f'{equipment["agr"]["nb"]} NB',
            1,
        ),
        _row("NG PILOT LINE", "FLEXIBLE HOSE (Pilot Burner)", "20 NB, 1500 mm", 1),
        _row("NG PILOT LINE", "FLEXIBLE HOSE (UV LINE)", "15 NB, 1500 mm", 1),
        _row("NG PILOT LINE", "FLEXIBLE HOSE PIPE", "15 NB - 1500 mm LONG", 1),
    ]

    # MG LINE
    gas_train_name = f'GAS TRAIN {equipment["ng_gas_train"]["max_flow"]:.0f} NM3/Hr'
    rows += [
        _row(
            "MG LINE", gas_train_name,
            f'{equipment["ng_gas_train"]["inlet_nb"]} x '
            f'{equipment["ng_gas_train"]["outlet_nb"]} NB',
            1,
            unit_price_override=equipment["ng_gas_train"]["price"],
        ),
        _row(
            "MG LINE", "AGR",
            f'{equipment["agr"]["nb"]} NB',
            1,
            unit_price_override=equipment["agr"]["price"],
        ),
    ]

    # ENCON ITEMS
    rows += [
        _row(
            "ENCON ITEMS", equipment["burner"]["model"],
            f'NATURAL GAS FLOW: {equipment["burner"]["input_nm3hr"]} Nm3/hr',
            1,
            unit_price_override=equipment["burner"]["price"],
        ),
        _row(
            "ENCON ITEMS", equipment["blower"]["model"],
            f'{equipment["blower"]["hp"]} HP, '
            f'{equipment["blower"]["pressure"]} WC, '
            f'{equipment["blower"]["airflow_nm3hr"]} Nm3/hr',
            1,
            unit_price_override=equipment["blower"]["price_basic"],
        ),
        _row("ENCON ITEMS", "ENCON-PB (NG/LPG) - 100 KW", "", 1),
        _row("ENCON ITEMS", "Ignition Transformer", "", 1),
        _row("ENCON ITEMS", "Burner Control Unit", "", 1),
        _row("ENCON ITEMS", "UV Sensor with Air Jacket", "", 1),
    ]

    # STATIC ITEMS
    for media, item, ref, qty in static_items():
        rows.append(_row(media, item, ref, qty))

    df = pd.DataFrame(
        rows,
        columns=["MEDIA", "ITEM NAME", "REFERENCE", "QTY", "UNIT PRICE", "TOTAL"],
    )

    # Sort by legacy sequence
    order_map = {name: i for i, name in enumerate(LEGACY_ITEM_SEQUENCE)}
    df["_order"] = df["ITEM NAME"].map(order_map).fillna(999)
    df = df.sort_values("_order").drop(columns="_order").reset_index(drop=True)

    # Summary rows
    bought_out_total = df.loc[
        (df["MEDIA"] != "ENCON ITEMS") & (~df["ITEM NAME"].isin(BOUGHT_OUT_EXCLUDE_ITEMS)),
        "TOTAL",
    ].sum()

    encon_total = df.loc[df["MEDIA"] == "ENCON ITEMS", "TOTAL"].sum()

    df = pd.concat(
        [
            df,
            pd.DataFrame(
                [
                    ("", "BOUGHT OUT ITEMS", "", "", "", bought_out_total),
                    ("", "ENCON ITEMS", "", "", "", encon_total),
                ],
                columns=df.columns,
            ),
        ],
        ignore_index=True,
    )

    return df
