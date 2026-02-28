import pandas as pd

from bom.static_items import static_items
from bom.price_master import PRICE_MASTER
from bom.selectors.selection_engine import select_equipment


# -------------------------------------------------
# LEGACY EXCLUSION RULES
# -------------------------------------------------
BOUGHT_OUT_EXCLUDE_ITEMS = {
    "RATIO CONTROLLER",
}


# -------------------------------------------------
# LEGACY ITEM SEQUENCE (EXACT MATCH)
# -------------------------------------------------
LEGACY_ITEM_SEQUENCE = [
    "COMPENSATOR",
    "PRESSURE GAUGE WITH TNV",
    "PRESSURE SWITCH LOW",
    "MOTORIZED CONTROL VALVE",
    "BUTTERFLY VALVE",
    "ROTARY JOINT",
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
    "NG GAS TRAIN",
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


# -------------------------------------------------
# HELPER
# -------------------------------------------------
def _row(media: str, item: str, ref: str, qty: int, unit_price_override=None):
    if unit_price_override is not None:
        unit_price = unit_price_override
    else:
        unit_price = PRICE_MASTER.get(item, 0)

    total = unit_price * qty

    return (media, item, ref, qty, unit_price, total)


# -------------------------------------------------
# MAIN BUILDER
# -------------------------------------------------
def build_vlph_120t_df(*, burner_results, pipe_results) -> pd.DataFrame:

    equipment = select_equipment(
        ng_flow_nm3hr=burner_results.extra_firing_rate_nm3hr,
        air_flow_nm3hr=burner_results.air_qty_nm3hr,
    )

    rows = []

    # -------------------------------------------------
    # COMBUSTION AIR LINE
    # -------------------------------------------------
    rows += [

        _row(
            "COMB AIR",
            "COMPENSATOR",
            f'{equipment["compensator"]["nb"]} NB F150#',
            1,
            unit_price_override=equipment["compensator"]["price"],
        ),

        _row(
            "COMB AIR",
            "PRESSURE GAUGE WITH TNV",
            '0–2000 mm WC, Dial 4"',
            1,
        ),

        _row(
            "COMB AIR",
            "PRESSURE SWITCH LOW",
            '0–150 mBAR',
            1,
        ),

        _row(
            "COMB AIR",
            "MOTORIZED CONTROL VALVE",
            f'{equipment["motorized_control_valve"]["nb"]} NB, '
            f'FLOW – {equipment["motorized_control_valve"]["flow_nm3hr"]} Nm3/hr',
            1,
            unit_price_override=equipment["motorized_control_valve"]["price"],
        ),

        _row(
            "COMB AIR",
            "BUTTERFLY VALVE",
            f'{equipment["butterfly_valve"]["nb"]} NB',
            1,
            unit_price_override=equipment["butterfly_valve"]["price"],
        ),

        _row(
            "COMB AIR",
            "ROTARY JOINT",
            f'{equipment["rotary_joint"]["nb"]} NB',
            1,
            unit_price_override=equipment["rotary_joint"]["price"],
        ),
    ]

    # -------------------------------------------------
    # NG PILOT LINE
    # -------------------------------------------------
    rows += [
        _row("NG PILOT LINE", "BALL VALVE", "20 NB", 2),
        _row("NG PILOT LINE", "BALL VALVE (Pilot Burner)", "20 NB", 1),
        _row("NG PILOT LINE", "BALL VALVE (UV LINE)", "15 NB", 1),

        _row(
            "NG PILOT LINE",
            "PRESSURE GAUGE WITH NV",
            '0–1600 mm WC, Dial 4"',
            1,
        ),

        _row("NG PILOT LINE", "PRESSURE SWITCH HIGH + LOW", "", 2),
        _row("NG PILOT LINE", "SOLENOID VALVE", "15 NB", 1),

        _row(
            "NG PILOT LINE",
            "PRESSURE REGULATING VALVE",
            f'{equipment["agr"]["nb"]} NB',
            1,
        ),

        _row(
            "NG PILOT LINE",
            "FLEXIBLE HOSE (Pilot Burner)",
            "20 NB, 1500 mm",
            1,
        ),

        _row(
            "NG PILOT LINE",
            "FLEXIBLE HOSE (UV LINE)",
            "15 NB, 1500 mm",
            1,
        ),

        _row(
            "NG PILOT LINE",
            "FLEXIBLE HOSE PIPE",
            "15 NB - 1500 mm LONG",
            1,
        ),
    ]

    # -------------------------------------------------
    # MG LINE
    # -------------------------------------------------
    rows += [

        _row(
            "MG LINE",
            "NG GAS TRAIN",
            f'FLOW: {equipment["ng_gas_train"]["flow_nm3hr"]} Nm3/hr',
            1,
            unit_price_override=equipment["ng_gas_train"]["price"],
        ),

        _row(
            "MG LINE",
            "AGR",
            f'{equipment["agr"]["nb"]} NB',
            1,
            unit_price_override=equipment["agr"]["price"],
        ),
    ]

    # -------------------------------------------------
    # ENCON ITEMS
    # -------------------------------------------------
    rows += [

        _row(
            "ENCON ITEMS",
            "ENCON MG BURNER WITH B. BLOCK",
            f'NATURAL GAS FLOW: '
            f'{equipment["encon_burner"]["max_flow_nm3hr"]} Nm3/hr '
            f'{equipment["encon_burner"]["model"]}',
            1,
            unit_price_override=equipment["encon_burner"]["price"],
        ),

        _row(
            "ENCON ITEMS",
            "COMBUSTION AIR BLOWER",
            f'{equipment["blower"]["hp"]} HP, '
            f'{equipment["blower"]["pressure_mm_wc"]}" WC, '
            f'{equipment["blower"]["flow_nm3hr"]} Nm3/hr',
            1,
            unit_price_override=equipment["blower"]["price"],
        ),

        _row("ENCON ITEMS", "PILOT BURNER", "10 KW", 1),
        _row("ENCON ITEMS", "IGNITION TRANSFORMER", "", 1),
        _row("ENCON ITEMS", "SEQUENCE CONTROLLER", "", 1),
        _row("ENCON ITEMS", "UV SENSOR WITH AIR JACKET", "", 1),
    ]

    # -------------------------------------------------
    # STATIC ITEMS
    # -------------------------------------------------
    for media, item, ref, qty in static_items():
        rows.append(_row(media, item, ref, qty))

    # -------------------------------------------------
    # DATAFRAME
    # -------------------------------------------------
    df = pd.DataFrame(
        rows,
        columns=["MEDIA", "ITEM NAME", "REFERENCE", "QTY", "UNIT PRICE", "TOTAL"],
    )

    # -------------------------------------------------
    # LEGACY ORDER
    # -------------------------------------------------
    order_map = {name: i for i, name in enumerate(LEGACY_ITEM_SEQUENCE)}
    df["_order"] = df["ITEM NAME"].map(order_map).fillna(999)
    df = df.sort_values("_order").drop(columns="_order").reset_index(drop=True)

    # -------------------------------------------------
    # TOTALS
    # -------------------------------------------------
    bought_out_total = df.loc[
        (df["MEDIA"] != "ENCON ITEMS")
        & (~df["ITEM NAME"].isin(BOUGHT_OUT_EXCLUDE_ITEMS)),
        "TOTAL",
    ].sum()

    encon_total = df.loc[
        df["MEDIA"] == "ENCON ITEMS",
        "TOTAL",
    ].sum()

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