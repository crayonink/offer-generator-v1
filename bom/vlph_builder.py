# bom/vlph_builder.py
import pandas as pd

from bom.vlph_rules import vlph_rules
from bom.static_items import static_items
from bom.price_master import PRICE_MASTER


# -------------------------------------------------
# LEGACY EXCLUSION RULES
# -------------------------------------------------
BOUGHT_OUT_EXCLUDE_ITEMS = {
    "RATIO CONTROLLER",
    
}


# -------------------------------------------------
# LEGACY ITEM SEQUENCE (DISPLAY + SUM RANGE MATCH)
# -------------------------------------------------
LEGACY_ITEM_SEQUENCE = [
    
"COMPENSATOR",
"PRESSURE GAUGE WITH TNV",
"PRESSURE SWITCH LOW (Set Pt -L)",
"MOTERIZED CONTROL VALVE",
"BUTTERFLY VALVE",
"ROTARY JOINT",
"BALL VALVE (Pilot Burner)",
"BALL VALVE (UV - LINE)",
"FLEXIBLE HOSE (Pilot Burner)",
"FLEXIBLE HOSE (UV - LINE)",
"BALL VALVE",
"PRESSURE GAUGE WITH NV",
"PRESSURE SWITCH HIGH + LOW",
"SOLENOID VALVE",
"PRESSURE REGULATING VALVE",
"FLEXIBALE HOSE PIPE",
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
"RATIO CONTROLLER"

]


def _row(media: str, item: str, ref: str, qty: int):
    unit_price = PRICE_MASTER.get(item, 0)
    return (media, item, ref, qty, unit_price * qty)


def build_vlph_120t_df(*, burner_results, pipe_results) -> pd.DataFrame:
    rules = vlph_rules(burner_results.extra_firing_rate_nm3hr)
    rows = []

    # -------------------------------------------------
    # COMBUSTION AIR LINE
    # -------------------------------------------------
    rows += [
        _row("COMB AIR", "COMPENSATOR", f'{rules["air_pipe_nb"]} NB F150#', 1),
        _row("COMB AIR", "PRESSURE GAUGE WITH TNV", '0–2000 mm WC, Dial 4"', 1),
        _row("COMB AIR", "PRESSURE SWITCH LOW", '0–150 mBAR', 1),
        _row(
            "COMB AIR",
            "MOTORIZED CONTROL VALVE",
            f'250 NB, FLOW – {round(burner_results.air_qty_nm3hr)} Nm3/hr',
            1,
        ),
        _row("COMB AIR", "BUTTERFLY VALVE", f'{rules["air_pipe_nb"]} NB', 1),
        _row("COMB AIR", "ROTARY JOINT", f'{rules["air_pipe_nb"]} NB', 1),
    ]

    # -------------------------------------------------
    # NG PILOT LINE
    # -------------------------------------------------
    rows += [
        _row("NG PILOT LINE", "BALL VALVE", "20 NB", 2),  # ✅ critical
        _row("NG PILOT LINE", "BALL VALVE (Pilot Burner)", "20 NB", 1),
        _row("NG PILOT LINE", "BALL VALVE (UV LINE)", "15 NB", 1),
        _row("NG PILOT LINE", "PRESSURE GAUGE WITH NV", '0–1600 mm WC, Dial 4"', 1),
        _row("NG PILOT LINE", "PRESSURE SWITCH HIGH + LOW", "", 2),
        _row("NG PILOT LINE", "SOLENOID VALVE", "15 NB", 1),
        _row("NG PILOT LINE", "PRESSURE REGULATING VALVE", "15 NB", 1),
        _row("NG PILOT LINE", "FLEXIBLE HOSE (Pilot Burner)", "20 NB, 1500 mm", 1),
        _row("NG PILOT LINE", "FLEXIBLE HOSE (UV LINE)", "15 NB, 1500 mm", 1),
        _row("NG PILOT LINE", "FLEXIBLE HOSE PIPE", "15 NB - 1500 mm LONG", 1),
    ]

    # -------------------------------------------------
    # MG LINE
    # -------------------------------------------------
    rows += [
        _row("MG LINE", "NG GAS TRAIN", "FLOW: 400 Nm3/hr", 1),
        _row("MG LINE", "AGR", "80 NB", 1),
    ]

    # -------------------------------------------------
    # ENCON ITEMS
    # -------------------------------------------------
    rows += [
        _row(
            "ENCON ITEMS",
            "ENCON MG BURNER WITH B. BLOCK",
            "NATURAL GAS FLOW: 440 Nm3/hr G7A",
            1,
        ),
        _row(
            "ENCON ITEMS",
            "COMBUSTION AIR BLOWER",
            '25 HP, 28" WC, 5100 Nm3/hr',
            1,
        ),
        _row("ENCON ITEMS", "PILOT BURNER", "10 KW", 1),
        _row("ENCON ITEMS", "IGNITION TRANSFORMER", "", 1),
        _row("ENCON ITEMS", "SEQUENCE CONTROLLER", "", 1),
        _row("ENCON ITEMS", "UV SENSOR WITH AIR JACKET", "", 1),
    ]

    # -------------------------------------------------
    # STATIC / MISC ITEMS
    # -------------------------------------------------
    for media, item, ref, qty in static_items():
        rows.append(_row(media, item, ref, qty))

    # -------------------------------------------------
    # DATAFRAME
    # -------------------------------------------------
    df = pd.DataFrame(
        rows,
        columns=["MEDIA", "ITEM NAME", "REFERENCE", "QTY", "TOTAL"],
    )

    # -------------------------------------------------
    # APPLY LEGACY ORDER
    # -------------------------------------------------
    order_map = {name: i for i, name in enumerate(LEGACY_ITEM_SEQUENCE)}
    df["_order"] = df["ITEM NAME"].map(order_map).fillna(999)
    df = df.sort_values("_order").drop(columns="_order").reset_index(drop=True)

    # -------------------------------------------------
    # TOTALS (ARCHITECTURE PRESERVED)
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
                    ("", "BOUGHT OUT ITEMS", "", "", bought_out_total),
                    ("", "ENCON ITEMS", "", "", encon_total),
                ],
                columns=df.columns,
            ),
        ],
        ignore_index=True,
    )

    return df
