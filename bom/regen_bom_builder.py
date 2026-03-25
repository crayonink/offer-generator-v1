"""
bom/regen_bom_builder.py

Builds the BOM DataFrame for a Regenerative Burner system.
Corrected for proper scaling (per burner, multipliers, system-level).
"""

import pandas as pd
from bom.price_master import PRICE_MASTER


# -------------------------------------------------
# REGEN-SPECIFIC PRICE OVERRIDES
# -------------------------------------------------

REGEN_PRICE_OVERRIDES = {
    "REGEN BURNER BODY (MS FABRICATED)": 0,
    "REGENERATOR (MS + SS FABRICATED)": 0,
    "BURNER BLOCK (CASTABLE REFRACTORY)": 0,
    "CERAMIC BALLS (HEAT STORAGE MEDIA)": 0,
    "REGEN REFRACTORY LINING": 0,
    "REVERSING VALVE (4-WAY)": 0,
    "REVERSING VALVE ACTUATOR": 0,
    "SEQUENCER / PLC FOR REVERSING": 0,
    "HOT GAS BY-PASS VALVE": 0,
    "REGEN BLOWER": 0,
    "REGEN PILOT BURNER": 0,
    "IGNITION TRANSFORMER (REGEN)": 0,
    "UV SENSOR WITH AIR JACKET (REGEN)": 0,
    "SEQUENCE CONTROLLER (REGEN)": 0,
}


# -------------------------------------------------
# ITEM ORDER
# -------------------------------------------------

REGEN_ITEM_SEQUENCE = [
    "COMPENSATOR",
    "PRESSURE GAUGE WITH TNV",
    "PRESSURE SWITCH LOW",
    "MOTORIZED CONTROL VALVE",
    "BUTTERFLY VALVE",

    "BALL VALVE",
    "PRESSURE GAUGE WITH NV",
    "PRESSURE SWITCH HIGH + LOW",
    "SOLENOID VALVE",
    "PRESSURE REGULATING VALVE",
    "FLEXIBLE HOSE PIPE",
    "NG GAS TRAIN",
    "AGR",

    "REVERSING VALVE (4-WAY)",
    "REVERSING VALVE ACTUATOR",
    "HOT GAS BY-PASS VALVE",
    "SEQUENCER / PLC FOR REVERSING",

    "THERMOCOUPLE",
    "COMPENSATING LEAD",
    "LIMIT SWITCHES",
    "TEMPERATURE TRANSMITTER",
    "CONTROL PANEL",

    "REGEN BURNER BODY (MS FABRICATED)",
    "REGENERATOR (MS + SS FABRICATED)",
    "BURNER BLOCK (CASTABLE REFRACTORY)",
    "CERAMIC BALLS (HEAT STORAGE MEDIA)",
    "REGEN REFRACTORY LINING",
    "REGEN BLOWER",
    "REGEN PILOT BURNER",
    "IGNITION TRANSFORMER (REGEN)",
    "UV SENSOR WITH AIR JACKET (REGEN)",
    "SEQUENCE CONTROLLER (REGEN)",
]


# -------------------------------------------------
# HELPER
# -------------------------------------------------

def _row(media, item, ref, qty, unit_price_override=None):
    if unit_price_override is not None:
        unit_price = unit_price_override
    else:
        unit_price = REGEN_PRICE_OVERRIDES.get(item) or PRICE_MASTER.get(item, 0)

    total = unit_price * qty
    return (media, item, ref, qty, unit_price, total)


# -------------------------------------------------
# MAIN BUILDER
# -------------------------------------------------

def build_regen_bom_df(regen_results) -> pd.DataFrame:

    r = regen_results
    s = r.sizing
    p = r.pipe_sizing

    n = r.num_burners
    kw = r.power_kw

    air_dn = p.air_in_dn if p else "—"
    fume_dn = p.fume_out_dn if p else "—"
    ng_dn = p.ng_in_dn if p else "—"

    rows = []

    # -------------------------------------------------
    # COMBUSTION AIR LINE
    # -------------------------------------------------

    rows += [
        _row("COMB AIR", "COMPENSATOR", f"{air_dn} NB F150#", n),
        _row("COMB AIR", "PRESSURE GAUGE WITH TNV", '0–2000 mm WC, Dial 4"', n),
        _row("COMB AIR", "PRESSURE SWITCH LOW", "0–150 mBAR", n),
        _row("COMB AIR", "MOTORIZED CONTROL VALVE", f"{air_dn} NB, Regen Air Flow", 1),  # system
        _row("COMB AIR", "BUTTERFLY VALVE", f"{air_dn} NB", n),
    ]

    # -------------------------------------------------
    # NG LINE
    # -------------------------------------------------

    rows += [
        _row("NG LINE", "BALL VALVE", f"{ng_dn} NB", 5 * n),  # 🔴 corrected
        _row("NG LINE", "PRESSURE GAUGE WITH NV", '0–1600 mm WC, Dial 4"', n),
        _row("NG LINE", "PRESSURE SWITCH HIGH + LOW", "", 2 * n),
        _row("NG LINE", "SOLENOID VALVE", f"{ng_dn} NB", n),
        _row("NG LINE", "PRESSURE REGULATING VALVE", f"{ng_dn} NB", n),
        _row("NG LINE", "FLEXIBLE HOSE PIPE", f"{ng_dn} NB – 1500 mm LONG", 5 * n),  # 🔴 corrected
        _row("NG LINE", "NG GAS TRAIN", f"{ng_dn} NB | {r.fuel_type}", 1),  # 🔴 system
        _row("NG LINE", "AGR", f"{ng_dn} NB", n),
    ]

    # -------------------------------------------------
    # REVERSING SYSTEM
    # -------------------------------------------------

    rows += [
        _row("REVERSING SYSTEM", "REVERSING VALVE (4-WAY)", f"{air_dn} NB / {fume_dn} NB", n),
        _row("REVERSING SYSTEM", "REVERSING VALVE ACTUATOR", "Pneumatic", n),
        _row("REVERSING SYSTEM", "HOT GAS BY-PASS VALVE", f"{fume_dn} NB", n),
        _row("REVERSING SYSTEM", "SEQUENCER / PLC FOR REVERSING", "Cycle time: 30–120 sec", 1),
    ]

    # -------------------------------------------------
    # INSTRUMENTATION
    # -------------------------------------------------

    rows += [
        _row("INSTRUMENTATION", "THERMOCOUPLE", "K-TYPE", 2 * n),
        _row("INSTRUMENTATION", "COMPENSATING LEAD", "K-TYPE, 10m", 2 * n),
        _row("INSTRUMENTATION", "LIMIT SWITCHES", "", 2 * n),
        _row("INSTRUMENTATION", "TEMPERATURE TRANSMITTER", "4–20 mA", n),
        _row("INSTRUMENTATION", "CONTROL PANEL", f"For {n} Regen Burner Pair(s)", 1),
    ]

    # -------------------------------------------------
    # ENCON ITEMS
    # -------------------------------------------------

    rows += [
        _row("ENCON ITEMS", "REGEN BURNER BODY (MS FABRICATED)",
             f"{kw} kW | L={s['burner_length']}m Ø{s['burner_dia']}m",
             n),

        _row("ENCON ITEMS", "REGENERATOR (MS + SS FABRICATED)",
             f"{kw} kW | {s['regen_L']}×{s['regen_H']}×{s['regen_W']} m",
             n),

        _row("ENCON ITEMS", "BURNER BLOCK (CASTABLE REFRACTORY)",
             f"Ø{s['burner_block_inner_dia']}m inner / Ø{s['burner_block_outer_dia']}m outer",
             n),

        _row("ENCON ITEMS", "CERAMIC BALLS (HEAT STORAGE MEDIA)",
             f"{r.weights.ceramic_balls_kg:.1f} kg/burner",
             n),

        _row("ENCON ITEMS", "REGEN REFRACTORY LINING",
             f"Castable, thk={s['refractory_thk_m']*1000:.0f} mm",
             n),

        _row("ENCON ITEMS", "REGEN BLOWER",
             f"For {kw} kW Regen Burner",
             1),  # 🔴 system

        _row("ENCON ITEMS", "REGEN PILOT BURNER", "10 kW", n),
        _row("ENCON ITEMS", "IGNITION TRANSFORMER (REGEN)", "", n),
        _row("ENCON ITEMS", "UV SENSOR WITH AIR JACKET (REGEN)", "", n),
        _row("ENCON ITEMS", "SEQUENCE CONTROLLER (REGEN)", "", 1),
    ]

    # -------------------------------------------------
    # BUILD DATAFRAME
    # -------------------------------------------------

    df = pd.DataFrame(
        rows,
        columns=["MEDIA", "ITEM NAME", "REFERENCE", "QTY", "UNIT PRICE", "TOTAL"],
    )

    order_map = {name: i for i, name in enumerate(REGEN_ITEM_SEQUENCE)}
    df["_order"] = df["ITEM NAME"].map(order_map).fillna(999)

    df = df.sort_values("_order").drop(columns="_order").reset_index(drop=True)

    bought_out_total = df.loc[df["MEDIA"] != "ENCON ITEMS", "TOTAL"].sum()
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