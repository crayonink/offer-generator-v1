# bom/hlph_builder.py
"""
BOM builder for Horizontal Ladle Pre-Heater (HLPH).

Differences from VLPH:
  - No ROTARY JOINT (horizontal — no swirling mechanism)
  - Has TROLLEY DRIVE MECHANISM instead
  - Uses get_hlph_params for structure/ceramic/trolley sizing

Shares fuel line logic, vendor support, and base_only with VLPH.
"""

import pandas as pd

from bom.price_master import get_price, DB_PATH
from bom.ladle_params import get_hlph_params
from bom.vlph_builder import (
    _row, _fuel_line_rows, _get_cheapest_ball_valve,
    _get_cheapest_solenoid_valve, _get_flexible_hose_price,
    _get_price_fuzzy, FUEL_NAMES, OIL_FUELS, GAS_FUELS,
)
from bom.selectors.gas_regulator_selector import select_gas_regulator


def build_hlph_df(
    equipment: dict,
    ladle_tons: float = 10.0,
    fuel1_type: str = "ng",
    fuel2_type: str = "none",
    equipment2: dict = None,
    control_mode: str = "automatic",
    auto_control_type: str = "agr",
    control_valve_vendor: str = "dembla",
    butterfly_valve_vendor: str = "lt_lever",
    shutoff_valve_vendor: str = "aira",
    pressure_gauge_vendor: str = "baumer",
    pilot_burner: str = "auto",
    pilot_line_fuel: str = "lpg",
    pipeline_weight_kg: float = 1000.0,
    purging_line: str = "no",
) -> pd.DataFrame:
    """
    Builds HLPH BOM DataFrame (automatic mode).
    Same logic as VLPH but no rotary joint, has trolley drive.
    """

    f1_label = FUEL_NAMES.get(fuel1_type, fuel1_type.upper())
    f2_label = FUEL_NAMES.get(fuel2_type, fuel2_type.upper()) if fuel2_type != "none" else None
    is_dual = fuel2_type != "none" and equipment2 is not None

    rows = []
    params = get_hlph_params(ladle_tons)

    pg_vendor = pressure_gauge_vendor.upper()
    pg_item = f'PRESSURE GAUGE WITH TNV ({pg_vendor})'

    is_plc = control_mode == "automatic" and auto_control_type == "plc"
    is_plc_agr = control_mode == "automatic" and auto_control_type == "plc_agr"
    is_pid = control_mode == "automatic" and auto_control_type == "pid"

    # ── STRUCTURE & SYSTEM ────────────────────────────────────────────────
    rows += [
        _row("ENCON ITEMS", "FABRICATION/ STRUCTURE",
             f'{params["ms_structure_kg"]} kg',
             1, unit_price_override=params["ms_structure_kg"] * get_price("FABRICATION RATE")),
        _row("ENCON ITEMS", "AIR-GAS PIPELINE",
             f'{pipeline_weight_kg:.0f} kg', 1,
             unit_price_override=pipeline_weight_kg * get_price("PIPELINE RATE")),
        _row("ENCON ITEMS", "CERAMIC FIBRE",
             f'{params["ceramic_rolls"]} Rolls @ Rs.{params.get("ceramic_rate", 0):,.0f}/roll',
             params["ceramic_rolls"],
             unit_price_override=params.get("ceramic_rate", 0)),
    ]

    # ── TROLLEY MECHANISM (individual items from DB) ──────────────────────
    rows += [
        _row("MISC ITEMS", "GEARED MOTOR MECHANISM", "3 HP", 1, make="POWERTEK"),
        _row("MISC ITEMS", "TROLLEY WHEEL", "CastIron", 4, make="ENCON"),
        _row("MISC ITEMS", "PLUMMER BLOCK", "MS IS-2062", 6, make="ENCON"),
        _row("MISC ITEMS", "SHAFT (1 long, 2 Short)", "EN-8", 1, make="ENCON"),
        _row("MISC ITEMS", "BEARING", "NU2214", 4, make="ENCON"),
        _row("MISC ITEMS", "SLIEVE", "", 4, make="ENCON"),
        _row("MISC ITEMS", "FRAME", "MS Structure", 1000, make="ENCON"),
    ]

    # ── COMBUSTION AIR LINE ───────────────────────────────────────────────
    air_nb = max(125, equipment["air_duct"]["nb"])

    rows += [
        _row("COMB AIR", "COMPENSATOR", "", 1),
        _row("COMB AIR", pg_item, '', 1, make=pg_vendor),
        _row("COMB AIR", "PRESSURE SWITCH LOW", '', 1, make="MADAS"),
    ]
    if is_plc:
        from bom.vlph_builder import _get_orifice_price
        op_nb, op_price = _get_orifice_price(air_nb)
        rows += [
            _row("COMB AIR", "ORIFICE PLATE", f'{op_nb} NB', 1,
                 unit_price_override=op_price, make="ENCON"),
            _row("COMB AIR", "DPT", '', 1, make="HONEYWELL"),
        ]
    if is_plc or is_plc_agr or is_pid:
        from calculations.pipes import STANDARD_PIPE_NB
        from bom.vlph_builder import _get_valve_price
        try:
            cv_nb = STANDARD_PIPE_NB[max(0, STANDARD_PIPE_NB.index(air_nb) - 1)]
        except ValueError:
            cv_nb = air_nb
        _, cv_price = _get_valve_price(cv_nb, "control", control_valve_vendor)
        vendor_label = control_valve_vendor.upper()
        rows.append(_row(
            "COMB AIR", "CONTROL VALVE",
            f'{cv_nb} NB',
            1, unit_price_override=cv_price, make=vendor_label,
        ))
        bfv = equipment["butterfly_valve"]
        rows.append(_row(
            "COMB AIR", "BUTTERFLY VALVE",
            f'{bfv["nb"]} NB',
            1, unit_price_override=bfv["price"], make=bfv.get("make", "L&T"),
        ))
    # No ROTARY JOINT for HLPH
    rows += [
        _row("COMB AIR", "BALL VALVE (Pilot Burner)", "20 NB", 1,
             unit_price_override=_get_cheapest_ball_valve(20), make="L&T"),
        _row("COMB AIR", "BALL VALVE (UV LINE)", "15 NB", 1,
             unit_price_override=_get_cheapest_ball_valve(15), make="L&T"),
        _row("COMB AIR", "FLEXIBLE HOSE (Pilot Burner)",
             f'{_get_flexible_hose_price(20)[0]} NB, 1500mm', 1,
             unit_price_override=_get_flexible_hose_price(20)[1], make="BENGAL IND."),
        _row("COMB AIR", "FLEXIBLE HOSE (UV LINE)",
             f'{_get_flexible_hose_price(15)[0]} NB, 1500mm', 1,
             unit_price_override=_get_flexible_hose_price(15)[1], make="BENGAL IND."),
    ]

    # ── FUEL 1 LINE ──────────────────────────────────────────────────────
    rows += _fuel_line_rows(f1_label, fuel1_type, equipment, control_mode,
                            auto_control_type, control_valve_vendor,
                            pressure_gauge_vendor, butterfly_valve_vendor,
                            shutoff_valve_vendor)

    # ── FUEL 2 LINE (dual fuel only) ─────────────────────────────────────
    if is_dual:
        rows += _fuel_line_rows(f2_label, fuel2_type, equipment2, control_mode,
                                auto_control_type, control_valve_vendor,
                                pressure_gauge_vendor, butterfly_valve_vendor,
                                shutoff_valve_vendor)

    # ── PURGING LINE (MG/COG only) ───────────────────────────────────────
    if purging_line == "yes":
        rows += [
            _row("PURGING LINE", "BALL VALVE", "20 NB", 1, unit_price_override=1800, make="AUDCO/L&T/LEADER"),
            _row("PURGING LINE", "PRESSURE GAUGE WITH TNV", "0-1600 mmWC", 1, unit_price_override=4000, make="HGURU/BAUMER"),
            _row("PURGING LINE", "PRESSURE REGULATING VALVE", "25 NB", 1, unit_price_override=35000, make="NIRMAL"),
            _row("PURGING LINE", "PRESSURE SWITCH HIGH", "", 1, unit_price_override=10000, make="SWITZER"),
            _row("PURGING LINE", "SOLENOID VALVE", "20 NB", 1, unit_price_override=5000, make="MADAS"),
            _row("PURGING LINE", "CHECK VALVE", "20 NB", 1, unit_price_override=3300, make="AUDCO/L&T/LEADER"),
        ]

    # ── PILOT LINE ────────────────────────────────────────────────────────
    pl_media = f"{pilot_line_fuel.upper()} PILOT LINE"
    rows += [
        _row(pl_media, "BALL VALVE", "20 NB", 1,
             unit_price_override=_get_cheapest_ball_valve(20), make="L&T"),
        _row(pl_media, pg_item, '', 1, make=pg_vendor),
        _row(pl_media, "BALL VALVE", "15 NB", 1,
             unit_price_override=_get_cheapest_ball_valve(15), make="L&T"),
        _row(pl_media, "SOLENOID VALVE", "15 NB", 1,
             unit_price_override=_get_cheapest_solenoid_valve(15), make="MADAS"),
    ]
    reg_nb_request = 20
    try:
        reg = select_gas_regulator(reg_nb_request, category="Standard 5 Bar")
        rows.append(_row(
            pl_media, "PRESSURE REGULATING VALVE",
            f'{reg["nb"]} NB, P2={reg["p2_range"]} ({reg["part_code"]})',
            1, unit_price_override=reg["price"], make="MADAS",
        ))
    except ValueError:
        rows.append(_row(
            pl_media, "PRESSURE REGULATING VALVE",
            f'{reg_nb_request} NB', 1, make="MADAS",
        ))
    rows += [
        _row(pl_media, "FLEXIBLE HOSE",
             f'{_get_flexible_hose_price(15)[0]} NB, 1500mm', 1,
             unit_price_override=_get_flexible_hose_price(15)[1], make="BENGAL IND."),
    ]

    # ── ENCON ITEMS ───────────────────────────────────────────────────────
    burner_desc = "ENCON DUAL FUEL Burner" if is_dual else equipment["burner"]["model"]
    rows += [
        _row("ENCON ITEMS", burner_desc,
             f'GAS FLOW: {equipment["burner"]["input_nm3hr"]} Nm3/hr',
             1, unit_price_override=equipment["burner"]["price"]),
        _row("ENCON ITEMS", equipment["blower"]["model"],
             f'{equipment["blower"]["hp"]} HP, {equipment["blower"]["pressure"]} WC, '
             f'{equipment["blower"]["airflow_nm3hr"]} Nm3/hr',
             1, unit_price_override=equipment["blower"]["price_premium"]),
        _row("ENCON ITEMS", "Ignition Transformer", "", 1, make="DANFOSS"),
        _row("ENCON ITEMS", "Sequence Controller", "", 1, make="LINEAR"),
        _row("ENCON ITEMS", "UV Sensor with Air Jacket", "", 1, make="LINEAR"),
        _row(
            "ENCON ITEMS",
            {
                "lpg_10":  "ENCON-PB-LPG-10KW",
                "ng_10":   "ENCON-PB NG 10 KW",
                "lpg_100": "ENCON-PB LPG 100 KW",
                "ng_100":  "ENCON-PB NG 100 KW",
                "cog_100": "ENCON PB COG 100 KW",
            }.get(pilot_burner, "ENCON-PB-LPG-10KW"),
            "", 1,
        ),
    ]

    # HPU — for oil fuels
    hpu = equipment.get("hpu")
    if hpu:
        rows.append(_row(
            "ENCON ITEMS", "Heating and Pumping Unit (HPU)",
            f'{hpu["model"]} — {hpu["unit_kw"]} KW {hpu["variant"]}',
            1, unit_price_override=hpu["price"],
        ))

    # ── MISC ITEMS ────────────────────────────────────────────────────────
    STATIC_SKIP = {"CONTROL PANEL", "HYDRAULIC POWER PACK & CYLINDER"}
    if is_plc or is_plc_agr:
        STATIC_SKIP.update({"P.PID", "RATIO CONTROLLER"})
    if is_pid:
        STATIC_SKIP.add("TEMPERATURE TRANSMITTER")
    from bom.static_items import static_items
    for media, item, ref, qty in static_items():
        if item not in STATIC_SKIP:
            rows.append(_row(media, item, ref, qty))

    rows.append(_row("MISC ITEMS", "CONTROL PANEL", "", 1))
    rows.append(_row("MISC ITEMS", "INSTRUMENTS BALL VALVE", "", 3, make="L&T"))
    if is_plc or is_plc_agr:
        rows.append(_row("MISC ITEMS", "PLC WITH HMI", "", 1))

    df = pd.DataFrame(
        rows,
        columns=["MEDIA", "ITEM NAME", "REFERENCE", "QTY", "MAKE", "UNIT PRICE", "TOTAL"],
    )

    bought_out_total = df.loc[
        (df["MEDIA"] != "ENCON ITEMS"), "TOTAL"
    ].sum()
    encon_total = df.loc[df["MEDIA"] == "ENCON ITEMS", "TOTAL"].sum()
    grand_total = bought_out_total + encon_total

    df = pd.concat(
        [
            df,
            pd.DataFrame(
                [
                    ("", "BOUGHT OUT ITEMS", "", "", "", "", bought_out_total),
                    ("", "ENCON ITEMS",      "", "", "", "", encon_total),
                    ("", "GRAND TOTAL",      "", "", "", "", grand_total),
                ],
                columns=df.columns,
            ),
        ],
        ignore_index=True,
    )

    return df


def build_hlph_manual_df(
    equipment: dict,
    ladle_tons: float = 10.0,
    fuel1_type: str = "ng",
    pressure_gauge_vendor: str = "baumer",
    pilot_burner: str = "auto",
    pipeline_weight_kg: float = 1000.0,
    include_pilot: bool = True,
    pilot_line_fuel: str = "lpg",
) -> pd.DataFrame:
    """
    Manual / simplified HLPH BOM.
    Same as VLPH manual but no rotary joint, has trolley drive.
    No automation items (AGR, orifice, DPT, control valve, pressure switch).
    """

    pg_vendor = pressure_gauge_vendor.upper()
    pg_item = f'PRESSURE GAUGE WITH TNV ({pg_vendor})'
    params = get_hlph_params(ladle_tons)

    rows = []
    air_nb = max(125, equipment["air_duct"]["nb"])

    # ── BOUGHT OUT ITEMS ──────────────────────────────────────────────────
    rows += [
        _row("COMB AIR", "COMPENSATOR", f'{air_nb} NB F150#', 1),
        _row("COMB AIR", pg_item, 'RANGE- 0-1600 mBAR', 1, make=pg_vendor),
        _row("COMB AIR", "BUTTERFLY VALVE",
             f'{equipment["butterfly_valve"]["nb"]} NB', 1,
             unit_price_override=equipment["butterfly_valve"]["price"],
             make=equipment["butterfly_valve"].get("make", "L&T")),
        _row("MISC ITEMS", "CONTROL PANEL", "", 1),
    ]

    # ── FUEL LINE (base items only) ──────────────────────────────────────
    f1_label = FUEL_NAMES.get(fuel1_type, fuel1_type.upper())
    rows += _fuel_line_rows(
        f1_label, fuel1_type, equipment,
        pressure_gauge_vendor=pressure_gauge_vendor,
        base_only=True,
    )

    # ── PILOT LINE (only if pilot burner is included) ────────────────────
    pilot_media = f"{pilot_line_fuel.upper()} PILOT LINE"
    if include_pilot:
        rows += [
            _row(pilot_media, "BALL VALVE", "20 NB", 1,
                 unit_price_override=_get_cheapest_ball_valve(20), make="L&T"),
            _row(pilot_media, pg_item, '', 1, make=pg_vendor),
            _row(pilot_media, "BALL VALVE", "15 NB", 1,
                 unit_price_override=_get_cheapest_ball_valve(15), make="L&T"),
            _row(pilot_media, "SOLENOID VALVE", "15 NB", 1,
                 unit_price_override=_get_cheapest_solenoid_valve(15), make="MADAS"),
        ]
        reg_nb_request = 20
        try:
            reg = select_gas_regulator(reg_nb_request, category="Standard 5 Bar")
            rows.append(_row(
                pilot_media, "PRESSURE REGULATING VALVE",
                f'{reg["nb"]} NB, P2={reg["p2_range"]} ({reg["part_code"]})',
                1, unit_price_override=reg["price"], make="MADAS",
            ))
        except ValueError:
            rows.append(_row(
                pilot_media, "PRESSURE REGULATING VALVE",
                f'{reg_nb_request} NB', 1, make="MADAS",
            ))
        rows += [
            _row(pilot_media, "FLEXIBLE HOSE",
                 f'{_get_flexible_hose_price(15)[0]} NB, 1500mm', 1,
                 unit_price_override=_get_flexible_hose_price(15)[1], make="BENGAL IND."),
        ]

    # ── IN-HOUSE / ENCON ITEMS ────────────────────────────────────────────
    rows += [
        _row("ENCON ITEMS", equipment["burner"]["model"],
             f'GAS FLOW: {equipment["burner"]["input_nm3hr"]} Nm3/hr',
             1, unit_price_override=equipment["burner"]["price"]),
        _row("ENCON ITEMS", "FABRICATION",
             f'{params["ms_structure_kg"]} KG',
             1, unit_price_override=params["ms_structure_kg"] * get_price("FABRICATION RATE")),
    ]

    # HPU — for oil fuels
    hpu = equipment.get("hpu")
    if hpu:
        rows.append(_row(
            "ENCON ITEMS", "Heating and Pumping Unit (HPU)",
            f'{hpu["model"]} — {hpu["unit_kw"]} KW {hpu["variant"]}',
            1, unit_price_override=hpu["price"],
        ))

    rows += [
        _row("MISC ITEMS", "GEARED MOTOR MECHANISM", "3 HP", 1, make="POWERTEK"),
        _row("MISC ITEMS", "TROLLEY WHEEL", "CastIron", 4, make="ENCON"),
        _row("MISC ITEMS", "PLUMMER BLOCK", "MS IS-2062", 6, make="ENCON"),
        _row("MISC ITEMS", "SHAFT (1 long, 2 Short)", "EN-8", 1, make="ENCON"),
        _row("MISC ITEMS", "BEARING", "NU2214", 4, make="ENCON"),
        _row("MISC ITEMS", "SLIEVE", "", 4, make="ENCON"),
        _row("MISC ITEMS", "FRAME", "MS Structure", 1000, make="ENCON"),
        _row("ENCON ITEMS", equipment["blower"]["model"],
             f'{equipment["blower"]["hp"]} HP, {equipment["blower"]["pressure"]} WC, '
             f'{equipment["blower"]["airflow_nm3hr"]} Nm3/hr',
             1, unit_price_override=equipment["blower"]["price_premium"]),
    ]
    if include_pilot:
        rows += [
            _row(
                "ENCON ITEMS",
                {
                    "lpg_10":  "ENCON-PB-LPG-10KW",
                    "ng_10":   "ENCON-PB NG 10 KW",
                    "lpg_100": "ENCON-PB LPG 100 KW",
                    "ng_100":  "ENCON-PB NG 100 KW",
                    "cog_100": "ENCON PB COG 100 KW",
                }.get(pilot_burner, "ENCON-PB-LPG-10KW"),
                "", 1,
            ),
            _row("ENCON ITEMS", "Ignition Transformer", "", 1, make="DANFOSS"),
            _row("ENCON ITEMS", "Sequence Controller", "", 1, make="LINEAR"),
            _row("ENCON ITEMS", "UV Sensor with Air Jacket", "", 1, make="LINEAR"),
        ]
    rows += [
        _row("ENCON ITEMS", "AIR-GAS PIPELINE",
             f'{pipeline_weight_kg:.0f} kg', 1,
             unit_price_override=pipeline_weight_kg * get_price("PIPELINE RATE")),
        _row("ENCON ITEMS", "CERAMIC FIBRE",
             f'{params["ceramic_rolls"]} Rolls @ Rs.{params.get("ceramic_rate", 0):,.0f}/roll',
             params["ceramic_rolls"],
             unit_price_override=params.get("ceramic_rate", 0)),
    ]

    df = pd.DataFrame(
        rows,
        columns=["MEDIA", "ITEM NAME", "REFERENCE", "QTY", "MAKE", "UNIT PRICE", "TOTAL"],
    )

    bought_out_total = df.loc[(df["MEDIA"] != "ENCON ITEMS"), "TOTAL"].sum()
    encon_total = df.loc[df["MEDIA"] == "ENCON ITEMS", "TOTAL"].sum()
    grand_total = bought_out_total + encon_total

    df = pd.concat(
        [
            df,
            pd.DataFrame(
                [
                    ("", "BOUGHT OUT ITEMS", "", "", "", "", bought_out_total),
                    ("", "ENCON ITEMS",      "", "", "", "", encon_total),
                    ("", "GRAND TOTAL",      "", "", "", "", grand_total),
                ],
                columns=df.columns,
            ),
        ],
        ignore_index=True,
    )

    return df
