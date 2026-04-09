import pandas as pd

from bom.static_items import static_items
from bom.price_master import get_price, DB_PATH
from bom.ladle_params import get_vlph_params


BOUGHT_OUT_EXCLUDE_ITEMS = {
    "RATIO CONTROLLER",
}

FUEL_NAMES = {
    "ng": "NG", "lpg": "LPG", "cog": "COG",
    "bg": "BG", "rlng": "RLNG", "ldo": "LDO",
}


def _get_price_fuzzy(item_name: str) -> float:
    """Try exact match first, then partial match on item name."""
    try:
        return get_price(item_name)
    except ValueError:
        pass
    # Fuzzy: try without parenthetical suffix e.g. "MOTORIZED CONTROL VALVE 25NB (Globe)" → "MOTORIZED CONTROL VALVE 25NB"
    base = item_name.split("(")[0].strip()
    if base != item_name:
        try:
            return get_price(base)
        except ValueError:
            pass
    # Try DB LIKE query as last resort
    import sqlite3
    conn = sqlite3.connect(DB_PATH)
    row = conn.execute(
        "SELECT price FROM component_price_master WHERE item LIKE ? LIMIT 1",
        (f"%{base}%",)
    ).fetchone()
    conn.close()
    return float(row[0]) if row else 0


def _row(media: str, item: str, ref: str, qty, unit_price_override=None):
    qty = qty if qty else 1
    if unit_price_override is not None:
        unit_price = unit_price_override
    else:
        unit_price = _get_price_fuzzy(item)

    if unit_price == 0:
        print(f"WARNING: No price found for '{item}'")

    return (media, item, ref, qty, unit_price, unit_price * qty)


OIL_FUELS = {"ldo"}
GAS_FUELS = {"ng", "rlng", "lpg", "cog", "bg"}


def _fuel_line_rows(label: str, fuel_type: str, equipment: dict,
                    control_mode: str = "automatic", auto_control_type: str = "plc",
                    control_valve_vendor: str = "dembla"):
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
                gas_nb = equipment["agr"]["nb"]
                _, gcv_price = _get_valve_price(gas_nb, "control", control_valve_vendor)
                gcv_vendor = "DEMBLA" if control_valve_vendor == "dembla" else "CAIR"
                rows += [
                    _row(media, "ORIFICE PLATE", f'{gas_nb} NB', 1),
                    _row(media, "DPT", "Output 4-20 mA", 1),
                    _row(media, f"CONTROL VALVE ({gcv_vendor})", f'{gas_nb} NB', 1,
                         unit_price_override=gcv_price),
                ]
            elif fuel_type in OIL_FUELS:
                rows += [
                    _row(media, "FLOWMETER", f'{oil_nb} NB', 1),
                    _row(media, "MOTORIZED CONTROL VALVE 025NB (Globe)", "CAIR", 1),
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


def _get_valve_price(nb: int, valve_type: str, vendor: str) -> tuple:
    """Look up valve price from DB by NB and vendor. Returns (item_name, price)."""
    import sqlite3
    conn = sqlite3.connect(DB_PATH)
    nb_str = f'{nb:03d}'

    if vendor == "dembla":
        if valve_type == "control":
            item = f'CONTROL VALVE {nb_str}NB'
        else:
            item = f'SHUT OFF VALVE {nb_str}NB'
        row = conn.execute(
            "SELECT price FROM component_price_master WHERE item=? AND company='DEMBLA'", (item,)
        ).fetchone()
        if not row:
            # Try next bigger NB
            row = conn.execute(
                "SELECT item, price FROM component_price_master WHERE item LIKE ? AND company='DEMBLA' ORDER BY item LIMIT 1",
                (f'{"CONTROL" if valve_type == "control" else "SHUT OFF"} VALVE %NB',)
            ).fetchone()
            if row:
                return row[0], row[1]
        make = "DEMBLA"
    else:
        if valve_type == "control":
            item = f'MOTORIZED CONTROL VALVE {nb_str}NB'
        else:
            item = f'SHUT OFF VALVE {nb_str}NB (Butterfly)'
        row = conn.execute(
            "SELECT price FROM component_price_master WHERE item=? AND company='CAIR'", (item,)
        ).fetchone()
        if not row:
            row = conn.execute(
                "SELECT item, price FROM component_price_master WHERE item LIKE ? AND company='CAIR' ORDER BY item LIMIT 1",
                (f'{"MOTORIZED CONTROL" if valve_type == "control" else "SHUT OFF"} VALVE %NB%',)
            ).fetchone()
            if row:
                return row[0], row[1]
        make = "CAIR"

    conn.close()
    price = row[0] if row else 0
    return item, price


def build_vlph_120t_df(
    equipment: dict,
    ladle_tons: float = 10.0,
    fuel1_type: str = "ng",
    fuel2_type: str = "none",
    equipment2: dict = None,
    control_mode: str = "automatic",
    auto_control_type: str = "agr",
    control_valve_vendor: str = "dembla",
    shutoff_valve_vendor: str = "dembla",
    pressure_gauge_vendor: str = "baumer",
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

    # Air pipe NB — minimum 125 NB, or next bigger from pipe sizing
    air_nb = max(125, equipment["air_duct"]["nb"])

    pg_vendor = pressure_gauge_vendor.upper()
    pg_item = f'PRESSURE GAUGE WITH TNV ({pg_vendor})'
    rows += [
        _row("COMB AIR", "COMPENSATOR", "", 1),
        _row("COMB AIR", pg_item, '', 1),
        _row("COMB AIR", "PRESSURE SWITCH LOW", '', 1),
    ]
    # PLC: air gets orifice plate + DPT + control valve
    if is_plc:
        rows += [
            _row("COMB AIR", "ORIFICE PLATE (Air)", f'{air_nb} NB', 1),
            _row("COMB AIR", "DPT (Air)", f'{air_nb} NB, Output 4-20 mA', 1),
        ]
    # PLC, PLC+AGR, PID: air gets control valve (vendor-selected)
    if is_plc or is_plc_agr or is_pid:
        _, cv_price = _get_valve_price(air_nb, "control", control_valve_vendor)
        vendor_label = "DEMBLA" if control_valve_vendor == "dembla" else "CAIR"
        rows.append(_row(
            "COMB AIR", f"CONTROL VALVE ({vendor_label})",
            f'{air_nb} NB, '
            f'FLOW - {equipment["motorized_control_valve"]["flow_nm3hr"]} Nm3/hr',
            1,
            unit_price_override=cv_price,
        ))
    rows += [
        _row(
            "COMB AIR", "ROTARY JOINT",
            f'{air_nb} NB',
            1,
            unit_price_override=equipment["rotary_joint"]["price"],
        ),
        _row("COMB AIR", "BALL VALVE (Pilot Burner)", "20 NB", 1),
        _row("COMB AIR", "BALL VALVE (UV LINE)", "15 NB", 1),
        _row("COMB AIR", "FLEXIBLE HOSE (Pilot Burner)", "20 NB, 1500 mm", 1),
        _row("COMB AIR", "FLEXIBLE HOSE (UV LINE)", "15 NB, 1500 mm", 1),
    ]

    # ── FUEL 1 LINE ────────────────────────────────────────────────────────
    rows += _fuel_line_rows(f1_label, fuel1_type, equipment, control_mode, auto_control_type, control_valve_vendor)

    # ── FUEL 2 LINE (dual fuel only) ──────────────────────────────────────
    if is_dual:
        rows += _fuel_line_rows(f2_label, fuel2_type, equipment2, control_mode, auto_control_type, control_valve_vendor)

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
