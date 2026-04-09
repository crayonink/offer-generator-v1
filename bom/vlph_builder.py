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


def _row(media: str, item: str, ref: str, qty, unit_price_override=None, make: str = ""):
    qty = qty if qty else 1
    if unit_price_override is not None:
        unit_price = unit_price_override
    else:
        unit_price = _get_price_fuzzy(item)

    if unit_price == 0:
        print(f"WARNING: No price found for '{item}'")

    return (media, item, ref, qty, make, unit_price, unit_price * qty)


OIL_FUELS = {"ldo"}
GAS_FUELS = {"ng", "rlng", "lpg", "cog", "bg"}


def _get_flexible_hose_price(nb: int) -> tuple:
    """Get flexible hose price by NB (next bigger if exact not found). Returns (actual_nb, price)."""
    import sqlite3
    conn = sqlite3.connect(DB_PATH)
    row = conn.execute(
        "SELECT dn, price FROM flexible_hose_master WHERE dn >= ? ORDER BY dn LIMIT 1",
        (nb,)
    ).fetchone()
    conn.close()
    return (int(row[0]), float(row[1])) if row else (nb, 0)


def _get_cheapest_solenoid_valve(nb: int) -> float:
    """Get cheapest MADAS solenoid valve price for a given NB."""
    import sqlite3
    conn = sqlite3.connect(DB_PATH)
    nb_str = f'{nb:03d}'
    row = conn.execute(
        "SELECT list_price FROM solenoidvalve_component_master WHERE size=? ORDER BY list_price ASC LIMIT 1",
        (nb_str,)
    ).fetchone()
    conn.close()
    return float(row[0]) if row else 0


def _get_cheapest_ball_valve(nb: int) -> float:
    """Get cheapest L&T ball valve price for a given NB."""
    import sqlite3
    conn = sqlite3.connect(DB_PATH)
    row = conn.execute(
        "SELECT price FROM component_price_master WHERE company='L&T' AND item LIKE ? ORDER BY price ASC LIMIT 1",
        (f'BALL VALVE {nb:03d}NB%',)
    ).fetchone()
    conn.close()
    return float(row[0]) if row else 0


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
            1, unit_price_override=equipment["ng_gas_train"]["price"], make="MADAS",
        ))

    # Oil line size is always 20 NB
    oil_nb = 20

    # Control-type-specific instrumentation
    if control_mode == "automatic":
        if auto_control_type == "plc":
            if fuel_type in GAS_FUELS:
                gas_nb = equipment["agr"]["nb"]
                _, gcv_price = _get_valve_price(gas_nb, "control", control_valve_vendor)
                gcv_vendor = "DEMBLA" if control_valve_vendor == "dembla" else "CAIR"
                gas_op_nb, gas_op_price = _get_orifice_price(gas_nb)
                rows += [
                    _row(media, "ORIFICE PLATE", f'{gas_op_nb} NB', 1,
                         unit_price_override=gas_op_price, make="ES"),
                    _row(media, "DPT", "", 1, make="HONEYWELL"),
                    _row(media, "CONTROL VALVE", f'{gas_nb} NB', 1,
                         unit_price_override=gcv_price, make=gcv_vendor),
                ]
            elif fuel_type in OIL_FUELS:
                bv_price = _get_cheapest_ball_valve(oil_nb)
                rows += [
                    _row(media, "BALL VALVE", f'{oil_nb} NB', 1,
                         unit_price_override=bv_price, make="L&T"),
                    _row(media, "FLOWMETER", f'{oil_nb} NB', 1, make=""),
                    _row(media, "MOTORIZED CONTROL VALVE", "025NB (Globe)", 1, make="CAIR"),
                    _row(media, "SOLENOID VALVE", f'{oil_nb} NB', 1,
                         unit_price_override=_get_cheapest_solenoid_valve(oil_nb), make="MADAS"),
                    _row(media, "PRESSURE SWITCH LOW", '', 1, make="MADAS"),
                ]
        elif auto_control_type in ("plc_agr", "pid"):
            if fuel_type in GAS_FUELS:
                rows.append(_row(
                    media, "AGR",
                    f'{equipment["agr"]["nb"]} NB',
                    1, unit_price_override=equipment["agr"]["price"], make="ENCON",
                ))
            elif fuel_type in OIL_FUELS:
                rows.append(_row(media, "AIR OIL REGULATOR", f'{oil_nb} NB', 1, make=""))

    # AGR for non-PLC+AGR/PID modes (gas fuels)
    if fuel_type in GAS_FUELS and not (control_mode == "automatic" and auto_control_type in ("plc_agr", "pid")):
        rows.append(_row(
            media, "AGR",
            f'{equipment["agr"]["nb"]} NB',
            1, unit_price_override=equipment["agr"]["price"], make="ENCON",
        ))

    return rows


def _get_orifice_price(nb: int) -> tuple:
    """Look up orifice plate total price by NB (next bigger if exact not found).
    Returns (orifice_nb, price)."""
    import sqlite3
    conn = sqlite3.connect(DB_PATH)
    row = conn.execute(
        "SELECT nb, total_price FROM orifice_plate_master WHERE nb >= ? ORDER BY nb LIMIT 1",
        (nb,)
    ).fetchone()
    conn.close()
    return (int(row[0]), float(row[1])) if row else (nb, 0)


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
        _row("COMB AIR", "COMPENSATOR", "", 1, make="ENCON"),
        _row("COMB AIR", pg_item, '', 1, make=pg_vendor),
        _row("COMB AIR", "PRESSURE SWITCH LOW", '', 1, make="MADAS"),
    ]
    # PLC: air gets orifice plate + DPT + control valve
    if is_plc:
        op_nb, op_price = _get_orifice_price(air_nb)
        rows += [
            _row("COMB AIR", "ORIFICE PLATE", f'{op_nb} NB', 1,
                 unit_price_override=op_price, make="ES"),
            _row("COMB AIR", "DPT", '', 1, make="HONEYWELL"),
        ]
    # PLC, PLC+AGR, PID: air gets control valve (vendor-selected)
    if is_plc or is_plc_agr or is_pid:
        _, cv_price = _get_valve_price(air_nb, "control", control_valve_vendor)
        vendor_label = "DEMBLA" if control_valve_vendor == "dembla" else "CAIR"
        rows.append(_row(
            "COMB AIR", "CONTROL VALVE",
            f'{air_nb} NB, FLOW - {equipment["motorized_control_valve"]["flow_nm3hr"]} Nm3/hr',
            1, unit_price_override=cv_price, make=vendor_label,
        ))
    rows += [
        _row("COMB AIR", "ROTARY JOINT",
             f'{equipment["rotary_joint"]["nb"]} NB', 1,
             unit_price_override=equipment["rotary_joint"]["price"], make="THIRD PARTY"),
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

    # ── FUEL 1 LINE ────────────────────────────────────────────────────────
    rows += _fuel_line_rows(f1_label, fuel1_type, equipment, control_mode, auto_control_type, control_valve_vendor)

    # ── FUEL 2 LINE (dual fuel only) ──────────────────────────────────────
    if is_dual:
        rows += _fuel_line_rows(f2_label, fuel2_type, equipment2, control_mode, auto_control_type, control_valve_vendor)

    # ── NG PILOT LINE ──────────────────────────────────────────────────────
    rows += [
        _row("NG PILOT LINE", "BALL VALVE", "20 NB", 1,
             unit_price_override=_get_cheapest_ball_valve(20), make="L&T"),
        _row("NG PILOT LINE", pg_item, '', 1, make=pg_vendor),
        _row("NG PILOT LINE", "BALL VALVE", "15 NB", 1,
             unit_price_override=_get_cheapest_ball_valve(15), make="L&T"),
        _row("NG PILOT LINE", "SOLENOID VALVE", "15 NB", 1,
             unit_price_override=_get_cheapest_solenoid_valve(15), make="MADAS"),
        _row("NG PILOT LINE", "PRESSURE REGULATING VALVE",
             f'{equipment["agr"]["nb"]} NB', 1, make="MADAS"),
        _row("NG PILOT LINE", "FLEXIBLE HOSE",
             f'{_get_flexible_hose_price(15)[0]} NB, 1500mm', 1,
             unit_price_override=_get_flexible_hose_price(15)[1], make="BENGAL IND."),
    ]


    # ── ENCON ITEMS ────────────────────────────────────────────────────────
    burner_desc = "ENCON DUAL FUEL Burner" if is_dual else equipment["burner"]["model"]
    rows += [
        _row("ENCON ITEMS", burner_desc,
             f'GAS FLOW: {equipment["burner"]["input_nm3hr"]} Nm3/hr',
             1, unit_price_override=equipment["burner"]["price"], make="ENCON"),
        _row("ENCON ITEMS", "BEARING (24026)", "", 2, make="ENCON"),
        _row("ENCON ITEMS", "PLUMMER BLOCK", "", 1,
             unit_price_override=params.get("plummer_block_kg", 300) * 170, make="ENCON"),
        _row("ENCON ITEMS", "SHAFT", "", 1,
             unit_price_override=params.get("shaft_kg", 350) * 120, make="ENCON"),
        _row("ENCON ITEMS", "FABRICATION/ STRUCTURE",
             f'{params["ms_structure_kg"]} kg @ Rs.{params["ms_structure_rate"]}/kg',
             1, unit_price_override=params["ms_structure_cost"], make="ENCON"),
        _row("ENCON ITEMS", "AIR-GAS PIPELINE", "", 1,
             unit_price_override=params.get("pipeline_kg", 1000) * 125, make="ENCON"),
        _row("ENCON ITEMS", "COMBUSTION AIR BLOWER",
             f'{equipment["blower"]["hp"]} HP, {equipment["blower"]["pressure"]} WC, '
             f'{equipment["blower"]["airflow_nm3hr"]} Nm3/hr — {equipment["blower"]["model"]}',
             1, unit_price_override=equipment["blower"]["price_premium"], make="ENCON"),
        _row("ENCON ITEMS", "IGNITION TRANSFORMER", "", 1, make="ENCON"),
        _row("ENCON ITEMS", "SEQUENCE CONTROLLER", "", 1, make="ENCON"),
        _row("ENCON ITEMS", "UV SENSOR WITH AIR JACKET", "", 1, make="ENCON"),
        _row("ENCON ITEMS", "PILOT BURNER", "", 1, make="ENCON"),
        _row("ENCON ITEMS", "CERAMIC FIBRE",
             f'{params["ceramic_rolls"]} Rolls @ Rs.{params.get("ceramic_rate", 0):,.0f}/roll',
             params["ceramic_rolls"],
             unit_price_override=params.get("ceramic_rate", 0), make="ENCON"),
    ]

    # ── MISC ITEMS ─────────────────────────────────────────────────────────
    STATIC_SKIP = {"CONTROL PANEL"}
    if is_plc or is_plc_agr:
        STATIC_SKIP.update({"P.PID", "RATIO CONTROLLER"})
    if is_pid:
        STATIC_SKIP.add("TEMPERATURE TRANSMITTER")
    for media, item, ref, qty in static_items():
        if item not in STATIC_SKIP:
            rows.append(_row(media, item, ref, qty, make="ENCON"))

    rows.append(_row("MISC ITEMS", "CONTROL PANEL", "", 1, make="ENCON"))
    rows.append(_row("MISC ITEMS", "INSTRUMENTS BALL VALVE", "", 3, make="L&T"))
    if is_plc or is_plc_agr:
        rows.append(_row("MISC ITEMS", "PLC WITH HMI", "", 1, make="ENCON"))

    df = pd.DataFrame(
        rows,
        columns=["MEDIA", "ITEM NAME", "REFERENCE", "QTY", "MAKE", "UNIT PRICE", "TOTAL"],
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
                    ("", "BOUGHT OUT ITEMS",     "", "", "", "", bought_out_total),
                    ("", "ENCON ITEMS",          "", "", "", "", encon_total),
                    ("", "GRAND TOTAL",          "", "", "", "", grand_total),
                ],
                columns=df.columns,
            ),
        ],
        ignore_index=True,
    )

    return df
