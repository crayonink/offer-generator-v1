import pandas as pd

from bom.static_items import static_items
from bom.price_master import get_price, DB_PATH
from bom.ladle_params import get_vlph_params
from bom.selectors.gas_regulator_selector import select_gas_regulator


BOUGHT_OUT_EXCLUDE_ITEMS = {
    "RATIO CONTROLLER",
}

FUEL_NAMES = {
    # Gas fuels
    "ng":   "NG",
    "lpg":  "LPG",
    "cog":  "COG",
    "bg":   "BG",
    "rlng": "RLNG",
    "mg":   "Mixed Gas",
    # Oil-based fuels (sub-categories of OIL)
    "ldo":   "LDO",
    "fo":    "Furnace Oil",
    "lshs":  "LSHS",
    "hsd":   "HSD",
    "sko":   "SKO",
}


def _get_price_fuzzy(item_name: str) -> float:
    """Exact match only — no fuzzy lookups."""
    try:
        return get_price(item_name)
    except ValueError:
        return 0


def _get_company(item_name: str) -> str:
    """Look up company/make from component_price_master."""
    import sqlite3
    conn = sqlite3.connect(DB_PATH)
    row = conn.execute(
        "SELECT company FROM component_price_master WHERE item=? LIMIT 1",
        (item_name,),
    ).fetchone()
    conn.close()
    return row[0] if row and row[0] else ""


def _row(media: str, item: str, ref: str, qty, unit_price_override=None, make=None):
    qty = qty if qty else 1
    if unit_price_override is not None:
        unit_price = unit_price_override
    else:
        unit_price = _get_price_fuzzy(item)

    if unit_price == 0:
        print(f"WARNING: No price found for '{item}'")

    if make is None:
        make = _get_company(item)

    return (media, item, ref, qty, make, unit_price, unit_price * qty)


OIL_FUELS = {"ldo", "fo", "lshs", "hsd", "sko"}
GAS_FUELS = {"ng", "rlng", "lpg", "cog", "bg", "mg"}


def _get_gate_valve_price(nb: int) -> tuple:
    """L&T 113-8/IBR gate valve lookup by NB (next bigger if exact not found).
    Returns (actual_nb, price)."""
    import sqlite3, re
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute(
        "SELECT item, price FROM component_price_master WHERE item LIKE 'GATE VALVE % NB' ORDER BY item"
    ).fetchall()
    conn.close()
    candidates = []
    for item, price in rows:
        m = re.match(r'GATE VALVE (\d+) NB', item)
        if m:
            candidates.append((int(m.group(1)), float(price)))
    candidates.sort()
    for cnb, cprice in candidates:
        if cnb >= nb:
            return cnb, cprice
    return (candidates[-1][0], candidates[-1][1]) if candidates else (nb, 0)


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


SOLENOID_VALVE_DISCOUNT = 0.45  # 45% discount on MADAS list price


def _get_cheapest_solenoid_valve(nb: int) -> float:
    """Get cheapest MADAS solenoid valve from the
    'AUTOMATIC RESET WITH FLOW REGULATION' section for a given NB,
    with 45% discount applied to the list price.

    MADAS only makes this SV family up to DN100. If the requested NB is
    larger than anything in the catalogue, fall back to the largest
    available (DN100) — engineering convention is to step the main line
    down with a reducer at the SV.
    """
    import sqlite3
    conn = sqlite3.connect(DB_PATH)
    nb_str = f'{nb:03d}'

    # 1. Exact match
    row = conn.execute(
        """
        SELECT list_price
        FROM solenoidvalve_component_master
        WHERE size = ?
          AND section LIKE '%AUTOMATIC RESET WITH FLOW REGULATION%'
        ORDER BY list_price ASC
        LIMIT 1
        """,
        (nb_str,),
    ).fetchone()

    # 2. Fallback: largest available size ≤ requested NB (stepped down with reducer)
    if not row:
        row = conn.execute(
            """
            SELECT list_price
            FROM solenoidvalve_component_master
            WHERE section LIKE '%AUTOMATIC RESET WITH FLOW REGULATION%'
              AND CAST(size AS INTEGER) <= ?
              AND CAST(size AS INTEGER) > 0
            ORDER BY CAST(size AS INTEGER) DESC, list_price ASC
            LIMIT 1
            """,
            (nb,),
        ).fetchone()

    # 3. Last resort: cheapest row overall in this section (smallest available)
    if not row:
        row = conn.execute(
            """
            SELECT list_price
            FROM solenoidvalve_component_master
            WHERE section LIKE '%AUTOMATIC RESET WITH FLOW REGULATION%'
              AND CAST(size AS INTEGER) > 0
            ORDER BY CAST(size AS INTEGER) ASC, list_price ASC
            LIMIT 1
            """,
        ).fetchone()

    conn.close()
    if not row:
        return 0
    return float(row[0]) * (1 - SOLENOID_VALVE_DISCOUNT)


def _get_cheapest_ball_valve(nb: int) -> float:
    """Get cheapest L&T ball valve price for a given NB."""
    import sqlite3
    conn = sqlite3.connect(DB_PATH)
    row = conn.execute(
        "SELECT price FROM component_price_master WHERE company='L&T' AND item LIKE ? ORDER BY price ASC LIMIT 1",
        (f'BALL VALVE {nb} NB%',)
    ).fetchone()
    conn.close()
    return float(row[0]) if row else 0


def _get_cheapest_butterfly_valve(nb: int) -> float:
    """Get cheapest L&T butterfly valve (lever preferred) price for a given NB
    from lt_butterfly_valve_master. Falls back to gear type if lever is not
    available at that NB (levers only go up to 300 NB)."""
    import sqlite3
    conn = sqlite3.connect(DB_PATH)
    row = conn.execute(
        "SELECT price FROM lt_butterfly_valve_master WHERE nb = ? ORDER BY price ASC LIMIT 1",
        (nb,),
    ).fetchone()
    conn.close()
    return float(row[0]) if row else 0


CERAMIC_FIBRE_KG_PER_ROLL = 14.0


def lookup_ladle_fab_pipeline(ladle_tons: float, preheater_type: str) -> dict:
    """Look up fabrication, pipeline and ceramic-fibre weight for a given ladle
    capacity and preheater type (vertical/horizontal) from the
    fabrication_ladle_mapping table. Picks the row whose capacity is closest
    to the requested tons. Returns {} if no rows.

    The stored weights already include the 10% margin. Ceramic rolls are
    derived as ceil(ceramic_kg / 14) — 14 kg per roll, rounded up so a
    partial roll always counts."""
    import sqlite3, math
    if not ladle_tons or preheater_type not in ('vertical', 'horizontal'):
        return {}
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute(
        "SELECT ladle_capacity_ton, fabrication_kg, pipeline_kg, ceramic_kg "
        "FROM fabrication_ladle_mapping WHERE preheater_type = ?",
        (preheater_type,),
    ).fetchall()
    conn.close()
    if not rows:
        return {}
    best = min(rows, key=lambda r: abs(float(r[0]) - float(ladle_tons)))
    ceramic_kg = float(best[3])
    return {
        "fabrication_kg": round(float(best[1])),
        "pipeline_kg":    round(float(best[2])),
        "ceramic_kg":     round(ceramic_kg, 2),
        "ceramic_rolls":  int(math.ceil(ceramic_kg / CERAMIC_FIBRE_KG_PER_ROLL)),
    }


def _get_cheapest_shutoff_valve(nb: int) -> float:
    """Cheapest pneumatic shut-off valve price for a given NB."""
    import sqlite3
    conn = sqlite3.connect(DB_PATH)
    row = conn.execute(
        "SELECT price FROM component_price_master WHERE item LIKE ? ORDER BY price ASC LIMIT 1",
        (f'SHUT OFF VALVE {nb} NB%',),
    ).fetchone()
    conn.close()
    return float(row[0]) if row else 0


def _cog_line_rows(media: str, equipment: dict,
                   control_mode: str, auto_control_type: str,
                   pressure_gauge_vendor: str,
                   butterfly_valve_vendor: str = "lt_lever",
                   shutoff_valve_vendor: str = "aira",
                   base_only: bool = False):
    """
    Coke Oven Gas BOM — discrete components, prices pulled from DB.
    Matches the ENCON COG pricelist layout.

    PLC mode: all 9 items incl. ORIFICE PLATE + DPT
    PLC+AGR / PID / manual: ORIFICE PLATE + DPT replaced by AGR
    """
    from calculations.pipes import STANDARD_PIPE_NB
    from bom.selectors.air_valve_selector import select_butterfly_valve

    gas_pipe_nb = equipment["pipe"].ng_pipe_nb

    # Control valve NB is one pipe size smaller than the gas pipe NB
    try:
        cv_nb = STANDARD_PIPE_NB[max(0, STANDARD_PIPE_NB.index(gas_pipe_nb) - 1)]
    except ValueError:
        cv_nb = gas_pipe_nb

    # Gate valve — L&T, NB-scaled
    gv_nb, gv_price = _get_gate_valve_price(gas_pipe_nb)

    # Shut off valve — vendor-selected, NB-scaled
    _, shutoff_price = _get_valve_price(gas_pipe_nb, "shutoff", shutoff_valve_vendor)

    # Pneumatic control valve — DEMBLA, one pipe size smaller
    _, pcv_price = _get_valve_price(cv_nb, "control", "dembla")

    # Butterfly valve — follows user's L&T vendor choice
    try:
        bfv = select_butterfly_valve(gas_pipe_nb, vendor=butterfly_valve_vendor)
    except ValueError:
        if butterfly_valve_vendor == "lt_lever":
            try:
                bfv = select_butterfly_valve(gas_pipe_nb, vendor="lt_gear")
            except ValueError:
                bfv = {"nb": gas_pipe_nb, "price": 0, "make": "L&T"}
        else:
            bfv = {"nb": gas_pipe_nb, "price": 0, "make": "L&T"}

    is_plc = control_mode == "automatic" and auto_control_type == "plc"
    needs_agr = (
        control_mode == "manual"
        or (control_mode == "automatic" and auto_control_type in ("plc_agr", "pid"))
    )

    pg_vendor = pressure_gauge_vendor.upper()
    # Use clean display name; vendor only shown in MAKE column.
    # Keep vendor-suffixed key for DB price lookup.
    pg_price = _get_price_fuzzy(f'PRESSURE GAUGE WITH TNV ({pg_vendor})')
    pg_item = "PRESSURE GAUGE WITH TNV"

    rows = [
        _row(media, "GATE VALVE", f'{gv_nb} NB', 1,
             unit_price_override=gv_price, make="L&T"),
        _row(media, pg_item, "", 1, unit_price_override=pg_price, make=pg_vendor),
        _row(media, "SHUT OFF VALVE", f'{gas_pipe_nb} NB', 1,
             unit_price_override=shutoff_price, make=shutoff_valve_vendor.upper()),
        _row(media, "BUTTERFLY VALVE", f'{bfv["nb"]} NB', 1,
             unit_price_override=bfv["price"], make=bfv.get("make", "L&T")),
    ]

    if not base_only:
        rows.append(_row(media, "PRESSURE SWITCH LOW", "Set PT - L", 1, make="MADAS"))

        if is_plc:
            rows += [
                _row(media, "ORIFICE PLATE", "Output: 4-20 mA, 230 V AC", 1,
                     unit_price_override=_get_price_fuzzy("ORIFICE PLATE (COG)"),
                     make="ENGINEERING SPECIALITY"),
                _row(media, "DPT", "", 1,
                     unit_price_override=_get_price_fuzzy("DPT (COG)"),
                     make="HONEYWELL"),
            ]

        rows.append(_row(media, "PNEUMATIC CONTROL VALVE", f'{cv_nb} NB', 1,
             unit_price_override=pcv_price, make="DEMBLA"))

        if needs_agr:
            rows.append(_row(
                media, "AGR",
                f'{equipment["agr"]["nb"]} NB',
                1, unit_price_override=equipment["agr"]["price"],
            ))

    return rows


def _bfg_line_rows(media: str, equipment: dict,
                   control_mode: str, auto_control_type: str,
                   pressure_gauge_vendor: str,
                   base_only: bool = False):
    """
    Blast Furnace Gas BOM — discrete components instead of a packaged gas train.
    BFG runs at low pressure and with variable composition, so a pre-assembled
    IAPL/MADAS gas train isn't used. Engineer builds the main line from:
      - Butterfly valve (isolation — BFG needs full-bore, low-dP)
      - Shut-off valve x 2 (double-block safety, replaces solenoid)
      - Pressure gauge with TNV
      - Pressure switch low
      - Orifice plate + DPT (mass-flow measurement — PLC mode only)
      - Pneumatic control valve (mass-flow regulation — always)
      - Rotary joint
      - AGR (only on PLC+AGR / PID / manual control)
    """
    from calculations.pipes import STANDARD_PIPE_NB
    from bom.selectors.rotary_joint_selector import select_rotary_joint

    gas_pipe_nb = equipment["pipe"].ng_pipe_nb

    # Control valve is one pipe size smaller than the gas pipe NB
    try:
        cv_nb = STANDARD_PIPE_NB[max(0, STANDARD_PIPE_NB.index(gas_pipe_nb) - 1)]
    except ValueError:
        cv_nb = gas_pipe_nb

    # Rotary joint sized to gas pipe NB
    try:
        rj = select_rotary_joint(gas_pipe_nb)
    except ValueError:
        rj = {"nb": gas_pipe_nb, "price": 0, "company": "ENCON"}

    pg_vendor = pressure_gauge_vendor.upper()
    pg_price = _get_price_fuzzy(f'PRESSURE GAUGE WITH TNV ({pg_vendor})')

    bfv_price = _get_cheapest_butterfly_valve(gas_pipe_nb)
    so_price  = _get_cheapest_shutoff_valve(gas_pipe_nb)
    _, pcv_price = _get_valve_price(cv_nb, "control", "dembla")

    is_plc = control_mode == "automatic" and auto_control_type == "plc"

    rows = [
        _row(media, "BUTTERFLY VALVE", f'{gas_pipe_nb} NB', 1,
             unit_price_override=bfv_price, make="L&T"),
        _row(media, "SHUT OFF VALVE", f'{gas_pipe_nb} NB', 2,
             unit_price_override=so_price, make="AIRA"),
        _row(media, "PRESSURE GAUGE WITH TNV", f'{gas_pipe_nb} NB', 1,
             unit_price_override=pg_price, make=pg_vendor),
        _row(media, "PRESSURE SWITCH LOW", "Set PT - L", 1, make="MADAS"),
    ]

    # Mass flow control elements — orifice + DPT only in PLC (flow-measurement
    # needed only when the PLC regulates the valve directly, no AGR).
    if is_plc:
        rows += [
            _row(media, "ORIFICE PLATE", "Output: 4-20 mA, 230 V AC", 1,
                 unit_price_override=_get_price_fuzzy("ORIFICE PLATE (COG)"),
                 make="ENGINEERING SPECIALITY"),
            _row(media, "DPT", "", 1,
                 unit_price_override=_get_price_fuzzy("DPT (COG)"),
                 make="HONEYWELL"),
        ]
    rows.append(_row(media, "PNEUMATIC CONTROL VALVE", f'{cv_nb} NB', 1,
                     unit_price_override=pcv_price, make="DEMBLA"))

    rows.append(_row(media, "ROTARY JOINT", f'{rj["nb"]} NB', 1,
                     unit_price_override=rj["price"], make="ENCON"))

    if not base_only:
        needs_agr = (
            control_mode == "manual"
            or (control_mode == "automatic" and auto_control_type in ("plc_agr", "pid"))
        )
        if needs_agr:
            rows.append(_row(
                media, "AGR",
                f'{equipment["agr"]["nb"]} NB',
                1, unit_price_override=equipment["agr"]["price"],
            ))

    return rows


def _mix_gas_line_rows(media: str, equipment: dict,
                       control_mode: str, auto_control_type: str,
                       pressure_gauge_vendor: str,
                       butterfly_valve_vendor: str = "lt_lever",
                       shutoff_valve_vendor: str = "aira",
                       base_only: bool = False):
    """
    Mix Gas BOM — discrete components instead of a packaged gas train.
    Structure matches the ENCON Mix Gas pricelist layout.

    Common to all control modes: gate valve, pressure gauge with TNV,
    pressure switch low, pneumatic shut-off valve, pneumatic control valve,
    butterfly valve, rotary joint.

    PLC adds: (ORIFICE PLATE WITH DPT — currently on hold, skipped)
    PLC+AGR / PID / manual: AGR is added by the consolidated block below.
    """
    from calculations.pipes import STANDARD_PIPE_NB
    from bom.selectors.air_valve_selector import select_butterfly_valve
    from bom.selectors.rotary_joint_selector import select_rotary_joint

    pg_vendor = pressure_gauge_vendor.upper()
    # Use clean display name; vendor only shown in MAKE column.
    # Keep vendor-suffixed key for DB price lookup.
    pg_price = _get_price_fuzzy(f'PRESSURE GAUGE WITH TNV ({pg_vendor})')
    pg_item = "PRESSURE GAUGE WITH TNV"

    gas_pipe_nb = equipment["pipe"].ng_pipe_nb

    # Control valve is one pipe size smaller than the gas pipe NB
    try:
        cv_nb = STANDARD_PIPE_NB[max(0, STANDARD_PIPE_NB.index(gas_pipe_nb) - 1)]
    except ValueError:
        cv_nb = gas_pipe_nb

    # Shut-off valve — vendor-selected, sized to gas pipe NB
    _, shutoff_price = _get_valve_price(gas_pipe_nb, "shutoff", shutoff_valve_vendor)

    # Pneumatic control valve — always DEMBLA for Mix Gas (pneumatic only).
    # When additional pneumatic vendors are added, expand this.
    _, pcv_price = _get_valve_price(cv_nb, "control", "dembla")
    pcv_make = "DEMBLA"

    # Butterfly valve sized to gas pipe NB — follow user's vendor choice,
    # fall back from Lever to Gear if the requested NB is out of Lever range.
    try:
        bfv = select_butterfly_valve(gas_pipe_nb, vendor=butterfly_valve_vendor)
    except ValueError:
        if butterfly_valve_vendor == "lt_lever":
            try:
                bfv = select_butterfly_valve(gas_pipe_nb, vendor="lt_gear")
            except ValueError:
                bfv = {"nb": gas_pipe_nb, "price": 0, "make": "L&T"}
        else:
            bfv = {"nb": gas_pipe_nb, "price": 0, "make": "L&T"}

    # Rotary joint sized to gas pipe NB
    try:
        rj = select_rotary_joint(gas_pipe_nb)
    except ValueError:
        rj = {"nb": gas_pipe_nb, "price": 0, "company": "ENCON"}

    # Gate valve sized to gas pipe NB (L&T 113-8/IBR Class 150). Use next
    # bigger NB if the exact size isn't listed (65 NB and 125 NB are blank).
    gv_nb, gv_price = _get_gate_valve_price(gas_pipe_nb)

    is_plc = control_mode == "automatic" and auto_control_type == "plc"

    rows = [
        _row(media, "GATE VALVE", f'{gv_nb} NB', 1,
             unit_price_override=gv_price, make="L&T"),
        _row(media, pg_item, f'{gas_pipe_nb} NB', 1, unit_price_override=pg_price, make=pg_vendor),
        _row(media, "SHUT OFF VALVE", f'{gas_pipe_nb} NB', 1,
             unit_price_override=shutoff_price, make=shutoff_valve_vendor.upper()),
        _row(media, "BUTTERFLY VALVE", f'{bfv["nb"]} NB', 1,
             unit_price_override=bfv["price"], make=bfv.get("make", "L&T")),
        _row(media, "ROTARY JOINT", f'{rj["nb"]} NB', 1,
             unit_price_override=rj["price"], make=rj.get("company", "ENCON")),
    ]

    if not base_only:
        rows.append(_row(media, "PRESSURE SWITCH LOW", "", 1, make="MADAS"))

        if is_plc:
            rows += [
                _row(media, "ORIFICE PLATE", "", 1,
                     unit_price_override=_get_price_fuzzy("ORIFICE PLATE (Gas)"),
                     make="ENGINEERING SPECIALITY"),
                _row(media, "DPT", "", 1,
                     unit_price_override=_get_price_fuzzy("DPT"),
                     make="HONEYWELL"),
            ]

        rows.append(_row(media, "PNEUMATIC CONTROL VALVE", f'{cv_nb} NB', 1,
             unit_price_override=pcv_price, make=pcv_make))

        needs_agr = (
            control_mode == "manual"
            or (control_mode == "automatic" and auto_control_type in ("plc_agr", "pid"))
        )
        if needs_agr:
            rows.append(_row(
                media, "AGR",
                f'{equipment["agr"]["nb"]} NB',
                1, unit_price_override=equipment["agr"]["price"],
            ))

    return rows


def _fuel_line_rows(label: str, fuel_type: str, equipment: dict,
                    control_mode: str = "automatic", auto_control_type: str = "plc",
                    control_valve_vendor: str = "dembla",
                    pressure_gauge_vendor: str = "baumer",
                    butterfly_valve_vendor: str = "lt_lever",
                    shutoff_valve_vendor: str = "aira",
                    base_only: bool = False):
    """Generate fuel line BOM rows for a single fuel.
    base_only=True skips all instrumentation (AGR, orifice, DPT, control valve)."""
    media = f"{label} LINE"
    rows = []

    # Mix Gas has a dedicated discrete-component BOM (no packaged gas train)
    if fuel_type == "mg":
        return _mix_gas_line_rows(
            media, equipment, control_mode, auto_control_type,
            pressure_gauge_vendor, butterfly_valve_vendor, shutoff_valve_vendor,
            base_only=base_only,
        )

    # Coke Oven Gas has its own discrete-component BOM
    if fuel_type == "cog":
        return _cog_line_rows(
            media, equipment, control_mode, auto_control_type,
            pressure_gauge_vendor, butterfly_valve_vendor, shutoff_valve_vendor,
            base_only=base_only,
        )

    # Blast Furnace Gas has its own discrete-component BOM (no packaged gas train)
    if fuel_type == "bg":
        return _bfg_line_rows(
            media, equipment, control_mode, auto_control_type,
            pressure_gauge_vendor,
            base_only=base_only,
        )

    # Gas train (gas fuels only — BFG/COG/MG already returned above with
    # discrete-component BOMs). RLNG uses the NG pre-assembled gas train but
    # prints with the fuel-specific prefix so the offer reads correctly.
    if fuel_type in GAS_FUELS:
        fuel_prefix = {"rlng": "RLNG ", "lpg": "LPG "}.get(fuel_type, "")
        gas_train_name = f'{fuel_prefix}GAS TRAIN {equipment["ng_gas_train"]["max_flow"]:.0f} NM3/Hr'
        rows.append(_row(
            media, gas_train_name,
            f'{equipment["ng_gas_train"]["inlet_nb"]} x '
            f'{equipment["ng_gas_train"]["outlet_nb"]} NB',
            1, unit_price_override=equipment["ng_gas_train"]["price"], make="MADAS",
        ))

    # Oil line size is always 20 NB
    oil_nb = 20

    # Control-type-specific instrumentation (skipped in base_only / manual BOM)
    if not base_only:
        if control_mode == "automatic":
            if auto_control_type == "plc":
                if fuel_type in ("ng", "lpg", "rlng"):
                    # Orifice plate = gas train outlet NB (2nd DN)
                    import re as _re
                    gt_outlet_nb = int(_re.sub(r'[^\d]', '', str(equipment["ng_gas_train"]["outlet_nb"])) or 0)
                    gas_op_nb, gas_op_price = _get_orifice_price(gt_outlet_nb)
                    # Control valve = same NB as gas train outlet
                    gas_cv_nb = gt_outlet_nb
                    _, gcv_price = _get_valve_price(gas_cv_nb, "control", control_valve_vendor)
                    gcv_vendor = control_valve_vendor.upper()
                    rows += [
                        _row(media, "ORIFICE PLATE", f'{gas_op_nb} NB', 1,
                             unit_price_override=gas_op_price, make="ENCON"),
                        _row(media, "DPT", "", 1, make="HONEYWELL"),
                        _row(media, "CONTROL VALVE", f'{gas_cv_nb} NB', 1,
                             unit_price_override=gcv_price, make=gcv_vendor),
                    ]
                elif fuel_type in OIL_FUELS:
                    bv_price = _get_cheapest_ball_valve(oil_nb)
                    rows += [
                        _row(media, "BALL VALVE", f'{oil_nb} NB', 1,
                             unit_price_override=bv_price, make="L&T"),
                        _row(media, "FLOWMETER", f'{oil_nb} NB', 1, make="ELETA"),
                        _row(media, "MOTORIZED CONTROL VALVE", "25 NB (Globe)", 1,
                             unit_price_override=_get_price_fuzzy("MOTORIZED CONTROL VALVE 25 NB (Globe)"),
                             make="CAIR"),
                        _row(media, "SOLENOID VALVE", f'{oil_nb} NB', 1,
                             unit_price_override=_get_cheapest_solenoid_valve(oil_nb), make="MADAS"),
                        _row(media, "PRESSURE SWITCH LOW", '', 1, make="MADAS"),
                    ]
            elif auto_control_type in ("plc_agr", "pid"):
                if fuel_type in OIL_FUELS:
                    rows.append(_row(media, "AIR OIL REGULATOR", f'{oil_nb} NB', 1))
                # Gas-fuel AGR is added by the consolidated block below

        # AGR — only appears for gas fuels in modes that explicitly use it:
        #   manual, plc_agr, pid. In pure PLC mode the ratio is controlled by
        #   orifice plate + DPT + control valve (no AGR needed).
        needs_agr = (
            fuel_type in GAS_FUELS
            and (
                control_mode == "manual"
                or (control_mode == "automatic" and auto_control_type in ("plc_agr", "pid"))
            )
        )
        if needs_agr:
            rows.append(_row(
                media, "AGR",
                f'{equipment["agr"]["nb"]} NB',
                1, unit_price_override=equipment["agr"]["price"],
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


def _get_valve_price(nb, valve_type: str, vendor: str) -> tuple:
    """Look up valve price from DB by NB and vendor. Returns (item_name, price).

    Exact-match first; on miss, fall back to the smallest NB >= requested
    (numeric comparison — alphabetical sort mis-picks 200 NB when asked for 20 NB).
    """
    import sqlite3, re
    conn = sqlite3.connect(DB_PATH)
    nb = int(re.sub(r'[^\d]', '', str(nb)) or 0)
    nb_str = f'{nb} NB'

    if vendor in ("dembla", "aira"):
        company = vendor.upper()
        item = f'CONTROL VALVE {nb_str}' if valve_type == "control" else f'SHUT OFF VALVE {nb_str}'
        like_prefix = 'CONTROL VALVE ' if valve_type == "control" else 'SHUT OFF VALVE '
    elif vendor == "cair":
        company = "CAIR"
        item = (f'MOTORIZED CONTROL VALVE {nb_str}' if valve_type == "control"
                else f'SHUT OFF VALVE {nb_str} (Butterfly)')
        like_prefix = 'MOTORIZED CONTROL VALVE ' if valve_type == "control" else 'SHUT OFF VALVE '
    else:
        conn.close()
        return (f'CONTROL VALVE {nb_str}' if valve_type == "control"
                else f'SHUT OFF VALVE {nb_str}'), 0

    # Exact NB match
    row = conn.execute(
        "SELECT price FROM component_price_master WHERE item=? AND company=?",
        (item, company),
    ).fetchone()
    if row:
        conn.close()
        return item, row[0]

    # Fallback: smallest NB >= requested, sorted numerically
    candidates = conn.execute(
        "SELECT item, price FROM component_price_master WHERE item LIKE ? AND company=?",
        (f'{like_prefix}% NB%', company),
    ).fetchall()
    conn.close()

    best = None
    for cand_item, cand_price in candidates:
        m = re.search(r'(\d+)\s*NB', cand_item)
        if not m:
            continue
        cand_nb = int(m.group(1))
        if cand_nb >= nb and (best is None or cand_nb < best[0]):
            best = (cand_nb, cand_item, cand_price)

    if best:
        return best[1], best[2]
    return item, 0


def build_vlph_120t_df(
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
    num_burners: int = 1,
    ms_structure_kg_override: float = 0.0,
    ceramic_rolls_override: int = 0,
    hood_type: str = "up_down",
) -> pd.DataFrame:
    """
    Builds VLPH BOM DataFrame.
    For dual fuel, equipment2 contains the second fuel's gas line equipment.
    hood_type: 'up_down' (hydraulic) or 'swivel' (geared) — controls whether
    HYDRAULIC POWER PACK & CYLINDER is included.
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
    # Use clean display name; vendor only shown in MAKE column.
    # Keep vendor-suffixed key for DB price lookup.
    pg_price = _get_price_fuzzy(f'PRESSURE GAUGE WITH TNV ({pg_vendor})')
    pg_item = "PRESSURE GAUGE WITH TNV"
    rows += [
        _row("COMB AIR", "COMPENSATOR", "", 1),
        _row("COMB AIR", pg_item, '', 1, unit_price_override=pg_price, make=pg_vendor),
        _row("COMB AIR", "PRESSURE SWITCH LOW", '', 1, make="MADAS"),
    ]
    # PLC: air gets orifice plate + DPT + control valve
    if is_plc:
        op_nb, op_price = _get_orifice_price(air_nb)
        rows += [
            _row("COMB AIR", "ORIFICE PLATE", f'{op_nb} NB', 1,
                 unit_price_override=op_price, make="ENCON"),
            _row("COMB AIR", "DPT", '', 1, make="HONEYWELL"),
        ]
    # PLC, PLC+AGR, PID: air gets control valve (vendor-selected).
    # Control valve NB is one pipe size smaller than the air pipe NB.
    if is_plc or is_plc_agr or is_pid:
        from calculations.pipes import STANDARD_PIPE_NB
        try:
            cv_nb = STANDARD_PIPE_NB[max(0, STANDARD_PIPE_NB.index(air_nb) - 1)]
        except ValueError:
            cv_nb = air_nb
        _, cv_price = _get_valve_price(cv_nb, "control", control_valve_vendor)
        vendor_label = control_valve_vendor.upper()
        rows.append(_row(
            "COMB AIR", "CONTROL VALVE",
            f'{cv_nb} NB, FLOW - {equipment["motorized_control_valve"]["flow_nm3hr"]} Nm3/hr',
            1, unit_price_override=cv_price, make=vendor_label,
        ))
        # Butterfly valve (L&T) — sized to air pipe NB
        bfv = equipment["butterfly_valve"]
        rows.append(_row(
            "COMB AIR", "BUTTERFLY VALVE",
            f'{bfv["nb"]} NB',
            1, unit_price_override=bfv["price"], make=bfv["make"],
        ))
    rows += [
        _row("COMB AIR", "ROTARY JOINT",
             f'{equipment["rotary_joint"]["nb"]} NB', 1,
             unit_price_override=equipment["rotary_joint"]["price"], make="ENCON"),
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
    rows += _fuel_line_rows(f1_label, fuel1_type, equipment, control_mode, auto_control_type, control_valve_vendor, pressure_gauge_vendor, butterfly_valve_vendor, shutoff_valve_vendor)

    # ── FUEL 2 LINE (dual fuel only) ──────────────────────────────────────
    if is_dual:
        rows += _fuel_line_rows(f2_label, fuel2_type, equipment2, control_mode, auto_control_type, control_valve_vendor, pressure_gauge_vendor, butterfly_valve_vendor, shutoff_valve_vendor)

    # ── PURGING LINE LINE (MG/COG only, when user enabled it) ─────────
    # Prices are specific to the nitrogen purging assembly and don't match
    # the regular gas-line items' prices, so they're inlined here.
    if purging_line == "yes":
        rows += [
            _row("PURGING LINE", "BALL VALVE",                "20 NB",       1, unit_price_override=1800,  make="AUDCO/L&T/LEADER"),
            _row("PURGING LINE", "PRESSURE GAUGE WITH TNV",   "0-1600 mmWC", 1, unit_price_override=4000,  make="HGURU/BAUMER"),
            _row("PURGING LINE", "PRESSURE REGULATING VALVE", "25 NB",       1, unit_price_override=35000, make="NIRMAL"),
            _row("PURGING LINE", "PRESSURE SWITCH HIGH",      "",            1, unit_price_override=10000, make="SWITZER"),
            _row("PURGING LINE", "SOLENOID VALVE",            "20 NB",       1, unit_price_override=5000,  make="MADAS"),
            _row("PURGING LINE", "CHECK VALVE",               "20 NB",       1, unit_price_override=3300,  make="AUDCO/L&T/LEADER"),
        ]

    # ── PILOT LINE ─────────────────────────────────────────────────────────
    pl_media = f"{pilot_line_fuel.upper()} PILOT LINE"
    rows += [
        _row(pl_media, "BALL VALVE", "20 NB", 1,
             unit_price_override=_get_cheapest_ball_valve(20), make="L&T"),
        _row(pl_media, pg_item, '', 1, unit_price_override=pg_price, make=pg_vendor),
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


    # ── ENCON ITEMS ────────────────────────────────────────────────────────
    # Dual fuel: one physical dual-fuel burner — match pricelist naming
    # ("ENCON 5A" → "ENCON DUAL- 5A") and show both fuel flows in the ref.
    if is_dual:
        burner_desc = equipment["burner"]["model"].replace("ENCON ", "ENCON DUAL- ")
        burner_ref = (
            f'{f1_label}: {equipment["burner"]["input_nm3hr"]} Nm3/hr | '
            f'{f2_label}: {equipment2["burner"]["input_nm3hr"]} Nm3/hr'
        )
    else:
        burner_desc = equipment["burner"]["model"]
        burner_ref = f'GAS FLOW: {equipment["burner"]["input_nm3hr"]} Nm3/hr'
    rows += [
        _row("ENCON ITEMS", burner_desc,
             burner_ref,
             1, unit_price_override=equipment["burner"]["price"]),
        _row("ENCON ITEMS",
             "BEARING (24026)" if ladle_tons >= 50 else "BEARING (22222)",
             "", 2),
        _row("ENCON ITEMS", "PLUMMER BLOCK", "", 1,
             unit_price_override=params.get("plummer_block_kg", 300) * 170),
        _row("ENCON ITEMS", "SHAFT", "", 1,
             unit_price_override=params.get("shaft_kg", 350) * 120),
        _row("ENCON ITEMS", "FABRICATION/ STRUCTURE",
             f'{(ms_structure_kg_override or params["ms_structure_kg"])} kg',
             1, unit_price_override=(ms_structure_kg_override or params["ms_structure_kg"]) * get_price("FABRICATION RATE")),
        _row("ENCON ITEMS", "AIR-GAS PIPELINE",
             f'{pipeline_weight_kg:.0f} kg', 1,
             unit_price_override=pipeline_weight_kg * get_price("PIPELINE RATE")),
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
    cf_rolls = ceramic_rolls_override or params["ceramic_rolls"]
    rows.append(_row(
        "ENCON ITEMS", "CERAMIC FIBRE",
        f'{cf_rolls} Rolls @ Rs.{params.get("ceramic_rate", 0):,.0f}/roll',
        cf_rolls,
        unit_price_override=params.get("ceramic_rate", 0),
    ))

    # HEATING & PUMPING UNIT (oil-based fuels only).
    # Dual fuel: HPU may come from equipment2 if the oil fuel is fuel2.
    hpu = equipment.get("hpu") or (equipment2.get("hpu") if equipment2 else None)
    if hpu:
        rows.append(_row(
            "ENCON ITEMS", hpu.get("label", "Heating and Pumping Unit (HPU)"),
            f'{hpu["model"]} — {hpu["unit_kw"]} KW {hpu["variant"]}',
            1, unit_price_override=hpu["price"],
        ))

    # ── MISC ITEMS ─────────────────────────────────────────────────────────
    STATIC_SKIP = {"CONTROL PANEL"}
    if is_plc or is_plc_agr:
        STATIC_SKIP.update({"P.PID", "RATIO CONTROLLER"})
    if is_pid:
        STATIC_SKIP.add("TEMPERATURE TRANSMITTER")
    # Swivelling hoods use a geared drive — no hydraulic power pack needed.
    if hood_type == "swivel":
        STATIC_SKIP.add("HYDRAULIC POWER PACK & CYLINDER")
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

    # ── Tundish: replicate burner-line items by number of burners ──────────
    # Each burner has its own burner unit + ignition / UV / pilot / sensor set,
    # and its own pilot line. Main equipment (blower, gas train, control valve,
    # structure, panel, HPU, etc.) stays at qty 1.
    if num_burners > 1:
        burner_model = equipment["burner"]["model"]
        # Dual fuel renames the burner row to "ENCON DUAL- 5A" — cover both.
        burner_model_dual = burner_model.replace("ENCON ", "ENCON DUAL- ")
        per_burner_items = {
            burner_model,
            burner_model_dual,
            "Ignition Transformer",
            "Sequence Controller",
            "UV Sensor with Air Jacket",
            "BALL VALVE (Pilot Burner)",
            "BALL VALVE (UV LINE)",
            "FLEXIBLE HOSE (Pilot Burner)",
            "FLEXIBLE HOSE (UV LINE)",
            "ENCON-PB-LPG-10KW",
            "ENCON-PB NG 10 KW",
            "ENCON-PB LPG 100 KW",
            "ENCON-PB NG 100 KW",
            "ENCON PB COG 100 KW",
        }
        pilot_line_media = {
            f"{pl} PILOT LINE" for pl in ("LPG", "NG", "COG", "LNG", "RLNG")
        }
        mask = (
            df["ITEM NAME"].isin(per_burner_items)
            | df["MEDIA"].isin(pilot_line_media)
        )
        df.loc[mask, "QTY"]   = df.loc[mask, "QTY"]   * num_burners
        df.loc[mask, "TOTAL"] = df.loc[mask, "TOTAL"] * num_burners

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


def build_vlph_manual_df(
    equipment: dict,
    ladle_tons: float = 10.0,
    fuel1_type: str = "ng",
    pressure_gauge_vendor: str = "baumer",
    pilot_burner: str = "auto",
    pipeline_weight_kg: float = 1000.0,
    include_pilot: bool = True,
    pilot_line_fuel: str = "lpg",
    hood_type: str = "up_down",
) -> pd.DataFrame:
    """
    Manual / simplified VLPH BOM — matches the Lloyds manual costing format.
    Bought Out items are few; In-House items are grouped.
    LPG NG Pilot Line is expanded with individual items (same as automatic).
    include_pilot=False skips pilot burner, ignition transformer, UV sensor, pilot line.
    """

    pg_vendor = pressure_gauge_vendor.upper()
    # Use clean display name; vendor only shown in MAKE column.
    # Keep vendor-suffixed key for DB price lookup.
    pg_price = _get_price_fuzzy(f'PRESSURE GAUGE WITH TNV ({pg_vendor})')
    pg_item = "PRESSURE GAUGE WITH TNV"
    params = get_vlph_params(ladle_tons)

    rows = []

    # ── BOUGHT OUT ITEMS ──────────────────────────────────────────────────
    air_nb = max(125, equipment["air_duct"]["nb"])

    rows += [
        _row("COMB AIR", "COMPENSATOR", f'{air_nb} NB F150#', 1),
        _row("COMB AIR", pg_item, 'RANGE- 0-1600 mBAR', 1, unit_price_override=pg_price, make=pg_vendor),
        _row("COMB AIR", "BUTTERFLY VALVE",
             f'{equipment["butterfly_valve"]["nb"]} NB', 1,
             unit_price_override=equipment["butterfly_valve"]["price"],
             make=equipment["butterfly_valve"].get("make", "L&T")),
        _row("MISC ITEMS", "CONTROL PANEL", "", 1),
    ]

    # ── FUEL LINE (base items only — no AGR, orifice, DPT, control valve) ─
    f1_label = FUEL_NAMES.get(fuel1_type, fuel1_type.upper())
    rows += _fuel_line_rows(
        f1_label, fuel1_type, equipment,
        pressure_gauge_vendor=pressure_gauge_vendor,
        base_only=True,
    )

    # ── PILOT LINE (only if pilot burner is included) ──────────────
    pilot_media = f"{pilot_line_fuel.upper()} PILOT LINE"
    if include_pilot:
        rows += [
            _row(pilot_media, "BALL VALVE", "20 NB", 1,
                 unit_price_override=_get_cheapest_ball_valve(20), make="L&T"),
            _row(pilot_media, pg_item, '', 1, unit_price_override=pg_price, make=pg_vendor),
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
            "ENCON ITEMS", hpu.get("label", "Heating and Pumping Unit (HPU)"),
            f'{hpu["model"]} — {hpu["unit_kw"]} KW {hpu["variant"]}',
            1, unit_price_override=hpu["price"],
        ))

    rows += [
        _row("ENCON ITEMS", equipment["blower"]["model"],
             f'{equipment["blower"]["hp"]} HP, {equipment["blower"]["pressure"]} WC, '
             f'{equipment["blower"]["airflow_nm3hr"]} Nm3/hr',
             1, unit_price_override=equipment["blower"]["price_premium"]),
    ]
    if hood_type != "swivel":
        rows.insert(-1, _row("ENCON ITEMS", "HYDRAULIC POWER PACK & CYLINDER", "", 1))
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
