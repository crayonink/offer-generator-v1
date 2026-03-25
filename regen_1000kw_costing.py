"""
REGEN Burner 1000 KW - Costing Generator
Generates a detailed cost sheet from user inputs, matching the Excel template logic.
"""

import math

# ---------------------------------------------
#  BOM Price Lists (from Excel BOM sheet)
# ---------------------------------------------

PILOT_BURNER_PRICE         = 10000
BURNER_CONTROLLER_PRICE    = 3600
IGNITION_TRANSFORMER_PRICE = 3300
UV_SENSOR_PRICE            = 5500
PILOT_REGULATOR_PRICE      = 4400
PILOT_SOLENOID_PRICE       = 4300
FLEXIBLE_HOSE_15_PRICE     = 1500
BALL_VALVE_15_PRICE        = 1400
PRESSURE_GAUGE_500_PRICE   = 3000
PRESSURE_GAUGE_1000_PRICE  = 4000
BALL_VALVE_UV_PRICE        = 1400
FLEX_HOSE_UV_PRICE         = 1500
BALL_VALVE_PILOT_PRICE     = 1400
FLEX_HOSE_PILOT_PRICE      = 1500
MANUAL_DAMPER_PRICE        = 40000
DPT_PRICE                  = 45000
THERMOCOUPLE_K_PRICE       = 5000
THERMOCOUPLE_R_PRICE       = 25000

SOLENOID_VALVE_PRICES = {32: 13700, 40: 14720, 50: 17900, 65: 43000, 80: 44000, 100: 76000}
BALL_VALVE_PRICES     = {32: 4925, 40: 5100, 50: 7200, 65: 13400, 80: 17000, 100: 26600}
FLEX_HOSE_PRICES      = {32: 1750, 40: 2000, 50: 3000, 65: 4200, 80: 6900, 100: 7650}
THREE_WAY_VALVE_PRICES = {32: 0, 40: 0, 50: 0, 65: 0, 80: 0, 100: 0}

SHUT_OFF_AIR_PRICES   = {125: 50050, 200: 80000, 250: 125000, 300: 148000, 350: 177000, 400: 227500}
MANUAL_BF_AIR_PRICES  = {125: 12498, 200: 31178, 250: 38378, 300: 48055, 350: 61700, 400: 83750}
SHUT_OFF_FLUE_PRICES  = {200: 80000, 250: 125000, 300: 148000, 350: 177000, 400: 227500,
                          450: 361020, 500: 453470, 600: 838150, 700: 1048800}

AIR_CONTROL_VALVE_PRICES = {
    100: 110450, 150: 125600, 200: 144000, 250: 189540, 300: 213240,
    350: 242250, 400: 258503, 450: 327583, 500: 396680, 600: 561425,
    650: 625771, 700: 804856
}
AIR_FLOW_METER_PRICES = {
    125: 54000, 200: 57000, 250: 58000, 300: 60000, 350: 64000,
    400: 70500, 450: 80500, 500: 90500, 600: 100500, 650: 110500, 700: 120500
}
GAS_CONTROL_VALVE_PRICES = {
    25: 83000, 32: 83000, 40: 83000, 50: 96960, 65: 97810, 80: 101900,
    100: 110450, 150: 125600, 200: 144000, 250: 189540, 300: 213240,
    350: 242250, 400: 261000
}
GAS_FLOW_METER_PRICES = {
    32: 48000, 40: 49000, 50: 49700, 65: 50000, 80: 51000, 100: 52000,
    150: 54000, 200: 57000, 250: 58000, 300: 60000, 350: 64000, 400: 70500
}
PNEUMATIC_DAMPER_PRICES = {
    200: 80000, 250: 125000, 300: 148000, 350: 177000, 400: 350000,
    450: 350000, 500: 350000, 600: 350000
}

# Blower prices (HP -> price without motor, with 8% markup applied)
BLOWER_BASE_PRICES = {
    5: 56500, 7.5: 60500, 10: 76000, 15: 87000, 20: 91000,
    25: 111000, 30: 131000, 40: 151500, 50: 175000, 60: 198000
}

CONTROL_PANEL_PRICES = {500: 300000, 1000: 300000, 1500: 450000, 2000: 450000,
                         2500: 600000, 3000: 600000, 4500: 600000, 6000: 700000}
PLC_PRICE = 300000

NG_GAS_TRAIN_PRICES = {
    500: 88500, 1000: 110000, 1500: 139300, 2000: 144100,
    2500: 224000, 3000: 224000, 4500: 295200, 6000: 383000
}

# Burner with regenerator cost (from Burner Sizing sheet, row 21 = 1000 KW)
BURNER_REGEN_COST_1000KW = 162998.63


def closest_key(d, val):
    """Return the closest key in dict d to val."""
    return min(d.keys(), key=lambda k: abs(k - val))


def roundup(x, decimals=0):
    factor = 10 ** decimals
    return math.ceil(x / factor) * factor


# ---------------------------------------------
#  Pipe Size lookup for 1000 KW (from Burner Pipe Size sheet)
# ---------------------------------------------
# Air DN=200, Gas DN=40, Flue DN=250 for 1000 KW Natural Gas

PIPE_SIZES_1000KW = {
    "air_dn": 200,
    "gas_dn": 40,
    "flue_dn": 250,
}


# ---------------------------------------------
#  Main costing function
# ---------------------------------------------

def generate_costing(
    selling_price_multiplier: float = 1.8,
    pipeline_cost_extra: float = 0,
    blower_hp: float = 10,
    thermocouple_type: str = "K",   # "K" or "R"
    num_pairs: int = 2,              # number of burner pairs (default 2)
):
    """
    Generate the 1000 KW REGEN Burner costing.

    Parameters
    ----------
    selling_price_multiplier : markup factor (default 1.8)
    pipeline_cost_extra      : extra pipeline cost (default 0)
    blower_hp                : combustion blower HP (default 10)
    thermocouple_type        : 'K' (standard, cheaper) or 'R' (high temp, costlier)
    num_pairs                : number of burner pairs (default 2)
    """

    SPM = selling_price_multiplier
    n   = num_pairs           # quantity for most items
    nh  = n // 2              # half-quantity (temperature control system)

    thermo_price = THERMOCOUPLE_K_PRICE if thermocouple_type.upper() == "K" else THERMOCOUPLE_R_PRICE

    # Pipe sizes for 1000 KW
    air_dn  = PIPE_SIZES_1000KW["air_dn"]   # 200
    gas_dn  = PIPE_SIZES_1000KW["gas_dn"]   # 40
    flue_dn = PIPE_SIZES_1000KW["flue_dn"]  # 250

    # Blower price
    blower_base   = BLOWER_BASE_PRICES.get(blower_hp, BLOWER_BASE_PRICES[10])
    blower_price  = roundup(blower_base * 1.08, -4)   # 8% markup, round to nearest 10000

    items = []

    def add(description, size, qty, cost_unit):
        total_cost    = qty * cost_unit
        sell_unit     = cost_unit * SPM
        total_sell    = sell_unit * qty
        items.append({
            "description": description,
            "size":        size,
            "qty":         qty,
            "cost_unit":   cost_unit,
            "total_cost":  total_cost,
            "sell_unit":   sell_unit,
            "total_sell":  total_sell,
        })

    # -- Section 1: Burner & pilot ------------------------------------------
    add("Burner with Regenerator (1000 KW)",  "1000 KW", n,    BURNER_REGEN_COST_1000KW)
    add("Pilot Burner",                        "7 KW",    n,    PILOT_BURNER_PRICE)
    add("Burner Controller",                   "-",       n,    BURNER_CONTROLLER_PRICE)
    add("Ignition Transformer",                "-",       n,    IGNITION_TRANSFORMER_PRICE)
    add("UV Sensor",                           "-",       n,    UV_SENSOR_PRICE)

    # -- Section 2: Gas Line - Pilot ----------------------------------------
    add("Pilot Regulator",                     "NB15",    n,    PILOT_REGULATOR_PRICE)
    add("Pilot Solenoid Valve",                "NB15",    n,    PILOT_SOLENOID_PRICE)
    add("Flexible Hose (Pilot, 1000mm)",       "NB15",    n,    FLEXIBLE_HOSE_15_PRICE)
    add("Ball Valve (Pilot)",                  "NB15",    n,    BALL_VALVE_15_PRICE)
    add("Pressure Gauge 0-500mm (Gas)",        "-",       n,    PRESSURE_GAUGE_500_PRICE)

    # -- Section 3: Gas Line - Burner ---------------------------------------
    sv_price   = SOLENOID_VALVE_PRICES.get(gas_dn, SOLENOID_VALVE_PRICES[closest_key(SOLENOID_VALVE_PRICES, gas_dn)])
    bv_price   = BALL_VALVE_PRICES.get(gas_dn,     BALL_VALVE_PRICES[closest_key(BALL_VALVE_PRICES, gas_dn)])
    fh_price   = FLEX_HOSE_PRICES.get(gas_dn,      FLEX_HOSE_PRICES[closest_key(FLEX_HOSE_PRICES, gas_dn)])
    twv_price  = THREE_WAY_VALVE_PRICES.get(gas_dn, 0)

    add(f"Solenoid Valve (Gas Burner NB{gas_dn})", f"NB{gas_dn}", n,      sv_price)
    add(f"Ball Valve (Gas Burner NB{gas_dn})",     f"NB{gas_dn}", n * 5,  bv_price)
    add(f"Flexible Hose Pipe (Gas, 1000mm NB{gas_dn})", f"NB{gas_dn}", n * 5, fh_price)
    add(f"3-Way Valve NB{gas_dn}",                f"NB{gas_dn}", n,      twv_price)
    add("Pressure Gauge 0-500mm (Burner)",         "-",           n,      PRESSURE_GAUGE_500_PRICE)

    # -- Section 4: Air Line - Pilot, UV & UV Cooling -----------------------
    add("Ball Valve UV",                       "NB15",    n * 4,  BALL_VALVE_UV_PRICE)
    add("Flexible Hose UV (1000mm)",           "NB15",    n * 2,  FLEX_HOSE_UV_PRICE)
    add("Ball Valve Pilot",                    "NB15",    n * 2,  BALL_VALVE_PILOT_PRICE)
    add("Flexible Hose Pilot (1000mm)",        "NB15",    n * 2,  FLEX_HOSE_PILOT_PRICE)

    # -- Section 5: Air Line - Burner (shut-off valves) ---------------------
    soa_price  = SHUT_OFF_AIR_PRICES.get(air_dn,  SHUT_OFF_AIR_PRICES[closest_key(SHUT_OFF_AIR_PRICES, air_dn)])
    mbfa_price = MANUAL_BF_AIR_PRICES.get(air_dn, MANUAL_BF_AIR_PRICES[closest_key(MANUAL_BF_AIR_PRICES, air_dn)])
    sof_price  = SHUT_OFF_FLUE_PRICES.get(flue_dn, SHUT_OFF_FLUE_PRICES[closest_key(SHUT_OFF_FLUE_PRICES, flue_dn)])

    add(f"Shut-Off Valve Air (NB{air_dn})",         f"NB{air_dn}", n,     soa_price)
    add(f"Manual Butterfly Valve Air (NB{air_dn})", f"NB{air_dn}", n,     mbfa_price)
    add("Pressure Gauge 0-1000mm (Air)",             "-",           n,     PRESSURE_GAUGE_1000_PRICE)
    add(f"Shut-Off Valve Flue Gas (NB{flue_dn})",   f"NB{flue_dn}", n,    sof_price)
    add(f"Thermocouple with TT-{thermocouple_type}", "-",           n * 2, thermo_price)

    # -- Section 6: Temperature Control System -----------------------------
    acv_price = AIR_CONTROL_VALVE_PRICES.get(air_dn,  AIR_CONTROL_VALVE_PRICES[closest_key(AIR_CONTROL_VALVE_PRICES, air_dn)])
    afm_price = AIR_FLOW_METER_PRICES.get(air_dn,     AIR_FLOW_METER_PRICES[closest_key(AIR_FLOW_METER_PRICES, air_dn)])
    gcv_price = GAS_CONTROL_VALVE_PRICES.get(gas_dn,  GAS_CONTROL_VALVE_PRICES[closest_key(GAS_CONTROL_VALVE_PRICES, gas_dn)])
    gfm_price = GAS_FLOW_METER_PRICES.get(gas_dn,     GAS_FLOW_METER_PRICES[closest_key(GAS_FLOW_METER_PRICES, gas_dn)])

    add(f"Air Control Valve (NB{air_dn})",      f"NB{air_dn}", nh, acv_price)
    add(f"Air Flow Meter DPT (NB{air_dn})",     f"NB{air_dn}", nh, afm_price)
    add(f"Gas Control Valve (NB{gas_dn})",      f"NB{gas_dn}", nh, gcv_price)
    add(f"Gas Flow Meter DPT (NB{gas_dn})",     f"NB{gas_dn}", nh, gfm_price)
    add(f"Thermocouple with TT-R (TCS)",        "-",           nh, THERMOCOUPLE_R_PRICE)
    add("DPT (flow, pressure & temperature)",   "1 unit",       1,  DPT_PRICE)

    pd_price = PNEUMATIC_DAMPER_PRICES.get(flue_dn, PNEUMATIC_DAMPER_PRICES[closest_key(PNEUMATIC_DAMPER_PRICES, flue_dn)])
    add(f"Pneumatic Damper (NB{flue_dn})",      f"NB{flue_dn}", nh, pd_price)
    add("Manual Damper",                         "1 unit",        1,  MANUAL_DAMPER_PRICE)

    # -- Section 7: Blower --------------------------------------------------
    add(f"Combustion Blower ({blower_hp}HP/40\")", f"{blower_hp}HP", n, blower_price)

    # -- Section 8: Controls ------------------------------------------------
    add("PLC with HMI",    "-", 1, PLC_PRICE)
    cp_price = CONTROL_PANEL_PRICES.get(1000, 300000)
    add("Control Panel",   "-", 1, cp_price)

    # -- Section 9: Gas Train -----------------------------------------------
    gt_price = NG_GAS_TRAIN_PRICES.get(1000, 110000)
    add("NG - Gas Train (100 Nm3/hr)", "-", 1, gt_price)

    # -- Totals -------------------------------------------------------------
    total_cost  = sum(i["total_cost"] for i in items)
    total_sell  = sum(i["total_sell"] for i in items)
    grand_total = roundup(total_sell + pipeline_cost_extra, -4)

    return items, total_cost, total_sell, grand_total


# ---------------------------------------------
#  Display
# ---------------------------------------------

def print_costing(items, total_cost, total_sell, grand_total, params):
    SEP  = "-" * 110
    SEP2 = "=" * 110

    print(SEP2)
    print(f"  REGEN BURNER 1000 KW  -  COST OF GOODS (COG) REPORT")
    print(f"  Selling Price Multiplier : {params['selling_price_multiplier']}")
    print(f"  Number of Burner Pairs   : {params['num_pairs']}")
    print(f"  Combustion Blower HP     : {params['blower_hp']} HP")
    print(f"  Thermocouple Type        : {params['thermocouple_type']}")
    print(f"  Pipeline Cost Extra      : Rs.{params['pipeline_cost_extra']:,.0f}")
    print(SEP2)

    header = f"{'#':<4} {'Description':<48} {'Size':<10} {'Qty':>5}  {'Cost/Unit':>12}  {'Total Cost':>14}  {'Sell/Unit':>12}  {'Total Sell':>14}"
    print(header)
    print(SEP)

    for i, item in enumerate(items, 1):
        desc  = item["description"][:47]
        size  = str(item["size"])[:9]
        qty   = item["qty"]
        cu    = item["cost_unit"]
        tc    = item["total_cost"]
        su    = item["sell_unit"]
        ts    = item["total_sell"]
        print(f"{i:<4} {desc:<48} {size:<10} {qty:>5}  {cu:>12,.0f}  {tc:>14,.0f}  {su:>12,.0f}  {ts:>14,.0f}")

    print(SEP)
    print(f"{'SUBTOTAL':<73}  {total_cost:>14,.0f}  {'':>12}  {total_sell:>14,.0f}")
    print(f"{'PIPELINE COST EXTRA':<73}  {'':>14}  {'':>12}  {params['pipeline_cost_extra']:>14,.0f}")
    print(SEP2)
    print(f"{'GRAND TOTAL (Selling Price, rounded)':<73}  {'':>14}  {'':>12}  {grand_total:>14,.0f}")
    print(SEP2)
    print(f"\n  Rs. {grand_total:,.0f}  (approx. Rs. {grand_total/100000:.1f} Lakhs)")
    print()


# ---------------------------------------------
#  Interactive CLI
# ---------------------------------------------

def prompt_float(msg, default):
    raw = input(f"{msg} [{default}]: ").strip()
    return float(raw) if raw else default

def prompt_int(msg, default):
    raw = input(f"{msg} [{default}]: ").strip()
    return int(raw) if raw else default

def prompt_str(msg, default, choices=None):
    if choices:
        msg = f"{msg} ({'/'.join(choices)})"
    raw = input(f"{msg} [{default}]: ").strip()
    return raw.upper() if raw else default


def main():
    print("\n" + "=" * 60)
    print("  REGEN BURNER 1000 KW - Costing Generator")
    print("  Press Enter to accept default values shown in [brackets]")
    print("=" * 60 + "\n")

    spm           = prompt_float("Selling price multiplier",  1.8)
    num_pairs     = prompt_int  ("Number of burner pairs",    2)
    blower_hp     = prompt_float("Combustion blower HP",      10)
    thermo_type   = prompt_str  ("Thermocouple type",         "K", ["K", "R"])
    pipeline_cost = prompt_float("Pipeline cost extra (Rs.)",   0)

    params = {
        "selling_price_multiplier": spm,
        "num_pairs":                num_pairs,
        "blower_hp":                blower_hp,
        "thermocouple_type":        thermo_type,
        "pipeline_cost_extra":      pipeline_cost,
    }

    items, total_cost, total_sell, grand_total = generate_costing(**params)
    print()
    print_costing(items, total_cost, total_sell, grand_total, params)


if __name__ == "__main__":
    main()
