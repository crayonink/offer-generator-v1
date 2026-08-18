"""Mass flow control, off the zone line sizing.

Ported from the workbook's 'Mass Flow Control' sheet. Every fired zone gets an
air line and a gas line, and each line gets a motorised flow control valve, an
orifice plate, a DPT — and the air line a thermocouple as well. Then one set of
furnace-wide instruments and the panel.

The sheet types the line sizes in: 600 NB air, 80 NB gas, and so on. This
module takes them from calculations/brf.py's zone sizing instead, so a furnace
with different zones gets different valves rather than the 60 TPH job's.

The rates are the sheet's own, and they are a short ladder — four points for a
valve, four for an orifice, with a gap between 80 and 450 NB you could drive a
lorry through. A size that lands between two points is interpolated and the row
is marked estimated. That is a quotation stand-in, not a price list, and it is
worth replacing with real vendor rates before this leaves the office.
"""
from __future__ import annotations

from dataclasses import dataclass, field


# NB -> rupees, from the Mass Flow Control sheet.
VALVE_RATES = {65: 17000.0, 80: 17200.0, 450: 105300.0, 600: 190500.0}
ORIFICE_RATES = {80: 5500.0, 100: 5500.0, 500: 47000.0, 600: 61400.0}

# Flat-rate items, no size against them.
RATE_DPT = 45000.0
RATE_THERMOCOUPLE = 30000.0
RATE_RTD = 2000.0
RATE_DAMPER = 350000.0
RATE_SOLENOID = 46000.0        # one per burner, on the gas line
RATE_TT = 13000.0
RATE_PT = 45000.0
RATE_PLC = 700000.0
RATE_PANEL = 900000.0

RTD_COUNT = 4


def rate_for(ladder, nb):
    """A rate for this size, and whether it had to be estimated.

    The sheet's ladder is four points with a gap between 80 and 450 NB, so
    "the next size at or above" put a 100 NB gas valve at the 450 NB price —
    105,300 against the 17,200 an 80 NB one costs. Between two points the
    rate is interpolated; outside them it holds at the nearest end. Either
    way the row is marked estimated, because a four-point ladder is a
    quotation stand-in and not a price list.
    """
    if not nb:
        return 0.0, False
    sizes = sorted(ladder)
    if nb in ladder:
        return ladder[nb], False
    if nb < sizes[0]:
        return ladder[sizes[0]], True
    if nb > sizes[-1]:
        return ladder[sizes[-1]], True
    for lo, hi in zip(sizes, sizes[1:]):
        if lo < nb < hi:
            span = hi - lo
            frac = (nb - lo) / span if span else 0.0
            return round(ladder[lo] + frac * (ladder[hi] - ladder[lo]), 2), True
    return ladder[sizes[-1]], True


@dataclass
class BRFMassFlowResults:
    # [section, item, size (NB or ""), qty, unit price, total]
    rows:        list = field(default_factory=list)
    total_price: float = 0.0     # H61
    zone_count:  int = 0
    estimated_lines: int = 0     # rows priced off the ladder rather than on it


def calculate_mass_flow(zones, burner_count=0, zone_count=None) -> BRFMassFlowResults:
    """zones — the sizing chain's zone list, as objects or as dicts.

    zone_count trims it to the fired zones: the sizing table carries a
    standalone burner as a row of its own, and that one has no flow control
    loop of its own to bill.
    """
    rows = []

    def add(section, item, nb, qty, rate, est=False):
        rows.append([section, item, (f"{int(nb)} NB" if nb else ""),
                     float(qty), float(rate), round(float(qty) * float(rate), 2),
                     bool(est)])

    def get(z, key, default=0):
        return z.get(key, default) if isinstance(z, dict) else getattr(z, key, default)

    fired = list(zones or [])
    if zone_count is not None:
        fired = fired[:int(zone_count)]
    else:
        fired = [z for z in fired if get(z, "is_zone", True)]
    for z in fired:
        name = get(z, "name", "Zone")
        air = get(z, "air_line_nb", 0) or 0
        gas = get(z, "gas_line_nb", 0) or 0
        sec = f"{name} (Air Line)"
        r, e = rate_for(VALVE_RATES, air)
        add(sec, "Motorized Flow Control Valve", air, 1, r, e)
        r, e = rate_for(ORIFICE_RATES, air)
        add(sec, "Orifice Plate", air, 1, r, e)
        add(sec, "DPT", 0, 1, RATE_DPT)
        add(sec, "Thermocouple with TT — R type", 0, 1, RATE_THERMOCOUPLE)
        sec = f"{name} (Gas Line)"
        r, e = rate_for(VALVE_RATES, gas)
        add(sec, "Motorized Flow Control Valve", gas, 1, r, e)
        r, e = rate_for(ORIFICE_RATES, gas)
        add(sec, "Orifice Plate", gas, 1, r, e)
        add(sec, "DPT", 0, 1, RATE_DPT)

    # Furnace-wide instruments and the panel
    sec = "Furnace"
    add(sec, "RTD", 0, RTD_COUNT, RATE_RTD)
    add(sec, "DPT", 0, 1, RATE_DPT)
    add(sec, "Motorized Damper", 0, 1, RATE_DAMPER)
    add(sec, "Solenoid Valve for each gas line burner (65 NB)", 0,
        burner_count or 0, RATE_SOLENOID)
    add(sec, "TT in Gas Line", 0, 1, RATE_TT)
    add(sec, "PT in Gas Line", 0, 1, RATE_PT)
    add(sec, "TT in Air Line", 0, 1, RATE_TT)
    add(sec, "PT in Air Line", 0, 1, RATE_PT)
    add(sec, "PLC-S7 1500 with HMI", 0, 1, RATE_PLC)
    add(sec, "Control Panel", 0, 1, RATE_PANEL)

    return BRFMassFlowResults(
        rows=rows,
        total_price=round(sum(r[5] for r in rows), 2),
        zone_count=len(fired),
        estimated_lines=sum(1 for r in rows if r[6]),
    )
