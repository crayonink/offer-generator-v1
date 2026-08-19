"""The combustion equipment list, from the workbook's 'Combustion' sheet.

Nine lines: the recuperator, the blowers, the burner sets, the gas train, the
mass flow control system, the doors, the pusher and the ejector.

Four of them stop being typed here:

  * the recuperator takes its cost from the recuperator calculation, as the
    sheet does (E5 reads Recuperator!F55);
  * the mass flow control likewise (E9 reads 'Mass Flow Control'!H61);
  * the burner count is the zone table's own total rather than a typed 25;
  * the blower count is the computed duty divided by one blower's rating and
    rounded up — 208.46 HP over 100 HP a blower is the 3 the sheet types.

The gas train carries the firing rate in its name, so that follows the duty
too. What is left typed is what nobody derives: the unit prices, the number
of pneumatic doors, and the pusher and ejector.
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field


@dataclass
class BRFCombustionInputs:
    blower_hp_each:      float = 100.0      # the "100HP" in the sheet's label
    blower_size_label:   str   = '40"'
    blower_price:        float = 1287000.0  # E6
    # E7 is written 69482/1.8 — a selling price divided back to cost.
    burner_list_price:   float = 69482.0
    burner_markup:       float = 1.8
    burner_label:        str   = "85 Litre/hr (5A)"
    gas_train_price:     float = 1103550.0  # E8
    door_count:          float = 3.0        # E10
    door_price:          float = 80000.0
    pusher_price:        float = 1500000.0  # E11
    ejector_price:       float = 1500000.0  # E12
    operator_seat_price: float = 200000.0   # the +200000 in E12


@dataclass
class BRFCombustionResults:
    # [item, qty, unit price, total, derived?]
    rows:        list = field(default_factory=list)
    total_price: float = 0.0    # F13
    blower_count: int = 0
    burner_count: int = 0


def calculate_combustion(inp=None, recup_cost=0.0, mass_flow_cost=0.0,
                         burner_count=0, blower_hp=0.0,
                         firing_rate_nm3hr=0.0) -> BRFCombustionResults:
    c = inp or BRFCombustionInputs()
    rows = []

    def add(item, qty, price, derived=False):
        rows.append([item, float(qty), round(float(price), 2),
                     round(float(qty) * float(price), 2), bool(derived)])

    # One blower per rating's worth of duty, rounded up.
    blowers = (math.ceil(blower_hp / c.blower_hp_each - 1e-9)
               if blower_hp and c.blower_hp_each else 0)
    burners = int(burner_count or 0)
    burner_cost = (c.burner_list_price / c.burner_markup) if c.burner_markup else 0.0

    add("Recuperator", 1, recup_cost, True)
    add(f"Blower {c.blower_hp_each:g}HP / {c.blower_size_label}", blowers,
        c.blower_price, True)
    add(f"Burner Set {c.burner_label}", burners, burner_cost, True)
    add(f"Gas Train ({firing_rate_nm3hr:,.0f} Nm³/hr)", 1,
        c.gas_train_price, True)
    add("Mass flow control", 1, mass_flow_cost, True)
    add("Pneumatically operated doors", c.door_count, c.door_price)
    add("Pusher", 1, c.pusher_price)
    add("Ejector + Operator seating arrangement", 1,
        c.ejector_price + c.operator_seat_price)

    return BRFCombustionResults(
        rows=rows,
        total_price=round(sum(r[3] for r in rows), 2),
        blower_count=blowers,
        burner_count=burners,
    )
