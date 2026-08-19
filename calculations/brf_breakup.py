"""The commercial breakup, from the workbook's 'Breakup (TRL)' sheet.

The roll-up: every major item at its cost, a markup per line, and a selling
price. Nothing here is a new calculation — it reads what the other blocks have
already worked out, exactly as the sheet's own formulas do:

    C4  burner count      <- the zone table
    E4  burner unit       <- Combustion!E7
    E5  blower            <- Combustion!E6
    E6  gas train         <- Combustion!E8
    E7  recuperator       <- Recuperator!F55
    C13 CI skid weight    <- Furnace!G107
    E16 control panel     <- 'Mass Flow Control'!G59 + G60
    E17 mass flow         <- 'Mass Flow Control'!H61 less the panel
    E23 structure         <- 'Ref.+Str.'!N55
    E24 refractory        <- 'Ref.+Str.'!N43

The markup is per line and it matters: 1.8 on most of it, 1.5 on the blower,
the ejector and the pusher, 1.3 on the gas train and the refractory, 1.0 on
design. A single blended figure would misprice the job.

Two lines carry their total in the unit column rather than a rate — structure
and refractory, where the sheet writes F=E and leaves the tonnage in C as a
label. That is reproduced rather than tidied, because the tonnage is worth
seeing beside the money.
"""
from __future__ import annotations

from dataclasses import dataclass, field

# The groups the sheet summarises into, at the bottom.
G_COMBUSTION = "Combustion system with recuperator"
G_FURNACE = "Furnace with castings, steel structure and refractory"
G_AUTOMATION = "Automation & control"
G_MOVERS = "Ejector and pusher"
G_DESIGN = "Design & engineering"


@dataclass
class BRFBreakupInputs:
    # Markups, per line
    mk_burner: float = 1.8
    mk_blower: float = 1.5
    mk_gas_train: float = 1.3
    mk_recuperator: float = 1.8
    mk_ejector: float = 1.5
    mk_pusher: float = 1.5
    mk_casting: float = 1.8
    mk_cylinder: float = 1.8
    mk_automation: float = 1.8
    mk_design: float = 1.0
    mk_structure: float = 1.8
    mk_refractory: float = 1.3
    mk_pilot: float = 1.8
    # Rates and typed items the sheet carries
    ci_casting_rate: float = 150.0      # E10/E11/E12
    ci_skid_rate: float = 100.0         # E13
    door_plate_kg: float = 3000.0       # C11, a typed weight
    cylinder_price: float = 80000.0     # E14
    cylinder_count: float = 3.0         # C14 = 1 + 2
    design_price: float = 2500000.0     # E18
    design_qty: float = 0.0             # C18 — carried, not charged
    pilot_burner_price: float = 25000.0  # E20
    pilot_burner_qty: float = 0.0
    uv_sensor_price: float = 10000.0    # E21
    uv_sensor_qty: float = 0.0


@dataclass
class BRFBreakupResults:
    # [group, item, qty, uom, unit price, cost, markup, sell]
    rows:        list = field(default_factory=list)
    cost_total:  float = 0.0    # F29
    sell_total:  float = 0.0    # G29
    # [group, cost, sell] — the sheet's own summary block
    summary:     list = field(default_factory=list)


def _get(obj, key, default=0.0):
    if obj is None:
        return default
    if isinstance(obj, dict):
        return obj.get(key, default)
    return getattr(obj, key, default)


def _row_price(rows, needle):
    """A unit price out of the combustion table, by the start of its name."""
    for r in rows or []:
        if str(r[0]).lower().startswith(needle):
            return float(r[2])
    return 0.0


def build_breakup(combustion=None, recuperator=None, mass_flow=None,
                  gas_train=None, casting=None, priced=None,
                  inp=None) -> BRFBreakupResults:
    b = inp or BRFBreakupInputs()
    crows = _get(combustion, "rows", []) or []
    rows = []

    def add(group, item, qty, uom, unit, markup, cost=None):
        c = float(cost if cost is not None else qty * unit)
        rows.append([group, item, round(float(qty), 2), uom,
                     round(float(unit), 2), round(c, 2), markup,
                     round(c * markup, 2)])

    # ── Combustion equipment ────────────────────────────────────────
    burners = _get(combustion, "burner_count", 0)
    add(G_COMBUSTION, "Burner set", burners, "Set",
        _row_price(crows, "burner set"), b.mk_burner)

    blowers = _get(combustion, "blower_count", 0)
    model = _get(combustion, "blower_model", "") or ""
    add(G_COMBUSTION, f"Blower{' ' + model if model else ''}", blowers, "Nos",
        _row_price(crows, "blower"), b.mk_blower)

    trains = _get(gas_train, "train_count", 1) or 1
    add(G_COMBUSTION, "Gas train, assembled", trains, "Set",
        _get(gas_train, "per_train", _row_price(crows, "gas train")),
        b.mk_gas_train)

    add(G_COMBUSTION, "Recuperator", 1, "No.",
        _get(recuperator, "total_cost", 0.0), b.mk_recuperator)

    add(G_MOVERS, "Ejector + operator seating", 1, "Set",
        _row_price(crows, "ejector"), b.mk_ejector)
    add(G_MOVERS, "Pusher", 1, "Set", _row_price(crows, "pusher"), b.mk_pusher)

    # ── Castings, off the casting block ─────────────────────────────
    add(G_FURNACE, "CI casting for doors", _get(casting, "door_weight_kg", 0.0),
        "Kg", b.ci_casting_rate, b.mk_casting)
    add(G_FURNACE, "Inspection door, discharge door, plate door",
        b.door_plate_kg, "Kg", b.ci_casting_rate, b.mk_casting)
    add(G_FURNACE, "CI hanger", _get(casting, "hanger_weight_kg", 0.0), "Kg",
        b.ci_casting_rate, b.mk_casting)
    add(G_FURNACE, "CI skid for preheating zone",
        _get(casting, "skid_weight_kg", 0.0), "Kg", b.ci_skid_rate,
        b.mk_casting)
    add(G_FURNACE, "Pneumatic cylinder with door lifting arrangement",
        b.cylinder_count, "Set", b.cylinder_price, b.mk_cylinder)

    # ── Automation: the panel out of the mass flow list, and the rest ──
    mrows = _get(mass_flow, "rows", []) or []
    panel = sum(float(r[5]) for r in mrows
                if str(r[1]).lower().startswith(("plc", "control panel")))
    mf_total = _get(mass_flow, "total_price", 0.0)
    add(G_AUTOMATION, "Control panel with PLC", 1, "Set", panel, b.mk_automation)
    add(G_AUTOMATION, "Mass flow control", 1, "Set", mf_total - panel,
        b.mk_automation)

    add(G_DESIGN, "Design, engineering & purchase support", b.design_qty,
        "Set", b.design_price, b.mk_design)
    add(G_COMBUSTION, "Pilot burner with accessories", b.pilot_burner_qty,
        "Set", b.pilot_burner_price, b.mk_pilot)
    add(G_COMBUSTION, "UV sensor", b.uv_sensor_qty, "Set", b.uv_sensor_price,
        b.mk_pilot)

    # ── Structure and refractory: the whole cost sits in one line ────
    add(G_FURNACE, "Structure", _get(priced, "structure_kg", 0.0) / 1000.0,
        "Ton", _get(priced, "structure_cost", 0.0), b.mk_structure,
        cost=_get(priced, "structure_cost", 0.0))
    add(G_FURNACE, "Refractory", _get(priced, "refractory_kg", 0.0) / 1000.0,
        "Ton", _get(priced, "refractory_cost", 0.0), b.mk_refractory,
        cost=_get(priced, "refractory_cost", 0.0))

    cost_total = round(sum(r[5] for r in rows), 2)
    sell_total = round(sum(r[7] for r in rows), 2)

    summary = []
    for g in (G_DESIGN, G_FURNACE, G_COMBUSTION, G_AUTOMATION, G_MOVERS):
        c = round(sum(r[5] for r in rows if r[0] == g), 2)
        s = round(sum(r[7] for r in rows if r[0] == g), 2)
        summary.append([g, c, s])

    return BRFBreakupResults(rows=rows, cost_total=cost_total,
                             sell_total=sell_total, summary=summary)
