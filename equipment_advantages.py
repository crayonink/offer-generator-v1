# Per-equipment "Advantages" content for the stand-alone equipment offers.
# Rendered dynamically into the offer templates (adv_* placeholders); the
# heading + bullets + optional sub-blocks + closing all flow from here.

EQUIPMENT_ADVANTAGES = {
    "blower": {
        "title": "ADVANTAGES OF ENCON BLOWERS",
        "items": [
            "Only CRC Sheets used",
            "Only Steel from Standard companies such as SAIL used",
            "Maneuverability available (Blower can be positioned in any direction). "
            "This requires a DOUBLE Sheet Blower instead of SINGLE Sheet",
            "Only Standard Motors used (ABB, Crompton, etc.)",
            "Heavy Duty as thicker sheets used giving it strength for a longer life",
            "Designing superiority (which also adds to additional cost) ensures better performance",
        ],
        "sub": [],
    },
    "burner": {
        "title": "ADVANTAGES OF IIP-ENCON “FILM” BURNERS",
        "items": [
            "Huge savings in fuel cost (5-15%)",
            "Multi-fuel application (HSD, LDO, FO, LSHS)",
            "Higher turn down ratio (7:1)",
            "Eliminates Oil Dripping",
            "Stable Flame",
            "Simpler to Maintain",
            "Can handle Preheated air",
            "Reduces Heating time",
            "Eliminates Choking",
            "Reduces Sulphur corrosion",
        ],
        "sub": [
            {"title": "Disadvantage Of IIP-ENCON “FILM” Burner",
             "items": ["Higher Price"]},
            {"title": "Other advantages",
             "items": [
                 "Graded Castings used",
                 "Better air control through air adjuster",
                 "Nozzles fully of SS",
                 "Swirlers / Lock nuts fully SS (so cleaning does not wear them off leading to longer life)",
                 "Higher efficiency for a longer duration (other burners start consuming more oil in 3 months)",
             ]},
        ],
    },
    "pumping": {  # HPU + PU
        "title": "ADVANTAGES OF ENCON HEATING & PUMPING UNITS",
        "items": [
            "Only Standard Motors used (ABB, Crompton, etc.)",
            "Only Standard Pumps used (Tushaco, Apex, etc.)",
            "Only Standard Valves used (Audco or equivalent)",
            "Only ISI pipes used",
            "Heavy duty base for sturdiness",
            "Heavy duty Oil Tank for longer life",
        ],
        "sub": [],
    },
}

CLOSING_LINES = [
    "The extra money invested in our ENCON Equipment often gets paid off in just 2 weeks.",
    "Do we need to say any more?",
]

# drive_product / mode -> advantages key
_KIND_MAP = {"blower": "blower", "burner": "burner",
             "hpu": "pumping", "pu": "pumping", "pumping": "pumping"}


def build_advantages_ctx(kind: str) -> dict:
    """adv_* context for the offer template, for the given equipment kind."""
    a = EQUIPMENT_ADVANTAGES.get(_KIND_MAP.get((kind or "").lower(), ""), {})
    subs = a.get("sub", [])
    return {
        "adv_title":       a.get("title", ""),
        "adv_items":       [{"item": x} for x in a.get("items", [])],
        "adv_sub1_title":  subs[0]["title"] if len(subs) > 0 else "",
        "adv_sub1_items":  [{"item": x} for x in subs[0]["items"]] if len(subs) > 0 else [],
        "adv_sub2_title":  subs[1]["title"] if len(subs) > 1 else "",
        "adv_sub2_items":  [{"item": x} for x in subs[1]["items"]] if len(subs) > 1 else [],
        "adv_closing":     [{"item": x} for x in CLOSING_LINES] if a else [],
    }
