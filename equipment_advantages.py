# Per-equipment "Advantages" content for the stand-alone equipment offers.
# Rendered dynamically into the offer templates (adv_* placeholders); the
# heading + bullets + optional sub-blocks + closing all flow from here.

EQUIPMENT_ADVANTAGES = {
    "blower": {
        "title": "ADVANTAGES OF ENCON BLOWERS",
        "items": [
            "Only CRC Sheets used",
            "Only Steel from Standard companies such as SAIL used",
            "Manoeuvrability available (Blower can be positioned in any direction).",
            "Only Standard Motors used (ABB, Crompton, etc.)",
            "Heavy Duty as thicker sheets used to give it strength for a longer life",
            "Designing superiority (which also adds to additional cost) ensures better performance",
        ],
        "sub": [],
    },
    "burner_film": {
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
        "sub": [],
    },
    "burner_gas": {
        "title": "ADVANTAGES OF ENCON GAS BURNERS",
        "items": [
            "High Thermal Efficiency",
            "Low NOx Emissions",
            "Stable Flame Across Wide Operating Range",
            "High Turndown Ratio (Up to 10:1)",
            "Suitable for Natural Gas, LPG, COG, Producer Gas",
            "Precise Air-Fuel Ratio Control",
            "Fast Heating Response",
            "Low Maintenance Requirement",
            "Compatible with PLC/SCADA Automation",
            "Energy Saving Operation",
        ],
        "sub": [],
    },
    "burner_dual": {
        "title": "ADVANTAGES OF ENCON DUAL FUEL BURNERS",
        "items": [
            "Operates on Gas + Liquid Fuel (LDO/HSD/FO)",
            "Automatic Fuel Changeover Facility",
            "High Combustion Efficiency",
            "Reliable Operation During Fuel Supply Interruptions",
            "Wide Turndown Ratio",
            "Stable Flame in Both Fuel Modes",
            "Reduced Operating Cost",
            "Flexible Process Operation",
            "Easy Maintenance",
            "Suitable for Furnace, Ladle & Tundish Heating Applications",
        ],
        "sub": [],
    },
    "burner_hv": {
        "title": "ADVANTAGES OF ENCON HIGH VELOCITY BURNERS",
        "items": [
            "Uniform Temperature Distribution",
            "High Heat Transfer Rate",
            "Reduced Heating Cycle Time",
            "Excellent Temperature Uniformity (±5°C to ±10°C)",
            "Lower Fuel Consumption",
            "High Flame Momentum",
            "Reduced Oxidation and Scale Formation",
            "Suitable for Heat Treatment Furnaces",
            "Supports Preheated Combustion Air",
            "High Thermal Efficiency and Productivity",
        ],
        "sub": [],
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

CLOSING_LINES = []   # removed per request (no "extra money invested…" / "Do we need to say any more?")

# drive_product / mode / explicit kind -> advantages key. Specific burner
# kinds (burner_gas/dual/hv/film) fall through to themselves.
_KIND_MAP = {"blower": "blower", "burner": "burner_film",
             "hpu": "pumping", "pu": "pumping", "pumping": "pumping"}


# Standard ENCON equipment Terms & Conditions — used as defaults when the
# offer form leaves a field blank (so the T&C annexure is never empty).
STANDARD_TNC = {
    "tnc_prices":             "Ex-works Bhagola, Dist. Palwal, Haryana, INDIA.",
    "tnc_delivery":           "4–6 weeks from the date of confirmed order along with the relevant advance.",
    "tnc_gst":                "GST @ 18% extra.",
    "tnc_hsn_code":           "",
    "tnc_pan_gst":            "",
    "tnc_payment_terms":      "30% advance along with the PO; balance 70% against Proforma Invoice prior to despatch.",
    "tnc_packing_forwarding": "4% of equipment value towards packing + 2% towards forwarding charges.",
    "tnc_freight":            "At actual, on to-pay basis.",
    "tnc_transit_insurance":  "To be arranged by the client.",
    "tnc_validity":           "30 days.",
    "tnc_inspection":         "If required, materials can be inspected at our works before despatch at your cost with prior intimation.",
    "tnc_guarantee":          "12 months from the date of supply against manufacturing defects.",
}


def tnc_value(key: str, provided) -> str:
    """The provided T&C value, or the standard ENCON default if blank."""
    return provided if (provided or "").strip() else STANDARD_TNC.get(key, "")


# Regenerative-burner offers carry a different Terms & Conditions layout (a
# 9-row site/erection-oriented table, not the 12-field equipment set). These
# defaults match the Regen_Offer_Template.docx wording and fill in whenever the
# regen form leaves a field blank. Multi-line values use "\n" (docxtpl renders
# each as a line break).
REGEN_STANDARD_TNC = {
    "tnc_regen_material_custody":
        "All material supplied by us will be received by your store for safe "
        "custody and storage. The material will be issued to us as and when "
        "required for installation.",
    "tnc_regen_site_facilities":
        "Electricity and water to be provided free of charge for the erection work.\n"
        "Crane facilities available in your works to be provided free of charge.",
    "tnc_regen_delivery":
        "Drawings for the furnace will be prepared and submitted within 5 weeks "
        "from the date of your confirmed order for the design work.\n"
        "After approval of the drawings, items in our scope of supply will be "
        "delivered within 15 weeks from the date of drawing approval.\n"
        "The erection schedule will be mutually worked out based on your scope "
        "of erection work.\n"
        "We shall not be responsible for any delay due to non-availability of "
        "fronts, delay in issue of work permits, safety clearances and other "
        "clearances that fall in your purview.",
    "tnc_regen_transportation":
        "From our factory to your work site — extra at actuals and chargeable to you.",
    "tnc_regen_gst":
        "18% extra.",
    "tnc_regen_hsn":
        "84541000",
    "tnc_regen_pan_gst":
        "PAN: AAACE0327M  |  GST: 06AAACE0327M1ZV",
    "tnc_regen_payment":
        "30% along with the purchase order.\n"
        "70% against inspection of material at our works.",
    "tnc_regen_packing":
        "4% of the order value, charged extra.",
    "tnc_regen_warranty":
        "18 months from the date of material supplied (any manufacturing defect) "
        "or 12 months from date of commissioning, whichever is earlier.",
}

REGEN_TNC_KEYS = list(REGEN_STANDARD_TNC.keys())


def regen_tnc_value(key: str, provided) -> str:
    """The provided regen T&C value, or the regen default if blank."""
    return provided if (provided or "").strip() else REGEN_STANDARD_TNC.get(key, "")


def build_advantages_ctx(kind: str) -> dict:
    """adv_* context for the offer template, for the given equipment kind."""
    k = (kind or "").lower()
    a = EQUIPMENT_ADVANTAGES.get(_KIND_MAP.get(k, k), {})
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
