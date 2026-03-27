"""
Ladle preheater sizing lookup tables.
Source: Pricelist WorkBook 28-08-2025.xlsx — Vertical & Horizontal sheets.

MS Structure fabricated rate = Rs.151.2/kg (derived: amount/weight is
consistent at 151.2 across all VLPH and HLPH sizes in the pricebook).
"""

# Fabricated structure rate (material + machining + labour + paint)
# Verified: 1096 kg × 151.2 = 165,715 (10T VLPH) ✓
#           1300 kg × 151.2 = 196,560 (20T VLPH) ✓
#           1683 kg × 151.2 = 254,500 (60T HLPH) ✓
MS_STRUCTURE_RATE = 151.2  # Rs/kg

# --------------------------------------------------------------------------
# VLPH (Vertical Ladle Pre-Heater)
# Columns: ladle_tons_max, ms_kg, ceramic_rolls, hpu_kw,
#          pipeline_swirling_cost, control_panel_cost
# --------------------------------------------------------------------------
VLPH_SIZE_TABLE = [
    (10,  1096, 9,  3,  124929, 105000),
    (14,  1130, 8,  6,  139047, 105000),
    (15,  1215, 9,  6,  147000, 115500),
    (20,  1300, 10, 6,  158372, 115500),
    (30,  1354, 11, 12, 168000, 147000),
    (40,  1773, 13, 12, 178500, 147000),
    (60,  1800, 19, 20, 189000, 147000),
]

# --------------------------------------------------------------------------
# HLPH (Horizontal Ladle Pre-Heater)
# Columns: ladle_tons_max, ms_kg, ceramic_rolls, hpu_kw,
#          pipeline_cost, trolley_drive_cost, control_panel_cost
# --------------------------------------------------------------------------
HLPH_SIZE_TABLE = [
    (10,  1096, 9,  3,  46200,  124200, 126000),
    (15,  1236, 11, 6,  50400,  126900, 126000),
    (20,  1416, 13, 6,  115500, 129600, 126000),
    (30,  1450, 14, 12, 136500, 139500, 157500),
    (40,  1500, 15, 12, 163800, 144000, 157500),
    (60,  1683, 19, 20, 168000, 222600, 135000),
]


def get_vlph_params(ladle_tons: float) -> dict:
    """
    Return sizing parameters for a VLPH of given ladle capacity.
    Selects the first row where ladle_tons <= ladle_tons_max.
    Falls back to the largest size if capacity exceeds table range.
    """
    for max_t, ms_kg, cf_rolls, hpu_kw, pipeline, panel in VLPH_SIZE_TABLE:
        if ladle_tons <= max_t:
            return _vlph_dict(ladle_tons, ms_kg, cf_rolls, hpu_kw, pipeline, panel)
    # Beyond largest size — use last row
    max_t, ms_kg, cf_rolls, hpu_kw, pipeline, panel = VLPH_SIZE_TABLE[-1]
    return _vlph_dict(ladle_tons, ms_kg, cf_rolls, hpu_kw, pipeline, panel)


def _vlph_dict(ladle_tons, ms_kg, cf_rolls, hpu_kw, pipeline, panel):
    return {
        "ladle_tons":            ladle_tons,
        "ms_structure_kg":       ms_kg,
        "ms_structure_rate":     MS_STRUCTURE_RATE,
        "ms_structure_cost":     round(ms_kg * MS_STRUCTURE_RATE, 2),
        "ceramic_rolls":         cf_rolls,
        "hpu_kw":                hpu_kw,
        "pipeline_swirling_cost": pipeline,
        "control_panel_cost":    panel,
    }


def get_hlph_params(ladle_tons: float) -> dict:
    """
    Return sizing parameters for an HLPH of given ladle capacity.
    """
    for max_t, ms_kg, cf_rolls, hpu_kw, pipeline, trolley, panel in HLPH_SIZE_TABLE:
        if ladle_tons <= max_t:
            return _hlph_dict(ladle_tons, ms_kg, cf_rolls, hpu_kw, pipeline, trolley, panel)
    max_t, ms_kg, cf_rolls, hpu_kw, pipeline, trolley, panel = HLPH_SIZE_TABLE[-1]
    return _hlph_dict(ladle_tons, ms_kg, cf_rolls, hpu_kw, pipeline, trolley, panel)


def _hlph_dict(ladle_tons, ms_kg, cf_rolls, hpu_kw, pipeline, trolley, panel):
    return {
        "ladle_tons":          ladle_tons,
        "ms_structure_kg":     ms_kg,
        "ms_structure_rate":   MS_STRUCTURE_RATE,
        "ms_structure_cost":   round(ms_kg * MS_STRUCTURE_RATE, 2),
        "ceramic_rolls":       cf_rolls,
        "hpu_kw":              hpu_kw,
        "pipeline_cost":       pipeline,
        "trolley_drive_cost":  trolley,
        "control_panel_cost":  panel,
    }


if __name__ == "__main__":
    for t in [10, 15, 20, 30, 40, 60]:
        p = get_vlph_params(t)
        print(f"VLPH {t}T → MS={p['ms_structure_cost']:,.0f}  "
              f"CF={p['ceramic_rolls']}rolls  HPU={p['hpu_kw']}kW  "
              f"Pipeline={p['pipeline_swirling_cost']:,.0f}  Panel={p['control_panel_cost']:,.0f}")
