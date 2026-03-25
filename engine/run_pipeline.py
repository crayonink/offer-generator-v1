# engine/run_pipeline.py

"""
End-to-end execution pipeline.

Flow:
Input → Selection → BOM → Total Cost
"""

from bom.selectors.selection_engine import select_equipment, SystemType
from bom.regen_bom_builder import build_regen_bom_df
from bom.vlph_builder import build_vlph_120t_df

def run_pipeline(
    *,
    system_type: SystemType,
    capacity_kw: float,
    ng_flow_nm3hr: float,
    air_flow_nm3hr: float,
):
    """
    Runs full costing pipeline.

    Returns:
    {
        equipment: dict,
        bom: DataFrame,
        total_cost: float
    }
    """

    # -------------------------------------------------
    # 1️⃣ SELECTION
    # -------------------------------------------------
    equipment = select_equipment(
        system_type=system_type,
        capacity_kw=capacity_kw,
        ng_flow_nm3hr=ng_flow_nm3hr,
        air_flow_nm3hr=air_flow_nm3hr,
    )

    # -------------------------------------------------
    # 2️⃣ BUILD BOM
    # -------------------------------------------------
    if system_type == SystemType.REGEN:
        bom_df = build_regen_bom_df(equipment)

    elif system_type == SystemType.VLPH:
        bom_df = build_vlph_bom_df(equipment)

    else:
        raise ValueError(f"Unsupported system type: {system_type}")

    # -------------------------------------------------
    # 3️⃣ TOTAL COST
    # -------------------------------------------------
    # Ignore summary rows if present
    valid_rows = bom_df[pd.to_numeric(bom_df["TOTAL"], errors="coerce").notnull()]
    total_cost = valid_rows["TOTAL"].sum()

    # -------------------------------------------------
    # RETURN
    # -------------------------------------------------
    return {
        "equipment": equipment,
        "bom": bom_df,
        "total_cost": total_cost,
    }