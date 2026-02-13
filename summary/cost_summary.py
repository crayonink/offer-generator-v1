# summary/cost_summary.py
"""
Cost Summary builder
Pure calculation + DataFrame assembly
"""

import pandas as pd
from config import MARKUP, USD_RATE


def build_cost_summary_df(
    *,
    bought_out_cost: float,
    bought_out_sell: float,
    inhouse_cost: float,
    inhouse_sell: float,
    item_description: str = "Vertical Ladle Preheater",
    qty_per_set: int = 1,
) -> pd.DataFrame:
    """
    Builds the Cost Summary sheet (single row).

    All inputs are assumed to be FINAL values coming
    from dynamic BOM aggregation.

    Excel is OUTPUT ONLY.
    """

    # -----------------------------
    # Normalize inputs (safety)
    # -----------------------------
    bought_out_cost = float(bought_out_cost)
    bought_out_sell = float(bought_out_sell)
    inhouse_cost = float(inhouse_cost)
    inhouse_sell = float(inhouse_sell)

    # -----------------------------
    # Unit prices
    # -----------------------------
    unit_cost_price = bought_out_cost + inhouse_cost
    unit_sell_price = bought_out_sell + inhouse_sell

    # -----------------------------
    # Add-ons (commercial logic)
    # -----------------------------
    designing_10 = unit_sell_price * 0.10
    negotiation_10 = unit_sell_price * 0.10

    # -----------------------------
    # Final pricing
    # -----------------------------
    total_price = unit_sell_price + designing_10 + negotiation_10
    usd_price = total_price / USD_RATE

    # -----------------------------
    # Assemble DataFrame
    # -----------------------------
    cost_summary_df = pd.DataFrame(
        [[
            1,
            item_description,
            round(bought_out_cost, 2),
            round(bought_out_sell, 2),
            round(inhouse_cost, 2),
            round(inhouse_sell, 2),
            round(unit_cost_price, 2),
            round(unit_sell_price, 2),
            round(designing_10, 2),
            round(negotiation_10, 2),
            qty_per_set,
            round(total_price, 2),
            MARKUP,
            round(usd_price, 2),
        ]],
        columns=[
            "S.No.",
            "Item Description",
            "Bought Out Cost Price",
            "Bought Out Sell Price",
            "Inhouse Cost Price",
            "Inhouse Sell Price",
            "Unit Cost Price",
            "Unit Sell Price",
            "10% Designing",
            "10% Negotiation",
            "Qty/Set",
            "Total Price",
            "Markup",
            "USD",
        ],
    )

    return cost_summary_df
