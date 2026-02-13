"""
FINAL frozen Excel writer
Matches legacy Excel logic EXACTLY
"""

import pandas as pd
from typing import Dict
from openpyxl import load_workbook
from export.calculation_sheet import write_calculation_sheet


# -------------------------------------------------
# LEGACY EXCLUSION RULES
# These items APPEAR in BOM
# but DO NOT contribute to BOUGHT OUT ITEMS cost
# -------------------------------------------------
BOUGHT_OUT_EXCLUDE_ITEMS = {
    # Instruments / internal
    "COMPENSATOR",
    "PRESSURE GAUGE WITH TNV",
   
    # Controller counted under ENCON margin
    "RATIO CONTROLLER",
}


def write_excel(
    buffer,
    sheets: Dict[str, pd.DataFrame],
    burner_inputs,
    burner_results,
    pipe_results,
):
    """
    Writes Excel with:
    - VLPH BOM
    - Correct legacy BOUGHT OUT / ENCON totals
    - Cost Summary
    - Calculation sheet
    """

    # -------------------------------------------------
    # 1️⃣ Write BOM & Cost Summary (xlsxwriter)
    # -------------------------------------------------
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        workbook = writer.book

        # ---------- BOM ----------
        bom_df = sheets["VLPH-120T"].copy()

        # Ensure numeric totals
        bom_df["TOTAL"] = pd.to_numeric(
            bom_df["TOTAL"], errors="coerce"
        ).fillna(0)

        # ---------- ENCON TOTAL ----------
        encon_total = bom_df.loc[
            bom_df["MEDIA"] == "ENCON ITEMS",
            "TOTAL",
        ].sum()

        print(
        bom_df.loc[
        bom_df["ITEM NAME"].isin(BOUGHT_OUT_EXCLUDE_ITEMS),
        ["ITEM NAME", "QTY", "TOTAL"]
    ]
)


        # ---------- BOUGHT OUT TOTAL (LEGACY LOGIC) ----------
        bought_out_total = bom_df.loc[
            (bom_df["MEDIA"] != "ENCON ITEMS")
            & (~bom_df["ITEM NAME"].isin(BOUGHT_OUT_EXCLUDE_ITEMS)),
            "TOTAL",
        ].sum()

        # ---------- Append TOTAL rows (DISPLAY ONLY) ----------
        totals_df = pd.DataFrame(
            [
               
                
            ]
        )

        final_bom_df = pd.concat(
            [bom_df, totals_df],
            ignore_index=True,
        )

        final_bom_df.to_excel(
            writer,
            sheet_name="VLPH-120T",
            index=False,
        )

        worksheet = writer.sheets["VLPH-120T"]

        header_fmt = workbook.add_format({
            "bold": True,
            "border": 1,
            "align": "center",
            "valign": "vcenter",
        })

        for col_num, col_name in enumerate(final_bom_df.columns):
            worksheet.write(0, col_num, col_name, header_fmt)
            worksheet.set_column(col_num, col_num, 22)

        # ---------- Cost Summary ----------
        if "Cost Summary" in sheets:
            sheets["Cost Summary"].to_excel(
                writer,
                sheet_name="Cost Summary",
                index=False,
            )

    # -------------------------------------------------
    # 2️⃣ Inject Calculation Sheet (openpyxl)
    # -------------------------------------------------
    buffer.seek(0)
    wb = load_workbook(buffer)

    # Insert calculation sheet as FIRST sheet (legacy behavior)
    ws = wb.create_sheet("Calculation", 0)

    write_calculation_sheet(
        ws,
        burner_inputs,
        burner_results,
        pipe_results,
    )

    buffer.seek(0)
    wb.save(buffer)
