import pandas as pd

file_path = "Pricelist WorkBook 28-08-2025.xlsx"
output_file = "excel_inspection.txt"

# Always regenerate file (overwrite mode)
with open(output_file, "w", encoding="utf-8") as f:

    xls = pd.ExcelFile(file_path)
    f.write(f"Sheets: {xls.sheet_names}\n\n")

    for sheet in xls.sheet_names:
        f.write(f"\n==============================\n")
        f.write(f"===== SHEET: {sheet} =====\n")
        f.write(f"==============================\n")

        # -----------------------------
        # RAW READ (default header)
        # -----------------------------
        try:
            df_raw = pd.read_excel(file_path, sheet_name=sheet)

            f.write("\n--- RAW HEADER (default) ---\n")
            f.write("Columns:\n")
            for col in df_raw.columns:
                f.write(f"- {col}\n")

            f.write("\nSample Rows:\n")
            f.write(df_raw.head(3).to_string())
            f.write("\n\n")

        except Exception as e:
            f.write(f"Error reading raw: {e}\n\n")

        # -----------------------------
        # HEADER = 1 READ
        # -----------------------------
        try:
            df_h1 = pd.read_excel(file_path, sheet_name=sheet, header=1)

            f.write("\n--- HEADER = 1 ---\n")
            f.write("Columns:\n")
            for col in df_h1.columns:
                f.write(f"- {col}\n")

            f.write("\nSample Rows:\n")
            f.write(df_h1.head(3).to_string())
            f.write("\n\n")

        except Exception as e:
            f.write(f"Error reading header=1: {e}\n\n")

print(f"Fresh inspection file generated: {output_file}")