import sqlite3
import os
import re
import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "vlph.db")
FILE_PATH = os.path.join(BASE_DIR, "Pricelist WorkBook 28-08-2025.xlsx")

# =================================================
# CONNECTION
# =================================================

conn = sqlite3.connect(DB_PATH)
cursor = conn.cursor()

xls = pd.ExcelFile(FILE_PATH)
print("SHEETS FOUND:", xls.sheet_names)


def get_sheet(keyword):
    for s in xls.sheet_names:
        if keyword.lower() in s.lower():
            return s
    raise Exception(f"Sheet not found for keyword: {keyword}")


def safe_float(v):
    try:
        return float(str(v).replace(",", "").strip())
    except:
        return None


def to_sql(df, table, conn):
    df.to_sql(table, conn, if_exists="replace", index=False)
    print(f"  ✅ {table}: {len(df)} rows")


# =================================================
# DROP ALL RAW TABLES (cleanup)
# =================================================

cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name LIKE '%_raw'")
raw_tables = [r[0] for r in cursor.fetchall()]
for t in raw_tables:
    cursor.execute(f"DROP TABLE IF EXISTS [{t}]")
    print(f"  🗑️  Dropped: {t}")

conn.commit()


# =================================================
# 1. HORIZONTAL MASTER
# =================================================

print("\n📦 HORIZONTAL")
cursor.execute("DROP TABLE IF EXISTS horizontal_master")

df = pd.read_excel(FILE_PATH, sheet_name="Horizontal", header=None)
records = []
current_model = None

for _, row in df.iterrows():
    row_values = [str(x).strip() for x in row if pd.notna(x)]
    if not row_values:
        continue

    found_model = False
    for val in row_values:
        if "HORIZONTAL LADLE PREHEATER" in val.upper():
            current_model = val.upper()
            found_model = True
            break
    if found_model:
        continue

    first_val = row_values[0].upper()
    if "TOTAL" in first_val or "COMBUSTION EQUIPMENT" in first_val:
        continue

    if len(row_values) >= 2 and current_model:
        if row_values[0].isdigit():
            particular = row_values[1]
            values = row_values[2:]
        else:
            particular = row_values[0]
            values = row_values[1:]

        particular = particular.strip().upper()
        if not particular:
            continue

        qty = None
        amount = None
        for val in values:
            if any(x in val.upper() for x in ["KGS", "ROLLS", "SET", "NO"]):
                qty = val
            try:
                amount = float(val.replace(",", ""))
            except:
                continue

        records.append({"model": current_model, "particular": particular, "qty": qty, "amount": amount})

df_out = pd.DataFrame(records)
if not df_out.empty:
    to_sql(df_out[["model", "particular", "qty", "amount"]], "horizontal_master", conn)
else:
    print("❌ Horizontal parsing failed")


# =================================================
# 2. VERTICAL MASTER
# =================================================

print("\n📦 VERTICAL")
cursor.execute("DROP TABLE IF EXISTS vertical_master")

df = pd.read_excel(FILE_PATH, sheet_name="Vertical", header=None)
records = []

for start_col in range(0, df.shape[1], 5):
    block = df.iloc[:, start_col:start_col + 4].copy().dropna(how="all")
    if block.empty:
        continue

    current_model = None
    for _, row in block.iterrows():
        row_values = [str(x).strip() for x in row if pd.notna(x)]
        if not row_values:
            continue

        for val in row_values:
            if "VERTICAL" in val.upper():
                current_model = val.upper()
                break

        if current_model is None:
            continue

        first_val = row_values[0].upper()
        if "TOTAL" in first_val or "COMBUSTION EQUIPMENT" in first_val:
            continue

        if len(row_values) >= 2:
            if row_values[0].isdigit():
                particular = row_values[1]
                values = row_values[2:]
            else:
                particular = row_values[0]
                values = row_values[1:]

            particular = particular.strip().upper()
            if not particular:
                continue

            qty = None
            amount = None
            for val in values:
                if any(x in val.upper() for x in ["KGS", "ROLLS", "SET", "NO"]):
                    qty = val
                try:
                    amount = float(val.replace(",", ""))
                except:
                    continue

            records.append({"model": current_model, "particular": particular, "qty": qty, "amount": amount})

df_out = pd.DataFrame(records)
if not df_out.empty:
    to_sql(df_out[["model", "particular", "qty", "amount"]], "vertical_master", conn)
else:
    print("❌ Vertical parsing failed")


# =================================================
# 3. RATES MASTER
# =================================================

print("\n📦 RATES")
cursor.execute("DROP TABLE IF EXISTS rates_master")

df = pd.read_excel(FILE_PATH, sheet_name="Rates", header=None).dropna(how="all")
records = []
current_category = None

for _, row in df.iterrows():
    row_values = [str(x).strip() for x in row if pd.notna(x)]
    if not row_values:
        continue
    if len(row_values) == 1:
        current_category = row_values[0].upper()
        continue
    if len(row_values) >= 2:
        if row_values[0].isdigit():
            item = row_values[1]
            values = row_values[2:]
        else:
            item = row_values[0]
            values = row_values[1:]

        item = item.strip().upper()
        rate = None
        for val in values:
            try:
                rate = float(val)
                break
            except:
                continue

        unit = "NA"
        for val in values:
            if val.upper() in ["KGS", "KG", "NOS", "NO", "SET"]:
                unit = val.upper()

        if rate is None:
            continue

        records.append({"item": item, "rate": rate, "unit": unit, "category": current_category})

df_out = pd.DataFrame(records)
if not df_out.empty:
    to_sql(df_out, "rates_master", conn)
else:
    print("❌ Rates parsing failed")


# =================================================
# 4. OIL BURNER PARTS MASTER
# =================================================

print("\n📦 OIL BURNER PARTS")

def parse_parts_sheet(sheet_name, table_name):
    df = pd.read_excel(FILE_PATH, sheet_name=sheet_name, header=None).dropna(how="all")
    records = []
    current_section = None

    for _, row in df.iterrows():
        vals = [str(x).strip() for x in row if pd.notna(x) and str(x).strip() not in ("nan", "")]
        if not vals:
            continue
        if len(vals) == 1 and safe_float(vals[0]) is None:
            current_section = vals[0].upper()
            continue
        if len(vals) < 2:
            continue
        if vals[0].upper() in ("S.NO.", "SNO", "SL NO"):
            continue

        idx = 1 if re.match(r'^\d+$', vals[0]) else 0
        particular = vals[idx].upper().strip() if idx < len(vals) else None
        if not particular or particular in ("PARTICULARS", "TOTAL", ""):
            continue

        qty = unit = rate = amount = None
        rest = vals[idx + 1:]
        for v in rest:
            if v.upper() in ("KGS", "KG", "NOS", "NO", "SET", "MTR", "ROLLS"):
                unit = v.upper()
            else:
                f = safe_float(v)
                if f is not None:
                    if rate is None:
                        rate = f
                    else:
                        amount = f
        for v in rest:
            f = safe_float(v)
            if f is not None and qty is None and f != rate and f != amount:
                qty = f
                break

        records.append({"section": current_section, "particular": particular,
                        "qty": qty, "unit": unit, "rate": rate, "amount": amount})

    to_sql(pd.DataFrame(records), table_name, conn)

parse_parts_sheet(" Oil Burner", "oil_burner_parts_master")
parse_parts_sheet("HV  Oil Burner", "hv_oil_burner_parts_master")
parse_parts_sheet("Gas Burner", "gas_burner_parts_master")


# =================================================
# 5. BURNER PRICE LIST MASTER
# =================================================

print("\n📦 BURNER PRICE LIST")
df = pd.read_excel(FILE_PATH, sheet_name="BURNER", header=None)
records = []
current_section = None
headers = None

for _, row in df.iterrows():
    vals_raw = [str(x).strip() if pd.notna(x) else None for x in row]
    non_null = [v for v in vals_raw if v and v != "nan"]
    if not non_null:
        continue

    first = non_null[0]

    if len(non_null) == 1 and safe_float(first) is None:
        current_section = first.upper().replace("`", "'")
        headers = None
        continue

    if "BURNER SIZE" in first.upper() or first.upper() == "BURNER SIZE":
        headers = [v for v in non_null if v]
        continue

    if headers is None and "BURNER" in first.upper() and safe_float(first) is None and len(non_null) > 2:
        headers = [v for v in non_null if v]
        continue

    if headers and current_section:
        data_vals = [v for v in non_null if v]
        if len(data_vals) < 2:
            continue
        burner_size = data_vals[0]
        for i, h in enumerate(headers[1:], 1):
            if i < len(data_vals):
                price = safe_float(data_vals[i])
                if price is not None:
                    records.append({
                        "section": current_section,
                        "burner_size": burner_size.upper(),
                        "component": h.upper(),
                        "price": price
                    })

df_burner = pd.DataFrame(records)
df_burner["section"] = df_burner["section"].str.replace("`", "'", regex=False)
to_sql(df_burner, "burner_pricelist_master", conn)


# =================================================
# 6. BLOWER PRICELIST MASTER
# =================================================

print("\n📦 BLOWER")
df = pd.read_excel(FILE_PATH, sheet_name="Blower", header=None)
records = []
current_section = None
headers = None

for _, row in df.iterrows():
    vals_raw = [str(x).strip() if pd.notna(x) else None for x in row]
    non_null = [v for v in vals_raw if v and v != "nan"]
    if not non_null:
        continue

    first = non_null[0]

    if first.upper() in ("MEDIUM PRESSURE", "HIGH PRESSURE", "BLOWER IDM") or \
       ("BLOWER" in first.upper() and "ENCON" not in first.upper() and len(non_null) == 1):
        current_section = first.upper()
        headers = None
        continue

    if first.upper().startswith("NOTE") or first.startswith("("):
        continue

    if first.upper() == "MODEL" or "MODEL" in first.upper():
        headers = [v for v in non_null if v]
        continue

    if headers and current_section and (first.upper().startswith("ENCON") or re.match(r'^\d+', first)):
        data_vals = [v for v in non_null if v]
        if len(data_vals) < 2:
            continue
        rec = {"section": current_section, "model": data_vals[0]}
        for i, h in enumerate(headers[1:], 1):
            if i < len(data_vals):
                f = safe_float(data_vals[i])
                col = h.lower().replace(" ", "_").replace("/", "_per_").replace(".", "")[:40]
                rec[col] = f if f is not None else data_vals[i]
        records.append(rec)

to_sql(pd.DataFrame(records), "blower_pricelist_master", conn)


# =================================================
# 7. RECUPERATOR MASTER
# =================================================

print("\n📦 RECUPERATOR")
df = pd.read_excel(FILE_PATH, sheet_name="Recuperator", header=None)
records = []
current_type = None
headers = None

for _, row in df.iterrows():
    vals_raw = [str(x).strip() if pd.notna(x) else None for x in row]
    non_null = [v for v in vals_raw if v and v != "nan"]
    if not non_null:
        continue

    first = non_null[0]

    if "RECUPERATOR" in first.upper() and len(non_null) == 1:
        current_type = first.upper()
        headers = None
        continue

    if first.upper() == "MODEL":
        headers = [v for v in non_null if v]
        continue

    if headers and current_type and re.match(r'^\d+', first):
        data_vals = [v for v in non_null if v]
        if len(data_vals) < 2:
            continue
        rec = {"type": current_type, "model": data_vals[0]}
        for i, h in enumerate(headers[1:], 1):
            if i < len(data_vals):
                col = h.lower().replace('"', 'in').replace(' ', '_').replace('(', '').replace(')', '').replace('.', '').replace(',', '_')[:40]
                rec[col] = safe_float(data_vals[i])
        records.append(rec)

to_sql(pd.DataFrame(records), "recuperator_master", conn)


# =================================================
# 8. RAD HEAT MASTER (Standard + TATA)
# =================================================

print("\n📦 RAD HEAT")

def parse_rad_heat(sheet_name, table_name):
    df = pd.read_excel(FILE_PATH, sheet_name=sheet_name, header=None).dropna(how="all")
    records = []
    section = None

    for _, row in df.iterrows():
        non_null = [str(x).strip() for x in row if pd.notna(x) and str(x).strip() not in ("nan", "")]
        if not non_null:
            continue
        first = non_null[0]

        if "PRICE LIST OF RAD-HEAT" in first.upper():
            section = "models"
            continue
        if "SPARES PRICE" in first.upper():
            section = "spares"
            continue
        if first.upper() in ("MODEL", "ITEM"):
            continue

        if section == "models" and first.upper().startswith("ARE-"):
            records.append({
                "section": "MODEL",
                "item": first.upper(),
                "output_kw": non_null[1] if len(non_null) > 1 else None,
                "gas_lpg_m3hr": safe_float(non_null[2]) if len(non_null) > 2 else None,
                "gas_ng_m3hr": safe_float(non_null[3]) if len(non_null) > 3 else None,
                "price_with_ss_tubing": safe_float(non_null[4]) if len(non_null) > 4 else None,
            })

        elif section == "spares" and len(non_null) >= 2:
            price = safe_float(non_null[1]) or (safe_float(non_null[2]) if len(non_null) > 2 else None)
            if price is not None:
                records.append({
                    "section": "SPARE",
                    "item": first.upper(),
                    "output_kw": None,
                    "gas_lpg_m3hr": None,
                    "gas_ng_m3hr": None,
                    "price_with_ss_tubing": price,
                })

    to_sql(pd.DataFrame(records), table_name, conn)

parse_rad_heat("Rad Heat (TATA)", "rad_heat_tata_master")

# Standard Rad Heat — avoid matching TATA sheet
parse_rad_heat("Rad Heat", "rad_heat_master")


# =================================================
# 9. GAIL GAS BURNER MASTER
# =================================================

print("\n📦 GAIL GAS BURNER")
df = pd.read_excel(FILE_PATH, sheet_name="GAIL GAS Burner", header=None).dropna(how="all")
records = []
current_section = None
headers = None

for _, row in df.iterrows():
    non_null = [str(x).strip() for x in row if pd.notna(x) and str(x).strip() not in ("nan", "")]
    if not non_null:
        continue
    first = non_null[0]

    if "GAIL GAS BURNER" in first.upper() or "PILOT BURNER" in first.upper() or "ALL BURNER" in first.upper():
        current_section = first.upper().strip()
        headers = None
        continue

    if "BURNER SIZE" in first.upper():
        headers = non_null
        continue

    if headers and current_section and "ENCON" in first.upper():
        rec = {"section": current_section, "burner_size": first.upper()}
        for i, h in enumerate(headers[1:], 1):
            if i < len(non_null):
                col = h.lower().replace(" ", "_").replace(".", "").replace("(", "").replace(")", "")
                rec[col] = safe_float(non_null[i])
        records.append(rec)
    elif current_section and "PILOT" in current_section and len(non_null) >= 2:
        price = safe_float(non_null[1])
        if price is not None:
            records.append({"section": current_section, "burner_size": first.upper(), "burner": price})

to_sql(pd.DataFrame(records), "gail_gas_burner_master", conn)


# =================================================
# 10. HPU MASTER
# =================================================

print("\n📦 HPU")
df = pd.read_excel(FILE_PATH, sheet_name="HPU", header=None)

row0 = df.iloc[0]
row1 = df.iloc[1]
row2 = df.iloc[2]

col_map = {}
title_cols = [(i, str(v)) for i, v in enumerate(row0) if pd.notna(v) and "Costing" in str(v)]
variant_cols = [(i, str(v)) for i, v in enumerate(row1) if pd.notna(v) and str(v).strip() not in ("nan", "")]

for title_col, title_text in title_cols:
    kw_match = re.search(r'(\d+)\s*KW', title_text, re.I)
    kw = int(kw_match.group(1)) if kw_match else None

    next_titles = [tc for tc, _ in title_cols if tc > title_col]
    end = next_titles[0] if next_titles else df.shape[1]

    for var_col, var_name in variant_cols:
        if title_col <= var_col < end:
            for dc in range(var_col, min(var_col + 8, df.shape[1])):
                cell = str(row2.iloc[dc]).strip().lower() if pd.notna(row2.iloc[dc]) else ""
                if "items" in cell or "item" in cell:
                    col_map[dc] = (kw, var_name, "item_col")
                elif "qty" in cell:
                    col_map[dc] = (kw, var_name, "qty_col")
                elif "unit" in cell:
                    col_map[dc] = (kw, var_name, "unit_col")
                elif "rate" in cell:
                    col_map[dc] = (kw, var_name, "rate_col")
                elif "amount" in cell:
                    col_map[dc] = (kw, var_name, "amount_col")

records = []
for _, row in df.iloc[3:].iterrows():
    block_data = {}
    for col_idx, (kw, variant, field) in col_map.items():
        key = (kw, variant)
        if key not in block_data:
            block_data[key] = {}
        val = row.iloc[col_idx] if col_idx < len(row) else None
        if pd.notna(val):
            block_data[key][field] = str(val).strip()

    for (kw, variant), fields in block_data.items():
        item = fields.get("item_col", "").strip()
        if not item or item.lower() in ("nan", "", "items", "total amount"):
            continue
        records.append({
            "unit_kw": kw,
            "variant": variant,
            "item": item.upper(),
            "qty": safe_float(fields.get("qty_col")),
            "unit": fields.get("unit_col", "").strip() or None,
            "rate": safe_float(fields.get("rate_col")),
            "amount": safe_float(fields.get("amount_col")),
        })

df_hpu = pd.DataFrame(records).dropna(subset=["item"])
df_hpu = df_hpu[df_hpu["item"].str.strip() != ""]
to_sql(df_hpu, "hpu_master", conn)


# =================================================
# FINAL
# =================================================

conn.commit()
conn.close()

print("\n✅ Database initialized successfully.")