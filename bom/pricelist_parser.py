"""
bom/pricelist_parser.py

Parses all sheets of the ENCON Pricelist WorkBook into SQLite tables.
Used by both init_db.py and the /api/upload-pricelist endpoint.
"""
import re
import math
import sqlite3
import pandas as pd


# ─────────────────────────────────────────────────────────────────
# SS Pipe price calculation
# Formula matches Excel:  E = C * G
#   G (kg/mtr) = π/4 × (OD² - ID²) × 7850 / 1,000,000
#   E (Rs/mtr) = price_per_kg × kg_per_mtr
# ─────────────────────────────────────────────────────────────────

def _ss_pipe_price_per_mtr(item_name: str, price_per_kg: float):
    """
    Parse OD and wall from name like 'SS Pipe 304 60 X 3mm'
    and return price per metre.  Returns None if name doesn't match.
    """
    m = re.search(r'(\d+)\s*[Xx]\s*(\d+(?:\.\d+)?)\s*mm', item_name, re.I)
    if not m:
        return None
    od_mm   = float(m.group(1))
    wall_mm = float(m.group(2))
    id_mm   = od_mm - 2 * wall_mm
    kg_per_mtr = math.pi * (od_mm**2 - id_mm**2) * 7850 / (4 * 1_000_000)
    return round(price_per_kg * kg_per_mtr, 4)


def safe_float(v):
    try:
        return float(str(v).replace(",", "").strip())
    except Exception:
        return None


def _find_sheet(xl, keyword):
    """Case-insensitive sheet name search."""
    for s in xl.sheet_names:
        if keyword.lower() in s.strip().lower():
            return s
    return None


# ─────────────────────────────────────────────────────────────────
# 1. RATES → component_price_master
# ─────────────────────────────────────────────────────────────────

def parse_rates(xl, conn):
    sheet = _find_sheet(xl, "rates")
    if sheet is None:
        return {"skipped": "Rates sheet not found"}

    df = xl.parse(sheet, header=None)

    SKIP_LOWER = {
        "price", "item", "previous", "bought out items",
        "encon purchase price", "specification", "s.no",
        "price list data", "price june", "out items",
    }

    def is_header(v):
        if not isinstance(v, str):
            return False
        return any(kw in v.lower() for kw in SKIP_LOWER)

    def clean_num(v):
        try:
            f = float(v)
            return f if f > 0 else None
        except (TypeError, ValueError):
            return None

    def clean_text(v):
        if not isinstance(v, str):
            return None
        s = v.strip()
        return s if len(s) >= 2 and not s.replace(".", "").replace(",", "").isdigit() else None

    rows = []
    for _, row in df.iterrows():
        # Group A: Raw Material (cols 1, 2, 3)
        item  = clean_text(row.iloc[1] if len(row) > 1 else None)
        price = clean_num(row.iloc[2]  if len(row) > 2 else None)
        prev  = clean_num(row.iloc[3]  if len(row) > 3 else None)
        if item and price and not is_header(item):
            # SS Pipe: compute price per metre from geometry formula
            if re.search(r'ss\s*pipe', item, re.I) and re.search(r'\d+\s*[Xx]\s*\d+\s*mm', item, re.I):
                price_mtr = _ss_pipe_price_per_mtr(item, price)
                if price_mtr is not None:
                    prev_mtr = _ss_pipe_price_per_mtr(item, prev) if prev else price_mtr
                    rows.append((item, "Raw Material", "mtr", price_mtr, prev_mtr))
                    continue
            unit = "kg" if price <= 500 else "nos"
            if "per mtr" in item.lower() or "(per mtr)" in item.lower():
                unit = "mtr"
            rows.append((item, "Raw Material", unit, price, prev or price))

        # Group B: Bought Out (cols 9, 10, 12)
        item  = clean_text(row.iloc[9]  if len(row) > 9  else None)
        price = clean_num(row.iloc[10]  if len(row) > 10 else None)
        prev  = clean_num(row.iloc[12]  if len(row) > 12 else None)
        if item and price and not is_header(item):
            rows.append((item, "Bought Out", "nos", price, prev or price))

        # Group C: ENCON Purchase (cols 15, 19, 20)
        item  = clean_text(row.iloc[15] if len(row) > 15 else None)
        price = clean_num(row.iloc[19]  if len(row) > 19 else None)
        prev  = clean_num(row.iloc[20]  if len(row) > 20 else None)
        if item and price and not is_header(item):
            rows.append((item, "ENCON Purchase", "nos", price, prev or price))

    if not rows:
        return {"error": "No price data found in Rates sheet"}

    # Deduplicate — keep last occurrence
    seen = {}
    for r in rows:
        seen[r[0]] = r
    rows = list(seen.values())

    conn.execute("""
        CREATE TABLE IF NOT EXISTS component_price_master (
            item TEXT PRIMARY KEY, category TEXT,
            unit TEXT, price REAL, previous_price REAL
        )""")
    for item, category, unit, price, prev in rows:
        conn.execute("""
            INSERT INTO component_price_master (item, category, unit, price, previous_price)
            VALUES (?,?,?,?,?)
            ON CONFLICT(item) DO UPDATE SET
                category=excluded.category, unit=excluded.unit,
                price=excluded.price, previous_price=excluded.previous_price
        """, (item, category, unit, price, prev))

    return {"rows": len(rows)}


# ─────────────────────────────────────────────────────────────────
# 2. HPU → hpu_master
# ─────────────────────────────────────────────────────────────────

def parse_hpu(xl, conn):
    sheet = _find_sheet(xl, "hpu")
    if sheet is None:
        return {"skipped": "HPU sheet not found"}

    df = xl.parse(sheet, header=None)
    row0, row1, row2 = df.iloc[0], df.iloc[1], df.iloc[2]

    col_map = {}
    title_cols = [(i, str(v)) for i, v in enumerate(row0) if pd.notna(v) and "Costing" in str(v)]
    variant_cols = [(i, str(v)) for i, v in enumerate(row1)
                    if pd.notna(v) and str(v).strip() not in ("nan", "")]

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

    df_out = pd.DataFrame(records).dropna(subset=["item"])
    df_out = df_out[df_out["item"].str.strip() != ""]
    df_out.to_sql("hpu_master", conn, if_exists="replace", index=False)
    return {"rows": len(df_out)}


# ─────────────────────────────────────────────────────────────────
# 3. BURNER → burner_pricelist_master
# ─────────────────────────────────────────────────────────────────

def parse_burner(xl, conn):
    sheet = _find_sheet(xl, "burner")
    if sheet is None:
        return {"skipped": "BURNER sheet not found"}

    df = xl.parse(sheet, header=None)
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
                            "price": price,
                        })

    df_out = pd.DataFrame(records)
    if not df_out.empty:
        df_out["section"] = df_out["section"].str.replace("`", "'", regex=False)
    df_out.to_sql("burner_pricelist_master", conn, if_exists="replace", index=False)
    return {"rows": len(df_out)}


# ─────────────────────────────────────────────────────────────────
# 4. BLOWER → blower_pricelist_master
# ─────────────────────────────────────────────────────────────────

def parse_blower(xl, conn):
    sheet = _find_sheet(xl, "blower")
    if sheet is None:
        return {"skipped": "Blower sheet not found"}

    df = xl.parse(sheet, header=None)
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

    df_out = pd.DataFrame(records)
    df_out.to_sql("blower_pricelist_master", conn, if_exists="replace", index=False)

    # Also sync to blower_master (used by blower_selector.py)
    _sync_blower_master(df_out, conn)

    return {"rows": len(df_out)}


def _sync_blower_master(df_pricelist: pd.DataFrame, conn):
    """
    Normalise blower_pricelist_master into the fixed-column schema that
    blower_selector.py expects:
        model, hp, airflow, cfm, pressure, price_basic, price_premium

    ENCON prices (cols F/G): read directly — manually set selling prices.
    PERKIN prices (cols L/M): computed from input amounts (col J, K):
        price_without_motor = J × 1.8
        price_with_motor    = K × 1.5 + J × 1.8
    This replicates the Excel formulas =J*1.8 and =K*1.5+L so that if
    J or K changes in the sheet, Python recalculates without needing
    Excel to recalculate first.
    """
    if df_pricelist.empty:
        return

    cols = df_pricelist.columns.tolist()

    def _find_col(*keywords):
        for c in cols:
            cl = c.lower()
            if all(kw in cl for kw in keywords):
                return c
        return None

    airflow_col  = _find_col("nm3") or _find_col("airflow") or _find_col("flow")
    cfm_col      = _find_col("cfm")
    pressure_col = _find_col("pressure")

    # ENCON price columns (manually set, read directly)
    encon_basic_col = _find_col("price_without") or _find_col("price without")
    # "price with motor" but NOT "without"
    encon_prem_col = None
    for c in cols:
        cl = c.lower()
        if "price" in cl and "with" in cl and "motor" in cl and "without" not in cl:
            encon_prem_col = c
            break

    # PERKIN input columns: Amount (J) and Motor Price (K)
    amount_col      = _find_col("amount")
    motor_price_col = _find_col("motor_price") or _find_col("motor price")

    rows = []
    for _, r in df_pricelist.iterrows():
        model = r.get("model")
        if not model or str(model).lower() in ("nan", ""):
            continue

        # Try ENCON price first; fall back to computing from PERKIN amounts
        price_basic = safe_float(r[encon_basic_col]) if encon_basic_col else None
        price_prem  = safe_float(r[encon_prem_col])  if encon_prem_col  else None

        # If ENCON prices are missing/zero, compute from PERKIN amounts
        if not price_basic and amount_col:
            j = safe_float(r[amount_col]) or 0
            k = safe_float(r[motor_price_col]) if motor_price_col else 0
            if j:
                price_basic = round(j * 1.8, 2)
                price_prem  = round(k * 1.5 + price_basic, 2) if k else None

        rows.append((
            str(model),
            str(r.get("hp", "")) if "hp" in cols else None,
            str(r[airflow_col])  if airflow_col  else None,
            str(r[cfm_col])      if cfm_col      else None,
            str(r[pressure_col]) if pressure_col else None,
            str(price_basic)     if price_basic is not None else None,
            str(price_prem)      if price_prem  is not None else None,
        ))

    conn.execute("DELETE FROM blower_master")
    conn.executemany(
        "INSERT INTO blower_master "
        "(model, hp, airflow, cfm, pressure, price_basic, price_premium) "
        "VALUES (?,?,?,?,?,?,?)",
        rows,
    )


# ─────────────────────────────────────────────────────────────────
# 5. HORIZONTAL → horizontal_master
# ─────────────────────────────────────────────────────────────────

def parse_horizontal(xl, conn):
    sheet = _find_sheet(xl, "horizontal")
    if sheet is None:
        return {"skipped": "Horizontal sheet not found"}

    df = xl.parse(sheet, header=None)
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
                except Exception:
                    continue

            records.append({"model": current_model, "particular": particular, "qty": qty, "amount": amount})

    df_out = pd.DataFrame(records)
    if not df_out.empty:
        df_out = df_out[["model", "particular", "qty", "amount"]]
    df_out.to_sql("horizontal_master", conn, if_exists="replace", index=False)
    return {"rows": len(df_out)}


# ─────────────────────────────────────────────────────────────────
# 6. VERTICAL → vertical_master
# ─────────────────────────────────────────────────────────────────

def parse_vertical(xl, conn):
    sheet = _find_sheet(xl, "vertical")
    if sheet is None:
        return {"skipped": "Vertical sheet not found"}

    df = xl.parse(sheet, header=None)
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
                    except Exception:
                        continue

                records.append({"model": current_model, "particular": particular, "qty": qty, "amount": amount})

    df_out = pd.DataFrame(records)
    if not df_out.empty:
        df_out = df_out[["model", "particular", "qty", "amount"]]
    df_out.to_sql("vertical_master", conn, if_exists="replace", index=False)
    return {"rows": len(df_out)}


# ─────────────────────────────────────────────────────────────────
# 7. PARTS SHEETS (Oil Burner, HV Oil Burner, Gas Burner)
# ─────────────────────────────────────────────────────────────────

def _parse_parts_sheet(xl, sheet_name, table_name, conn):
    sheet = _find_sheet(xl, sheet_name)
    if sheet is None:
        return {"skipped": f"'{sheet_name}' sheet not found"}

    df = xl.parse(sheet, header=None).dropna(how="all")
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

    pd.DataFrame(records).to_sql(table_name, conn, if_exists="replace", index=False)
    return {"rows": len(records)}


# ─────────────────────────────────────────────────────────────────
# 8. RECUPERATOR → recuperator_master
#
# Replicates the Excel formulas exactly:
#
#   Type F  : col_cost = (qty / 0.3) × rate  for tube_c
#                        qty × rate           for all others
#             selling  = SUM(cols B–G) × 1.3 × 1.8
#
#   Type HT : selling  = SUM(cols N–Q) × 1.3 × 1.8
#
#   Type FS : col_cost = qty × 4.27 × rate   for tube_c and tube_b
#                        qty × rate           for all others
#             selling  = SUM(cols B–I) × 1.3 × 1.8
#
# Material prices come from component_price_master (already populated
# by parse_rates which runs first in parse_all).
# ─────────────────────────────────────────────────────────────────

def parse_recuperator(xl, conn):
    sheet = _find_sheet(xl, "recuperator")
    if sheet is None:
        return {"skipped": "Recuperator sheet not found"}

    # ── fetch live rates from component_price_master ──────────────
    def _rate(item_prefix):
        try:
            row = conn.execute(
                "SELECT price FROM component_price_master WHERE item LIKE ?",
                (item_prefix.strip() + '%',)
            ).fetchone()
            return float(row[0]) if row else 0.0
        except Exception:
            return 0.0

    R = {
        'tube_c':    _rate('M.S. Tube "C" Class'),
        'tube_b':    _rate('M.S. Tube "B" Class'),
        'ci_gills':  _rate('C.I. Gills'),
        'ang_6550':  _rate('M.S. Angle 65,50'),
        'ang_100':   _rate('M.S. Angle 100,100'),
        'plate_5mm': _rate('M.S. Plate 16mm* 5mm'),
        'bolt':      _rate('Hardware Bolt'),
        'channel':   _rate('M.S. Chanel'),
        'plate_10mm':_rate('M.S. Plate 16mm*10mm'),
        'plate_5mm2':_rate('M.S. Plate 5mm'),
        'ang_50':    _rate('M.S. Angle 50*6'),
        'ss_sheet':  _rate('S.S. Sheet 3mm'),
    }

    # ── read the raw quantity table (data_only to get values not formulas) ──
    import openpyxl
    wb_vals = openpyxl.load_workbook(xl.io, read_only=True, data_only=True)
    ws = wb_vals[sheet]
    rows_vals = list(ws.iter_rows(values_only=True))
    wb_vals.close()

    def _v(row_idx, col_idx):
        """0-based row and col. Returns float or 0."""
        try:
            v = rows_vals[row_idx][col_idx]
            return float(v) if v is not None else 0.0
        except Exception:
            return 0.0

    MARKUP = 1.3 * 1.8  # = 2.34

    records = []

    # ── Type F (rows 3–8, 0-indexed 2–7) ─────────────────────────
    # cols: A=0,B=1,C=2,D=3,E=4,F=5,G=6,H=7
    F_MODELS_ROWS = range(2, 8)   # rows 3–8
    for ri in F_MODELS_ROWS:
        model = rows_vals[ri][0]
        if model is None or not str(model).strip():
            continue
        b = _v(ri, 1)   # MS Tube C qty
        c = _v(ri, 2)   # CI Gills qty
        d = _v(ri, 3)   # Angle 65,50
        e = _v(ri, 4)   # Angle 100,100
        f = _v(ri, 5)   # Plate 16mm*5mm
        g = _v(ri, 6)   # Hardware Bolt
        # (col H / MS Channel is calculated but excluded from the Excel SUM)
        tube_cost   = (b / 0.3) * R['tube_c'] if b else 0
        ci_cost     = c * R['ci_gills']
        ang65_cost  = d * R['ang_6550']
        ang100_cost = e * R['ang_100']
        plate_cost  = f * R['plate_5mm']
        bolt_cost   = g * R['bolt']
        fab_cost    = tube_cost + ci_cost + ang65_cost + ang100_cost + plate_cost + bolt_cost
        selling     = round(fab_cost * MARKUP, 2)
        records.append({"type": "F", "model": str(model).strip(),
                        "fabrication_cost": round(fab_cost, 2), "selling_price": selling})

    # ── Type HT (same input rows, cols M–Q: M=12,N=13,O=14,P=15,Q=16) ──
    HT_MODELS_ROWS = range(2, 6)   # rows 3–6 only (4 HT models)
    for ri in HT_MODELS_ROWS:
        model = rows_vals[ri][12]
        if model is None or not str(model).strip():
            continue
        n = _v(ri, 13)   # Plate 16mm*10mm
        o = _v(ri, 14)   # Plate 5mm
        p = _v(ri, 15)   # Angle 50*6
        q = _v(ri, 16)   # SS Sheet 3mm
        fab_cost = (n * R['plate_10mm'] + o * R['plate_5mm2'] +
                    p * R['ang_50']    + q * R['ss_sheet'])
        selling  = round(fab_cost * MARKUP, 2)
        records.append({"type": "HT", "model": str(model).strip(),
                        "fabrication_cost": round(fab_cost, 2), "selling_price": selling})

    # ── Type FS (rows 21–27, 0-indexed 20–26) ────────────────────
    # cols: A=0,B=1(tube_c),C=2(tube_b),D=3(ci_gills),E=4,F=5,G=6,H=7(bolt),I=8(channel)
    FS_MODELS_ROWS = range(20, 27)
    for ri in FS_MODELS_ROWS:
        model = rows_vals[ri][0]
        if model is None or not str(model).strip():
            continue
        b = _v(ri, 1)   # tube C qty
        c = _v(ri, 2)   # tube B qty
        e = _v(ri, 4)   # Angle 65,50
        f = _v(ri, 5)   # Angle 100,100
        g = _v(ri, 6)   # Plate 16mm*5mm
        h = _v(ri, 7)   # Hardware Bolt
        i = _v(ri, 8)   # MS Channel (included in FS sum)
        tube_c_cost = b * 4.27 * R['tube_c']
        tube_b_cost = c * 4.27 * R['tube_b']
        ang65_cost  = e * R['ang_6550']
        ang100_cost = f * R['ang_100']
        plate_cost  = g * R['plate_5mm']
        bolt_cost   = h * R['bolt']
        chan_cost    = i * R['channel']
        fab_cost    = (tube_c_cost + tube_b_cost + ang65_cost + ang100_cost +
                       plate_cost + bolt_cost + chan_cost)
        selling     = round(fab_cost * MARKUP, 2)
        records.append({"type": "FS", "model": str(model).strip(),
                        "fabrication_cost": round(fab_cost, 2), "selling_price": selling})

    pd.DataFrame(records).to_sql("recuperator_master", conn, if_exists="replace", index=False)
    return {"rows": len(records)}


# ─────────────────────────────────────────────────────────────────
# 9. GAIL GAS BURNER → gail_gas_burner_master
# ─────────────────────────────────────────────────────────────────

def parse_gail_gas_burner(xl, conn):
    sheet = _find_sheet(xl, "gail")
    if sheet is None:
        return {"skipped": "GAIL GAS Burner sheet not found"}

    df = xl.parse(sheet, header=None).dropna(how="all")
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

    pd.DataFrame(records).to_sql("gail_gas_burner_master", conn, if_exists="replace", index=False)
    return {"rows": len(records)}


# ─────────────────────────────────────────────────────────────────
# 10. RAD HEAT → rad_heat_master / rad_heat_tata_master
# ─────────────────────────────────────────────────────────────────

def _parse_rad_heat(xl, sheet_name, table_name, conn):
    sheet = _find_sheet(xl, sheet_name)
    if sheet is None:
        return {"skipped": f"'{sheet_name}' sheet not found"}

    df = xl.parse(sheet, header=None).dropna(how="all")
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

    pd.DataFrame(records).to_sql(table_name, conn, if_exists="replace", index=False)
    return {"rows": len(records)}


# ─────────────────────────────────────────────────────────────────
# MAIN: parse everything
# ─────────────────────────────────────────────────────────────────

def parse_all(file_path: str, conn: sqlite3.Connection) -> dict:
    """
    Run all parsers against the given pricebook file.
    Returns a dict mapping table_name → {rows: N} or {skipped: reason} or {error: msg}.
    """
    xl = pd.ExcelFile(file_path)
    results = {}

    parsers = [
        ("component_price_master",    lambda: parse_rates(xl, conn)),
        ("hpu_master",                lambda: parse_hpu(xl, conn)),
        ("burner_pricelist_master",   lambda: parse_burner(xl, conn)),
        ("blower_pricelist_master",   lambda: parse_blower(xl, conn)),
        ("horizontal_master",         lambda: parse_horizontal(xl, conn)),
        ("vertical_master",           lambda: parse_vertical(xl, conn)),
        ("oil_burner_parts_master",   lambda: _parse_parts_sheet(xl, "oil burner", "oil_burner_parts_master", conn)),
        ("hv_oil_burner_parts_master",lambda: _parse_parts_sheet(xl, "hv  oil burner", "hv_oil_burner_parts_master", conn)),
        ("gas_burner_parts_master",   lambda: _parse_parts_sheet(xl, "gas burner", "gas_burner_parts_master", conn)),
        ("recuperator_master",        lambda: parse_recuperator(xl, conn)),
        ("gail_gas_burner_master",    lambda: parse_gail_gas_burner(xl, conn)),
        ("rad_heat_tata_master",      lambda: _parse_rad_heat(xl, "rad heat (tata)", "rad_heat_tata_master", conn)),
        ("rad_heat_master",           lambda: _parse_rad_heat(xl, "rad heat", "rad_heat_master", conn)),
    ]

    for table, fn in parsers:
        try:
            results[table] = fn()
        except Exception as exc:
            results[table] = {"error": str(exc)}

    conn.commit()
    return results
