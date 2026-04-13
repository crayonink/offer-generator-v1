from fastapi import FastAPI, UploadFile, File, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, FileResponse
from pydantic import BaseModel
from typing import List, Optional
import pandas as pd
import sqlite3
import shutil
import os
from datetime import datetime

from bom.pricelist_parser import parse_all as _parse_pricelist_all

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "vlph.db")
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
QUOTES_FOLDER = os.path.join(BASE_DIR, "quotes")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(QUOTES_FOLDER, exist_ok=True)

VALID_TABLES = None

COUNTER_FILE = os.path.join(BASE_DIR, "quote_counter.txt")

def next_quote_seq():
    if os.path.exists(COUNTER_FILE):
        with open(COUNTER_FILE) as f:
            n = int(f.read().strip()) + 1
    else:
        n = 1
    with open(COUNTER_FILE, "w") as f:
        f.write(str(n))
    return str(n).zfill(3)

def ensure_log_table():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""CREATE TABLE IF NOT EXISTS table_update_log (
        table_name TEXT PRIMARY KEY, updated_at TEXT, uploaded_by TEXT)""")
    conn.commit()
    conn.close()

ensure_log_table()


def ensure_valve_sizes():
    """Ensure rotary joint table has entries up to 600 NB.
    (butterfly_valve_master was dropped — selector now reads from
    component_price_master and lt_butterfly_valve_master.)"""
    conn = sqlite3.connect(DB_PATH)
    rj_defaults = [(400, 65000), (450, 75000), (500, 90000), (600, 110000)]
    try:
        for nb, price in rj_defaults:
            conn.execute("INSERT OR IGNORE INTO rotary_joint_master (nb, price) VALUES (?,?)", (nb, price))
        conn.commit()
    except sqlite3.OperationalError:
        pass  # rotary_joint_master may not exist yet on a fresh DB
    conn.close()

ensure_valve_sizes()


def ensure_extra_columns():
    """Add updated_at and company columns to component_price_master if not present."""
    conn = sqlite3.connect(DB_PATH)
    cols = [r[1] for r in conn.execute("PRAGMA table_info(component_price_master)").fetchall()]
    if "updated_at" not in cols:
        conn.execute("ALTER TABLE component_price_master ADD COLUMN updated_at TEXT")
    if "company" not in cols:
        conn.execute("ALTER TABLE component_price_master ADD COLUMN company TEXT")
    conn.commit()
    conn.close()

ensure_extra_columns()

# Duplicate name pairs in component_price_master — canonical -> [aliases]
_RATE_DUPLICATES = [
    ("FLEXIBLE HOSE-15NB*1000MM (OIL)", ["FLEXIBLE HOSE-15NB*1000MM (OIL )"]),
    ("FLEXIBLE HOSE-15NB*750MM (OIL)",  ["FLEXIBLE HOSE-15NB*750MM (OIL )"]),
    ("FLEXIBLE HOSE-20NB*1000MM (OIL)", ["FLEXIBLE HOSE-20NB*1000MM (OIL )"]),
    ("FLEXIBLE HOSE-25NB*1000MM (AIR)", ["FLEXIBLE HOSE-25NB*1000MM (AIR )"]),
    ("M.S. Sheet 2mm",       ["M.S. Sheet  2mm"]),
    ("M.S. Sheet 3mm",       ["M.S. Sheet  3mm"]),
    ("M.S. Plate 16mm*5mm",  ["M.S. Plate 16mm* 5mm"]),
    ("M.S. Tube B Class 1.5 in", ['M.S. Tube "B" Class 1.5 in']),
    ("M.S. Tube C Class 1.5 in", ['M.S. Tube "C" Class 1.5 in']),
    ("M.S. Chanel",          ["M.S.Chanel"]),
    ("Plumber block with Bearing", ["Plumber Block with Bearing"]),
    ("Pulley with V belt",   ["Pulley with V Belt"]),
    ("SS Pipe 304 60x3mm (per mtr)",  ["SS Pipe 304 60x3mm",  "SS Pipe 304 60 X 3mm"]),
    ("SS Pipe 304 76x3mm (per mtr)",  ["SS Pipe 304 76x3mm",  "SS Pipe 304 76 X 3mm"]),
    ("SS Pipe 304 100x3mm (per mtr)", ["SS Pipe 304 100x3mm", "SS Pipe 304 100 X 3mm"]),
    ("ID FAN (ARE 35)",      ["ID FAN  (ARE 35)"]),
    ("SEQUENCE CONTROLLER",  ["SEQUENCE"]),
]

def clean_duplicate_rates(conn):
    """Remove known duplicate/alias rows from component_price_master."""
    for canonical, aliases in _RATE_DUPLICATES:
        for alias in aliases:
            exists = conn.execute(
                "SELECT 1 FROM component_price_master WHERE item=?", (alias,)
            ).fetchone()
            if not exists:
                continue
            canonical_exists = conn.execute(
                "SELECT 1 FROM component_price_master WHERE item=?", (canonical,)
            ).fetchone()
            if canonical_exists:
                conn.execute("DELETE FROM component_price_master WHERE item=?", (alias,))
            else:
                conn.execute(
                    "UPDATE component_price_master SET item=? WHERE item=?", (canonical, alias)
                )
    conn.commit()

import glob
import tempfile

ALLOWED_EDIT_TABLES = {
    'hpu_master', 'oil_burner_parts_master', 'hv_oil_burner_parts_master',
    'gas_burner_parts_master', 'horizontal_master', 'vertical_master',
    'recuperator_master', 'blower_pricelist_master',
    'rad_heat_master', 'rad_heat_tata_master', 'gail_gas_burner_master',
}

def _find_latest_pricebook():
    """Find the most recently uploaded full pricebook Excel file."""
    candidates = []
    search_dirs = [UPLOAD_FOLDER, BASE_DIR]
    for d in search_dirs:
        for f in glob.glob(os.path.join(d, "*.xlsx")):
            # Skip stock files and other non-pricebook files
            bn = os.path.basename(f).lower()
            if "stock" in bn or "costing" in bn or "regen" in bn or "sample" in bn:
                continue
            try:
                xl = pd.ExcelFile(f)
                if len(xl.sheet_names) >= 8:
                    candidates.append((os.path.getmtime(f), f))
            except Exception:
                pass
    return sorted(candidates, reverse=True)[0][1] if candidates else None

def _find_regen_file():
    """Find the Regen Standard Costing workbook in BASE_DIR."""
    for f in glob.glob(os.path.join(BASE_DIR, "*.xlsx")):
        bn = os.path.basename(f).lower()
        if "regen" in bn and "costing" in bn:
            return f
    return None

def _ensure_regen_costing():
    """Parse the regen costing file if found and table not yet populated."""
    try:
        regen_file = _find_regen_file()
        if not regen_file:
            return
        conn = sqlite3.connect(DB_PATH)
        try:
            count = conn.execute("SELECT COUNT(*) FROM regen_costing_items").fetchone()[0]
            if count > 0:
                conn.close()
                return
        except Exception:
            pass
        from bom.regen_parser import parse_regen_costing
        parse_regen_costing(regen_file, conn)
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"Regen costing parse error: {e}")

_ensure_regen_costing()

def _cascade_recalculate(xl_path: str, conn):
    """Patch the Rates sheet with current DB rates, re-run all parsers."""
    import openpyxl
    # Get current rates with cell positions
    rate_cells = {}
    for row in conn.execute(
        "SELECT price, excel_row, excel_col FROM component_price_master WHERE excel_row IS NOT NULL AND excel_col IS NOT NULL"
    ):
        price, er, ec = row
        if er and ec:
            rate_cells[(int(er), int(ec))] = price

    wb = openpyxl.load_workbook(xl_path, data_only=False)
    rates_sn = next((s for s in wb.sheetnames if s.strip().lower() == 'rates'), None)
    if rates_sn and rate_cells:
        ws = wb[rates_sn]
        for (er, ec), price in rate_cells.items():
            ws.cell(er, ec).value = price

    tmp_fd, tmp_path = tempfile.mkstemp(suffix='.xlsx')
    os.close(tmp_fd)
    wb.save(tmp_path)
    wb.close()

    results = _parse_pricelist_all(tmp_path, conn)
    try:
        os.unlink(tmp_path)
    except Exception:
        pass
    return results

def ensure_rate_columns():
    conn = sqlite3.connect(DB_PATH)
    for col in ['excel_row', 'excel_col']:
        try:
            conn.execute(f"ALTER TABLE component_price_master ADD COLUMN {col} INTEGER")
            conn.commit()
        except Exception:
            pass
    conn.close()

ensure_rate_columns()

@app.get("/", response_class=HTMLResponse)
def root():
    html_path = os.path.join(BASE_DIR, "dashboard.html")
    with open(html_path, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())


@app.get("/api/last-pricebook-update")
def last_pricebook_update():
    try:
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("SELECT updated_at FROM table_update_log ORDER BY updated_at DESC LIMIT 1")
        row = c.fetchone()
        conn.close()
        if row:
            dt = datetime.fromisoformat(row[0])
            return {"date": dt.strftime("%d %b %Y")}
        return {"date": None}
    except Exception:
        return {"date": None}


@app.get("/api/system-health")
def system_health():
    """
    Returns health status for each key master table plus recent activity log.
    Used by the dashboard to show a product manager what data is loaded.
    """
    KEY_TABLES = [
        ("component_price_master", "Rates / Component Prices",      "Rates sheet"),
        ("hpu_master",             "HPU Components",                 "HPU sheet"),
        ("burner_pricelist_master","Burner Selling Prices",          "BURNER sheet"),
        ("blower_master",          "Blower Models & Prices",         "Blower sheet"),
        ("vertical_master",        "VLPH Structure Costs",           "Vertical sheet"),
        ("horizontal_master",      "HLPH Structure Costs",           "Horizontal sheet"),
        ("gas_burner_parts_master","Gas Burner Fabrication Parts",   "Gas Burner sheet"),
        ("gail_gas_burner_master", "GAIL Burner Prices",             "GAIL GAS Burner sheet"),
        ("ng_gas_train_master",    "Gas Train Models",               "manual / init_db"),
        ("agr_master",             "AGR Models",                     "manual / init_db"),
        ("rotary_joint_master",    "Rotary Joint Models",            "manual / init_db"),
        ("blower_pricelist_master","Blower Pricelist (raw)",         "Blower sheet"),
        ("recuperator_master",     "Recuperator Models",             "Recuperator sheet"),
    ]

    try:
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()

        c.execute("SELECT table_name, updated_at FROM table_update_log")
        log = {r[0]: r[1] for r in c.fetchall()}

        tables = []
        for table, label, source in KEY_TABLES:
            try:
                c.execute(f"SELECT COUNT(*) FROM [{table}]")
                count = c.fetchone()[0]
            except Exception:
                count = 0
            updated_at = log.get(table)
            if updated_at:
                dt = datetime.fromisoformat(updated_at)
                updated_str = dt.strftime("%d %b %Y, %H:%M")
            else:
                updated_str = None
            tables.append({
                "table":      table,
                "label":      label,
                "source":     source,
                "rows":       count,
                "ok":         count > 0,
                "updated_at": updated_str,
            })

        # Recent activity: last 8 table updates
        c.execute("SELECT table_name, updated_at FROM table_update_log ORDER BY updated_at DESC LIMIT 8")
        activity = []
        for tname, upd in c.fetchall():
            try:
                dt = datetime.fromisoformat(upd)
                activity.append({"table": tname, "at": dt.strftime("%d %b %Y, %H:%M")})
            except Exception:
                pass

        conn.close()
        return {"tables": tables, "activity": activity}
    except Exception as e:
        return {"error": str(e), "tables": [], "activity": []}

@app.get("/quote", response_class=HTMLResponse)
def quote_form():
    html_path = os.path.join(BASE_DIR, "quote_form.html")
    with open(html_path, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())

@app.get("/viewer", response_class=HTMLResponse)
def db_viewer():
    html_path = os.path.join(BASE_DIR, "db_viewer.html")
    with open(html_path, "r", encoding="utf-8") as f:
        content = f.read()
    content = content.replace('const API = "http://127.0.0.1:8000"', 'const API = ""')
    return HTMLResponse(content=content)


@app.get("/api/catalog")
def get_catalog():
    try:
        conn = sqlite3.connect(DB_PATH)
        def q(sql):
            return pd.read_sql(sql, conn).to_dict(orient="records")
        # Price master items (if table exists)
        pm_raw, pm_bo, pm_ep = [], [], []
        try:
            pm_raw = q("SELECT item as id, item as label, price as base_price, unit FROM component_price_master WHERE category='Raw Material' ORDER BY item")
            pm_bo  = q("SELECT item as id, item as label, price as base_price, unit FROM component_price_master WHERE category='Bought Out' ORDER BY item")
            pm_ep  = q("SELECT item as id, item as label, price as base_price, unit FROM component_price_master WHERE category='ENCON Purchase' ORDER BY item")
        except Exception:
            pass

        catalog = {
            "Horizontal Ladle Preheater": q("SELECT DISTINCT model as id, model as label FROM horizontal_master"),
            "Vertical Ladle Preheater": q("SELECT DISTINCT model as id, model as label FROM vertical_master"),
            "Blower": q("SELECT DISTINCT model as id, model as label, section, price_without_motor, price_with_motor, hp FROM blower_pricelist_master WHERE price_without_motor IS NOT NULL ORDER BY section, hp"),
            "HPU": q("SELECT DISTINCT unit_kw as id, CAST(unit_kw AS TEXT) || ' KW - ' || variant as label, unit_kw, variant FROM hpu_master ORDER BY unit_kw"),
            "Burner (Film)": q("SELECT DISTINCT burner_size as id, burner_size as label, price as base_price FROM burner_pricelist_master WHERE component='BURNER ALONE' AND section LIKE '%FILM%' GROUP BY burner_size"),
            "Burner (Dual Fuel)": q("SELECT DISTINCT burner_size as id, burner_size as label, price as base_price FROM burner_pricelist_master WHERE component='BURNER ALONE' AND section LIKE '%DUAL%' GROUP BY burner_size"),
            "Recuperator": q("SELECT DISTINCT model as id, type || ' - ' || model as label FROM recuperator_master"),
            "Rad Heat": q("SELECT item as id, item || ' (' || output_kw || ')' as label, price_with_ss_tubing as base_price FROM rad_heat_master WHERE section='MODEL'"),
            "GAIL Gas Burner": q("SELECT burner_size as id, burner_size as label, burner_set as base_price FROM gail_gas_burner_master WHERE burner_set IS NOT NULL"),
            "Raw Material": pm_raw,
            "Bought Out Item": pm_bo,
            "ENCON Component": pm_ep,
        }
        conn.close()
        return catalog
    except Exception as e:
        return {"error": str(e)}

@app.get("/api/price")
def get_price(product_type: str, model: str, qty: int = 1,
              with_motor: bool = False, variant: str = "Duplex 1"):
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        price = 0
        breakdown = []

        if product_type == "Horizontal Ladle Preheater":
            cursor.execute("SELECT particular, amount FROM horizontal_master WHERE model=? AND amount IS NOT NULL", (model,))
            for particular, amount in cursor.fetchall():
                if particular not in ("COMBUSTION EQUIPMENT:", "S.NO.") and amount:
                    breakdown.append({"item": particular, "amount": float(amount)})
                    price += float(amount)

        elif product_type == "Vertical Ladle Preheater":
            cursor.execute("SELECT particular, amount FROM vertical_master WHERE model=? AND amount IS NOT NULL", (model,))
            for particular, amount in cursor.fetchall():
                if particular not in ("COMBUSTION EQUIPMENT:", "S.NO.") and amount:
                    breakdown.append({"item": particular, "amount": float(amount)})
                    price += float(amount)

        elif product_type == "Blower":
            col = "price_with_motor" if with_motor else "price_without_motor"
            cursor.execute(f"SELECT {col} FROM blower_pricelist_master WHERE model=? AND {col} IS NOT NULL LIMIT 1", (model,))
            row = cursor.fetchone()
            if row and row[0]:
                price = float(row[0])
                breakdown = [{"item": f"{model} ({'with' if with_motor else 'without'} motor)", "amount": price}]

        elif product_type == "HPU":
            cursor.execute("SELECT SUM(amount) FROM hpu_master WHERE unit_kw=? AND variant=? AND amount IS NOT NULL", (model, variant))
            row = cursor.fetchone()
            if row and row[0]:
                price = float(row[0])
                breakdown = [{"item": f"HPU {model} KW ({variant})", "amount": price}]

        elif "Burner" in product_type:
            section_filter = "FILM" if "Film" in product_type else "DUAL"
            cursor.execute("SELECT price FROM burner_pricelist_master WHERE burner_size=? AND component='BURNER ALONE' AND section LIKE ? LIMIT 1", (model, f"%{section_filter}%"))
            row = cursor.fetchone()
            if row and row[0]:
                price = float(row[0])
                breakdown = [{"item": f"{model} Burner", "amount": price}]

        elif product_type == "Rad Heat":
            cursor.execute("SELECT price_with_ss_tubing FROM rad_heat_master WHERE item=? AND section='MODEL' LIMIT 1", (model,))
            row = cursor.fetchone()
            if row and row[0]:
                price = float(row[0])
                breakdown = [{"item": f"Rad Heat {model}", "amount": price}]

        elif product_type == "GAIL Gas Burner":
            cursor.execute("SELECT burner_set FROM gail_gas_burner_master WHERE burner_size=? LIMIT 1", (model,))
            row = cursor.fetchone()
            if row and row[0]:
                price = float(row[0])
                breakdown = [{"item": f"GAIL Gas Burner {model} Set", "amount": price}]

        elif product_type in ("Raw Material", "Bought Out Item", "ENCON Component"):
            cursor.execute("SELECT price, unit FROM component_price_master WHERE item=? LIMIT 1", (model,))
            row = cursor.fetchone()
            if row and row[0]:
                price = float(row[0])
                unit = row[1] or "nos"
                breakdown = [{"item": f"{model} ({unit})", "amount": price}]

        conn.close()
        return {"unit_price": price, "qty": qty, "total": price * qty, "breakdown": breakdown}
    except Exception as e:
        return {"error": str(e)}


class RateUpdateRequest(BaseModel):
    item: str
    price: float

class CompanyUpdateRequest(BaseModel):
    item: str
    company: str

class ItemUpdateRequest(BaseModel):
    table: str
    rowid: int
    qty: Optional[float] = None
    rate: Optional[float] = None

class VLPHCalcRequest(BaseModel):
    mode: str = "calc"                          # "calc" or "direct"
    Ti: float = 650.0
    Tf: float = 1200.0
    refractory_weight: float = 21500.0
    fuel_cv: float = 8500.0
    time_taken_hr: float = 2.0
    refractory_heat_factor: float = 0.25
    efficiency: float = 0.52
    ladle_tons: float = 10.0
    fuel1_type: str = "ng"
    fuel1_cv: float = 8500.0
    fuel2_type: str = "none"
    fuel2_cv: float = 0.0
    direct_burner_capacity: float = 0.0         # Nm3/hr (direct mode)
    blower_pressure: str = "28"                  # "28" or "40" (WG inches)
    control_mode: str = "automatic"              # "manual" or "automatic"
    auto_control_type: str = "agr"               # "plc", "plc_agr", "pid"
    control_valve_vendor: str = "dembla"         # "dembla" or "cair"
    shutoff_valve_vendor: str = "aira"           # "dembla", "aira", or "cair" — shut off valve vendor
    butterfly_valve_vendor: str = "lt_lever"     # "lt_lever" or "lt_gear" (L&T butterfly variants)
    pressure_gauge_vendor: str = "baumer"        # "baumer" or "hguru"
    hpu_variant: str = "Duplex 1"                # "Simplex" | "Duplex 1" | "Duplex 2" — for oil fuels
    # burner_pressure_wg is derived from blower_pressure (28→24, 40→36)
    pilot_burner: str = "auto"                   # "auto" | "lpg_10" | "nglpg_100" | "cog_100"
    pipeline_weight_kg: float = 1000.0           # Air-gas pipeline weight (700–2000 kg, step 100)
    purging_line: str = "no"                     # "yes" | "no" — nitrogen purging line for MG/COG
    manual_pilot_burner: str = "yes"             # "yes" | "no" — include pilot burner in manual BOM
    pilot_line_fuel: str = "lpg"                 # "lpg" | "ng" — pilot line fuel type (manual mode)


class QuoteItem(BaseModel):
    product_type: str
    model: str
    description: Optional[str] = None
    qty: int = 1
    with_motor: bool = False
    variant: str = "Duplex 1"
    unit_price: Optional[float] = None   # pre-set price (e.g. from VLPH costing)
    total: Optional[float] = None

class QuoteRequest(BaseModel):
    # Company / customer
    company_name: str
    company_address: Optional[str] = ""
    company_city: Optional[str] = ""
    company_state: Optional[str] = ""
    company_pin: Optional[str] = ""
    company_gstin: Optional[str] = ""
    # Point of contact
    poc_name: Optional[str] = ""
    poc_designation: Optional[str] = ""
    mobile_no: Optional[str] = ""
    email: Optional[str] = ""
    # Enquiry
    project_name: Optional[str] = ""
    ref_no: Optional[str] = ""
    # Items & commercial
    items: List[QuoteItem]
    gst_percent: float = 18
    freight: float = 0
    valid_days: int = 30


@app.get("/costing", response_class=HTMLResponse)
def costing_form():
    html_path = os.path.join(BASE_DIR, "vlph_costing.html")
    with open(html_path, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())

@app.get("/hlph", response_class=HTMLResponse)
def hlph_costing_form():
    html_path = os.path.join(BASE_DIR, "hlph_costing.html")
    with open(html_path, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())

@app.get("/regen", response_class=HTMLResponse)
def regen_costing_form():
    html_path = os.path.join(BASE_DIR, "regen_costing.html")
    with open(html_path, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())

@app.get("/price-master", response_class=HTMLResponse)
def price_master_page():
    html_path = os.path.join(BASE_DIR, "price_master.html")
    with open(html_path, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())


@app.get("/pricelist", response_class=HTMLResponse)
def pricelist_viewer_page():
    html_path = os.path.join(BASE_DIR, "pricelist_viewer.html")
    with open(html_path, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())


@app.get("/api/pricelist-summary")
def pricelist_summary():
    """Return all live-computed prices from every master table."""
    try:
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()

        def q(sql, *args):
            return c.execute(sql, args).fetchall()

        def _parts_sections(table, markup=1.25):
            """Convert section-header / item rows into [{title, items, fabrication_cost, selling_price}]."""
            sections, cur_title, cur_items, cur_total = [], None, [], 0.0
            for rowid, _, part, qty, unit, rate, amt in conn.execute(
                f"SELECT rowid, section, particular, qty, unit, rate, amount FROM {table} ORDER BY rowid"
            ):
                if amt is None:
                    if cur_title is not None:
                        sections.append({"title": cur_title, "items": cur_items,
                                         "fabrication_cost": round(cur_total, 2),
                                         "selling_price": round(cur_total * markup, 2)})
                    cur_title, cur_items, cur_total = part, [], 0.0
                else:
                    cur_items.append({"rowid": rowid, "particular": part, "qty": qty, "unit": unit,
                                      "rate": rate, "amount": amt})
                    cur_total += amt or 0
            if cur_title is not None:
                sections.append({"title": cur_title, "items": cur_items,
                                 "fabrication_cost": round(cur_total, 2),
                                 "selling_price": round(cur_total * markup, 2)})
            return sections

        # ── HPU ───────────────────────────────────────────────────────────
        _variant_order = "CASE variant WHEN 'Duplex 1' THEN 1 WHEN 'Duplex 2' THEN 2 WHEN 'Simplex' THEN 3 ELSE 4 END"
        hpu_kws = [r[0] for r in q("SELECT DISTINCT unit_kw FROM hpu_master ORDER BY unit_kw")]
        hpu = []
        for kw in hpu_kws:
            variants_data = []
            for (variant,) in q(f"SELECT DISTINCT variant FROM hpu_master WHERE unit_kw=? ORDER BY {_variant_order}", kw):
                rows = q("SELECT rowid, item, qty, unit, rate, amount FROM hpu_master WHERE unit_kw=? AND variant=? ORDER BY rowid", kw, variant)
                items = [{"rowid": r[0], "item": r[1], "qty": r[2], "unit": r[3], "rate": r[4], "amount": r[5]} for r in rows]
                mat = sum((r[5] or 0) for r in rows)
                variants_data.append({"name": variant, "material_cost": round(mat, 2),
                                      "selling_price": round(mat * 1.8, 2), "items": items})
            hpu.append({"kw": kw, "variants": variants_data})

        # ── Burners ───────────────────────────────────────────────────────
        burner_rows = q("SELECT section, burner_size, component, price FROM burner_pricelist_master ORDER BY section, burner_size")
        burners = {}
        for sec, size, comp, price in burner_rows:
            burners.setdefault(sec, {}).setdefault(size, {})[comp] = price

        # ── Blowers ───────────────────────────────────────────────────────
        blower_cols = [d[0] for d in c.execute("SELECT * FROM blower_pricelist_master LIMIT 0").description]
        blowers = [dict(zip(blower_cols, r)) for r in q("SELECT * FROM blower_pricelist_master ORDER BY section, CAST(hp AS REAL), model")]
        try:
            dm_cols = [d[0] for d in c.execute("SELECT * FROM blower_dm_idm_master LIMIT 0").description]
            blower_dm_idm = [dict(zip(dm_cols, r)) for r in q("SELECT * FROM blower_dm_idm_master ORDER BY section, CAST(SUBSTR(model,1,INSTR(model,'/')-1) AS REAL)")]
        except Exception:
            blower_dm_idm = []

        # ── Recuperator ───────────────────────────────────────────────────
        try:
            # Use cursor description (avoids variable-name clash with outer cursor 'c')
            c.execute("SELECT * FROM recuperator_master ORDER BY type, CAST(model AS REAL)")
            rcols = [d[0] for d in c.description]
            recup_rows = [dict(zip(rcols, r)) for r in c.fetchall()]
            # Fetch component rates used in recuperator calculations
            recup_rate_items = {
                'tube_c':     "M.S. Tube \"C\" Class",
                'tube_b':     "M.S. Tube \"B\" Class",
                'ci_gills':   "C.I. Gills",
                'ang_6550':   "M.S. Angle 65,50",
                'ang_100100': "M.S. Angle 100,100",
                'plate_16_5': "M.S. Plate 16mm*",
                'bolt':       "Hardware Bolt",
                'channel':    "M.S. Chanel",
                'plate_16_10':"M.S. Plate 16mm*10mm",
                'plate_5':    "M.S. Plate 5mm",
                'ang_50':     "M.S. Angle 50*6",
                'ss_sheet':   "S.S. Sheet 3mm",
            }
            recup_rates = {}
            for key, prefix in recup_rate_items.items():
                row = conn.execute(
                    "SELECT price FROM component_price_master WHERE item LIKE ? LIMIT 1",
                    (prefix + '%',)
                ).fetchone()
                recup_rates[key] = float(row[0]) if row else None
            recup = {"rows": recup_rows, "rates": recup_rates}
        except Exception:
            recup = {"rows": [], "rates": {}}

        # ── Rad Heat ──────────────────────────────────────────────────────
        def _rad_rows(table):
            models, spares = [], []
            for r in q(f"SELECT * FROM {table} ORDER BY section, item"):
                d = dict(zip([x[0] for x in conn.execute(f"SELECT * FROM {table} LIMIT 0").description], r))
                if d["section"] == "MODEL":
                    models.append(d)
                else:
                    spares.append({"item": d["item"], "price": d["price_with_ss_tubing"]})
            return {"models": models, "spares": spares}
        rad      = _rad_rows("rad_heat_master")
        rad_tata = _rad_rows("rad_heat_tata_master")

        # ── GAIL Gas Burner ───────────────────────────────────────────────
        gail_cols = [d[0] for d in c.execute("SELECT * FROM gail_gas_burner_master LIMIT 0").description]
        gail = [dict(zip(gail_cols, r)) for r in q("SELECT * FROM gail_gas_burner_master ORDER BY section, burner_size")]

        # ── Burner Parts (Oil / HV Oil / Gas) ─────────────────────────────
        oil_parts = _parts_sections("oil_burner_parts_master",    markup=1.25)
        hv_parts  = _parts_sections("hv_oil_burner_parts_master", markup=1.25)
        gas_parts = _parts_sections("gas_burner_parts_master",    markup=1.25)

        # ── Horizontal LPH ────────────────────────────────────────────────
        hlph = []
        for (model,) in q("SELECT DISTINCT model FROM horizontal_master ORDER BY model"):
            rows = q("SELECT particular, qty, amount FROM horizontal_master WHERE model=?", model)
            items = [{"particular": r[0], "qty": r[1], "amount": r[2]} for r in rows if r[2]]
            if items:
                hlph.append({"model": model, "total": round(sum(i["amount"] for i in items), 2), "items": items})

        # ── Vertical LPH ─────────────────────────────────────────────────
        vlph = []
        for (model,) in q("SELECT DISTINCT model FROM vertical_master ORDER BY model"):
            rows = q("SELECT particular, qty, amount FROM vertical_master WHERE model=?", model)
            items = [{"particular": r[0], "qty": r[1], "amount": r[2]} for r in rows if r[2]]
            if items:
                vlph.append({"model": model, "total": round(sum(i["amount"] for i in items), 2), "items": items})

        # ── Regen Costing ──────────────────────────────────────────────────
        try:
            rci_cols = [d[0] for d in conn.execute("SELECT * FROM regen_costing_items LIMIT 0").description]
            rci_rows = [dict(zip(rci_cols, r)) for r in conn.execute(
                "SELECT * FROM regen_costing_items ORDER BY kw, row_order"
            ).fetchall()]
            rsz_cols = [d[0] for d in conn.execute("SELECT * FROM regen_sizing LIMIT 0").description]
            rsz_rows = [dict(zip(rsz_cols, r)) for r in conn.execute(
                "SELECT * FROM regen_sizing ORDER BY kw"
            ).fetchall()]
            rpl_cols = [d[0] for d in conn.execute("SELECT * FROM regen_pricelist LIMIT 0").description]
            rpl_rows = [dict(zip(rpl_cols, r)) for r in conn.execute(
                "SELECT * FROM regen_pricelist ORDER BY kw"
            ).fetchall()]
            rmr_cols = [d[0] for d in conn.execute("SELECT * FROM regen_material_rates LIMIT 0").description]
            rmr_rows = [dict(zip(rmr_cols, r)) for r in conn.execute(
                "SELECT * FROM regen_material_rates"
            ).fetchall()]
            rnz_cols = [d[0] for d in conn.execute("SELECT * FROM regen_nozzle_sizing LIMIT 0").description]
            rnz_rows = [dict(zip(rnz_cols, r)) for r in conn.execute(
                "SELECT * FROM regen_nozzle_sizing ORDER BY power_kw"
            ).fetchall()]
            rps_cols = [d[0] for d in conn.execute("SELECT * FROM regen_pipe_sizes LIMIT 0").description]
            rps_rows = [dict(zip(rps_cols, r)) for r in conn.execute(
                "SELECT * FROM regen_pipe_sizes ORDER BY gas_type, burner_size_kw"
            ).fetchall()]
            regen_costing = {
                "items": rci_rows, "sizing": rsz_rows,
                "pricelist": rpl_rows, "material_rates": rmr_rows,
                "nozzle_sizing": rnz_rows, "pipe_sizes": rps_rows,
            }
        except Exception:
            regen_costing = {"items": [], "sizing": [], "pricelist": [], "material_rates": [], "nozzle_sizing": [], "pipe_sizes": []}

        conn.close()
        return {
            "hpu": hpu,
            "burners": burners,
            "blowers": blowers,
            "blower_dm_idm": blower_dm_idm,
            "recuperator": recup,
            "rad_heat": rad,
            "rad_heat_tata": rad_tata,
            "gail_gas_burner": gail,
            "oil_burner_parts": oil_parts,
            "hv_oil_burner_parts": hv_parts,
            "gas_burner_parts": gas_parts,
            "horizontal_lph": hlph,
            "vertical_lph": vlph,
            "regen_costing": regen_costing,
        }
    except Exception as e:
        import traceback
        return {"error": str(e), "detail": traceback.format_exc()}


@app.get("/api/pricelist/rates")
def get_pricelist_rates():
    """Return component_price_master for editing."""
    try:
        conn = sqlite3.connect(DB_PATH)
        rows = conn.execute(
            "SELECT rowid, item, category, price, previous_price, updated_at, company FROM component_price_master ORDER BY category, item"
        ).fetchall()
        conn.close()
        return [{"rowid": r[0], "item": r[1], "category": r[2],
                 "price": r[3], "previous_price": r[4], "updated_at": r[5], "company": r[6]} for r in rows]
    except Exception as e:
        return {"error": str(e)}


@app.put("/api/pricelist/rate")
def update_pricelist_rate(req: RateUpdateRequest):
    """Update a rate in component_price_master and cascade-recalculate all tables."""
    import re as _re

    def _norm(s):
        s = str(s or "").upper()
        s = _re.sub(r"[\s\-_\(\)\*\.\,\"\']+", " ", s).strip()
        s = _re.sub(r"\s+", " ", s)  # collapse multiple spaces
        # Also create a no-space version for comparison
        return s

    def _norm_compact(s):
        """Extra compact normalization — removes ALL spaces and dots."""
        s = str(s or "").upper()
        s = _re.sub(r"[^A-Z0-9]", "", s)
        return s

    try:
        conn = sqlite3.connect(DB_PATH)

        # Get old price before update
        old_row = conn.execute(
            "SELECT price FROM component_price_master WHERE item=?", (req.item,)
        ).fetchone()
        old_price = old_row[0] if old_row else None

        # Update master table
        conn.execute(
            "UPDATE component_price_master SET previous_price=price, price=?, updated_at=? WHERE item=?",
            (req.price, datetime.now().strftime("%Y-%m-%d %H:%M"), req.item)
        )
        conn.commit()

        # ── Direct cascade to all parts tables ──────────────────────────────
        # Update rows where name fuzzy-matches AND rate = old_price
        PARTS_TABLES = [
            ("oil_burner_parts_master",    "particular"),
            ("hv_oil_burner_parts_master", "particular"),
            ("gas_burner_parts_master",    "particular"),
            ("hpu_master",                 "item"),
        ]
        norm_item = _norm(req.item)
        compact_item = _norm_compact(req.item)
        cascade_counts = {}

        # ── Method 1: Use formula mapping table (exact legacy links) ──
        mapped_targets = conn.execute(
            "SELECT target_table, target_item FROM rate_cascade_map WHERE rates_item = ?",
            (req.item,)
        ).fetchall()

        for target_table, target_item in mapped_targets:
            name_col = "particular" if "parts" in target_table else "item"
            rows = conn.execute(
                f"SELECT rowid, {name_col}, qty FROM {target_table} WHERE {name_col} = ?",
                (target_item,)
            ).fetchall()
            updated = 0
            for rowid, part_name, qty in rows:
                new_amount = round(float(qty or 0) * req.price, 2) if qty is not None else None
                conn.execute(
                    f"UPDATE {target_table} SET rate=?, amount=? WHERE rowid=?",
                    (req.price, new_amount, rowid)
                )
                updated += 1
            if updated:
                cascade_counts[target_table] = cascade_counts.get(target_table, 0) + updated

        # ── Method 2: Fuzzy name match (for items not in mapping table) ──
        for table, name_col in PARTS_TABLES:
            rows = conn.execute(
                f"SELECT rowid, {name_col}, qty, rate FROM {table} WHERE rate IS NOT NULL"
            ).fetchall()
            updated = 0
            for rowid, part_name, qty, rate in rows:
                norm_part = _norm(part_name or "")
                compact_part = _norm_compact(part_name or "")
                if norm_part != norm_item and compact_part != compact_item:
                    continue
                new_amount = round(float(qty or 0) * req.price, 2) if qty is not None else None
                conn.execute(
                    f"UPDATE {table} SET rate=?, amount=? WHERE rowid=?",
                    (req.price, new_amount, rowid)
                )
                updated += 1
            if updated:
                cascade_counts[table] = cascade_counts.get(table, 0) + updated

        conn.commit()

        conn.close()
        return {
            "success": True,
            "cascaded": True,
            "direct_cascade": cascade_counts,
        }
    except Exception as e:
        import traceback
        return {"success": False, "error": str(e), "detail": traceback.format_exc()}


@app.put("/api/pricelist/company")
def update_pricelist_company(req: CompanyUpdateRequest):
    """Update the company/supplier name for an item in component_price_master."""
    try:
        conn = sqlite3.connect(DB_PATH)
        conn.execute(
            "UPDATE component_price_master SET company=? WHERE item=?",
            (req.company.strip() or None, req.item)
        )
        conn.commit()
        conn.close()
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}


@app.put("/api/pricelist/item")
def update_pricelist_item(req: ItemUpdateRequest):
    """Update qty and/or rate for a row in any master table; recalculate amount."""
    if req.table not in ALLOWED_EDIT_TABLES:
        return {"success": False, "error": f"Table '{req.table}' not editable"}
    try:
        conn = sqlite3.connect(DB_PATH)
        if req.qty is not None:
            conn.execute(f"UPDATE {req.table} SET qty=? WHERE rowid=?", (req.qty, req.rowid))
        if req.rate is not None:
            conn.execute(f"UPDATE {req.table} SET rate=? WHERE rowid=?", (req.rate, req.rowid))
        # Recalculate amount
        row = conn.execute(f"SELECT qty, rate FROM {req.table} WHERE rowid=?", (req.rowid,)).fetchone()
        if row and row[0] is not None and row[1] is not None:
            amount = round(float(row[0]) * float(row[1]), 2)
            conn.execute(f"UPDATE {req.table} SET amount=? WHERE rowid=?", (amount, req.rowid))
        conn.commit()
        row2 = conn.execute(f"SELECT qty, rate, amount FROM {req.table} WHERE rowid=?", (req.rowid,)).fetchone()
        conn.close()
        return {"success": True, "qty": row2[0], "rate": row2[1], "amount": row2[2]}
    except Exception as e:
        return {"success": False, "error": str(e)}


@app.post("/api/vlph-calculate")
def vlph_calculate(req: VLPHCalcRequest):
    try:
        from calculations.burner import BurnerInputs, calculate_burner
        from calculations.pipes import PipeInputs, calculate_pipe_sizes, select_oil_pipe_nb
        from bom.selectors.selection_engine import select_equipment
        from bom.vlph_builder import build_vlph_120t_df, build_vlph_manual_df

        FUEL_NAMES = {
            "ng": "Natural Gas", "lpg": "LPG", "cog": "COG", "bg": "BFG", "rlng": "RLNG", "mg": "Mixed Gas",
            "hsd": "Diesel (HSD)", "ldo": "LDO", "hdo": "HDO", "fo": "Furnace Oil",
            "sko": "Kerosene (SKO)", "cfo": "CFO", "lshs": "LSHS",
        }
        OIL_FUELS = {"hsd", "ldo", "hdo", "fo", "sko", "cfo", "lshs"}

        f1_cv = req.fuel1_cv if req.fuel1_cv > 0 else req.fuel_cv

        if req.mode == "direct":
            # --- Direct mode: burner capacity = higher-CV fuel flow ---
            # Heat output from the entered capacity (using higher CV fuel)
            is_dual = req.fuel2_type != "none" and req.fuel2_cv > 0
            if is_dual:
                higher_cv = max(f1_cv, req.fuel2_cv)
            else:
                higher_cv = f1_cv
            heat_kcal_hr = req.direct_burner_capacity * higher_cv
            # Each fuel's flow = heat / its CV
            ng_flow = heat_kcal_hr / f1_cv
            air_flow = heat_kcal_hr * 118 / 100000
            br1 = None
        else:
            # --- Calc mode: calculate from process params ---
            br1 = calculate_burner(BurnerInputs(
                Ti=req.Ti, Tf=req.Tf,
                refractory_weight=req.refractory_weight,
                fuel_cv=f1_cv,
                time_taken_hr=req.time_taken_hr,
                refractory_heat_factor=req.refractory_heat_factor,
                efficiency=req.efficiency,
            ))
            ng_flow = br1.extra_firing_rate_nm3hr
            air_flow = br1.air_qty_nm3hr

        pipes1 = calculate_pipe_sizes(PipeInputs(
            ng_flow_nm3hr=ng_flow,
            air_flow_nm3hr=air_flow,
        ))
        # --- Fuel 2 calculation (if dual fuel) ---
        is_dual = req.fuel2_type != "none" and req.fuel2_cv > 0

        # Burner pressure derived from blower pressure: 28" -> 24" w.g., 40" -> 36" w.g.
        burner_pressure_wg = 36 if req.blower_pressure == "40" else 24

        equip1 = select_equipment(
            ng_flow_nm3hr=ng_flow,
            air_flow_nm3hr=air_flow,
            is_dual_fuel=is_dual,
            fuel_cv=f1_cv,
            blower_pressure=req.blower_pressure,
            fuel_type=req.fuel1_type,
            hpu_variant=req.hpu_variant,
            burner_pressure_wg=burner_pressure_wg,
            butterfly_valve_vendor=req.butterfly_valve_vendor,
            shutoff_valve_vendor=req.shutoff_valve_vendor,
            control_mode=req.control_mode,
            auto_control_type=req.auto_control_type,
        )

        f1_is_oil = req.fuel1_type in OIL_FUELS
        f1_oil_lph = equip1["burner"].get("equivalent_lph", 0) if f1_is_oil else 0
        f1_oil_nb = select_oil_pipe_nb(f1_oil_lph) if f1_is_oil else 0

        br2 = None
        pipes2 = None
        equip2 = None
        ng_flow2 = 0
        air_flow2 = 0
        if is_dual:
            if req.mode == "direct":
                # Same heat, different CV → different flow
                ng_flow2 = heat_kcal_hr / req.fuel2_cv
                air_flow2 = air_flow  # air is CV-independent (same heat)
            else:
                br2 = calculate_burner(BurnerInputs(
                    Ti=req.Ti, Tf=req.Tf,
                    refractory_weight=req.refractory_weight,
                    fuel_cv=req.fuel2_cv,
                    time_taken_hr=req.time_taken_hr,
                    refractory_heat_factor=req.refractory_heat_factor,
                    efficiency=req.efficiency,
                ))
                ng_flow2 = br2.extra_firing_rate_nm3hr
                air_flow2 = br2.air_qty_nm3hr
            pipes2 = calculate_pipe_sizes(PipeInputs(
                ng_flow_nm3hr=ng_flow2,
                air_flow_nm3hr=air_flow2,
            ))
            equip2 = select_equipment(
                ng_flow_nm3hr=ng_flow2,
                air_flow_nm3hr=air_flow2,
                is_dual_fuel=is_dual,
                fuel_cv=req.fuel2_cv,
                blower_pressure=req.blower_pressure,
                fuel_type=req.fuel2_type,
                hpu_variant=req.hpu_variant,
                burner_pressure_wg=burner_pressure_wg,
                butterfly_valve_vendor=req.butterfly_valve_vendor,
                shutoff_valve_vendor=req.shutoff_valve_vendor,
                control_mode=req.control_mode,
                auto_control_type=req.auto_control_type,
            )

        # Air is CV-independent, so use fuel1 for air sizing
        if req.control_mode == "manual":
            bom_df = build_vlph_manual_df(
                equipment=equip1,
                ladle_tons=req.ladle_tons,
                fuel1_type=req.fuel1_type,
                pressure_gauge_vendor=req.pressure_gauge_vendor,
                pilot_burner=req.pilot_burner,
                pipeline_weight_kg=req.pipeline_weight_kg,
                include_pilot=req.manual_pilot_burner == "yes",
                pilot_line_fuel=req.pilot_line_fuel,
            )
        else:
            bom_df = build_vlph_120t_df(
                equipment=equip1,
                ladle_tons=req.ladle_tons,
                fuel1_type=req.fuel1_type,
                fuel2_type=req.fuel2_type,
                equipment2=equip2,
                control_mode=req.control_mode,
                auto_control_type=req.auto_control_type,
                control_valve_vendor=req.control_valve_vendor,
                butterfly_valve_vendor=req.butterfly_valve_vendor,
                shutoff_valve_vendor=req.shutoff_valve_vendor,
                pressure_gauge_vendor=req.pressure_gauge_vendor,
                pilot_burner=req.pilot_burner,
                pilot_line_fuel=req.pilot_line_fuel,
                pipeline_weight_kg=req.pipeline_weight_kg,
                purging_line=req.purging_line,
            )

        # Split summary rows from detail rows
        detail = bom_df[bom_df["MEDIA"] != ""].copy()
        bought_out_total = float(bom_df.loc[bom_df["ITEM NAME"] == "BOUGHT OUT ITEMS",   "TOTAL"].values[0]) if "BOUGHT OUT ITEMS" in bom_df["ITEM NAME"].values else 0
        encon_total      = float(bom_df.loc[bom_df["ITEM NAME"] == "ENCON ITEMS",        "TOTAL"].values[0]) if "ENCON ITEMS" in bom_df["ITEM NAME"].values else 0
        grand_total      = float(bom_df.loc[bom_df["ITEM NAME"] == "GRAND TOTAL",        "TOTAL"].values[0]) if "GRAND TOTAL" in bom_df["ITEM NAME"].values else 0

        # Build response — blower HP at user-selected pressure
        cfm = air_flow / 1.7
        blower_hp_calc = cfm * int(req.blower_pressure) / 3200
        resp = {
            "calculations": {
                "mode": req.mode,
                "Ti": req.Ti,
                "Tf": req.Tf,
                "refractory_weight": req.refractory_weight,
                "fuel_cv": f1_cv,
                "fuel1_type": req.fuel1_type,
                "fuel1_name": FUEL_NAMES.get(req.fuel1_type, req.fuel1_type),
                "fuel1_cv": f1_cv,
                "control_mode": req.control_mode,
                "auto_control_type": req.auto_control_type if req.control_mode == "automatic" else None,
                "time_taken_hr": req.time_taken_hr,
                "avg_temp_rise":                  round(br1.avg_temp_rise, 2) if br1 else 0,
                "firing_rate_kcal":               round(br1.firing_rate_kcal, 2) if br1 else 0,
                "heat_load_kcal":                 round(br1.heat_load_kcal, 2) if br1 else 0,
                "fuel_consumption_nm3":           round(br1.fuel_consumption_nm3, 2) if br1 else 0,
                "calculated_firing_rate_nm3hr":   round(br1.calculated_firing_rate_nm3hr, 2) if br1 else round(ng_flow / 1.1, 2),
                "extra_firing_rate_nm3hr":        round(ng_flow, 2),
                "equivalent_lph":                 round(equip1["burner"].get("equivalent_lph", 0), 2),
                "fuel_density":                   equip1["burner"].get("fuel_density", 0),
                "final_firing_rate_mw":           round(br1.final_firing_rate_mw, 2) if br1 else round(ng_flow * f1_cv / (860 * 1000), 2),
                "air_qty_nm3hr":                  round(air_flow, 2),
                "cfm":                            round(cfm, 2),
                "blower_hp_calc":                 round(blower_hp_calc, 2),
            },
            "pipes": {
                "fuel1_label": FUEL_NAMES.get(req.fuel1_type, "Fuel 1"),
                "fuel1_is_oil": f1_is_oil,
                "fuel1_oil_lph": round(f1_oil_lph, 2) if f1_is_oil else None,
                "ng_flow":      round(ng_flow, 2),
                "ng_velocity":  12.7 if not f1_is_oil else 0,
                "ng_dia_mm":    round(pipes1.ng_pipe_inner_dia_mm, 2) if not f1_is_oil else f1_oil_nb,
                "ng_nb":        pipes1.ng_pipe_nb if not f1_is_oil else f1_oil_nb,
                "gas_train_flow": round(equip1["ng_gas_train"]["max_flow"], 0) if not f1_is_oil else None,
                "gas_train_model": f'{equip1["ng_gas_train"]["inlet_nb"]} x {equip1["ng_gas_train"]["outlet_nb"]}' if not f1_is_oil else None,
                "air_flow":     round(air_flow, 2),
                "air_velocity": 15.0,
                "air_dia_mm":   round(pipes1.air_pipe_inner_dia_mm, 2),
                "air_nb":       pipes1.air_pipe_nb,
            },
            "equipment": {
                "burner_model":   equip1["burner"]["model"],
                "blower_model":   equip1["blower"]["model"],
                "blower_hp":      equip1["blower"]["hp"],
                "blower_airflow": equip1["blower"]["airflow_nm3hr"],
                "ng_gas_train":   f'{equip1["ng_gas_train"]["inlet_nb"]} x {equip1["ng_gas_train"]["outlet_nb"]} NB',
                "hpu": (
                    f'{equip1["hpu"]["model"]} — {equip1["hpu"]["unit_kw"]} KW {equip1["hpu"]["variant"]}'
                    if equip1.get("hpu") else None
                ),
            },
            "bom": detail[["MEDIA","ITEM NAME","REFERENCE","QTY","MAKE","UNIT PRICE","TOTAL"]].to_dict(orient="records"),
            "cost_summary": {
                "bought_out_total": round(bought_out_total, 2),
                "encon_total":      round(encon_total, 2),
                "grand_total":      round(grand_total, 2),
            },
        }

        # Add fuel2 data if dual fuel
        if is_dual and pipes2:
            resp["calculations"]["is_dual"] = True
            resp["calculations"]["fuel2_type"] = req.fuel2_type
            resp["calculations"]["fuel2_name"] = FUEL_NAMES.get(req.fuel2_type, req.fuel2_type)
            resp["calculations"]["fuel2_cv"] = req.fuel2_cv
            resp["calculations"]["fuel2_consumption_nm3"] = round(br2.fuel_consumption_nm3, 2) if br2 else 0
            resp["calculations"]["fuel2_firing_rate_nm3hr"] = round(br2.calculated_firing_rate_nm3hr, 2) if br2 else round(ng_flow2 / 1.1, 2)
            resp["calculations"]["fuel2_extra_firing_rate_nm3hr"] = round(ng_flow2, 2)
            resp["calculations"]["fuel2_equivalent_lph"] = round(equip2["burner"].get("equivalent_lph", 0), 2) if equip2 else 0
            resp["calculations"]["fuel2_fuel_density"] = equip2["burner"].get("fuel_density", 0) if equip2 else 0
            resp["pipes"]["fuel2_label"] = FUEL_NAMES.get(req.fuel2_type, "Fuel 2")
            resp["pipes"]["fuel2_flow"] = round(ng_flow2, 2)
            f2_is_oil = req.fuel2_type in OIL_FUELS
            f2_oil_lph = (equip2["burner"].get("equivalent_lph", 0) if (f2_is_oil and equip2) else 0)
            f2_oil_nb = select_oil_pipe_nb(f2_oil_lph) if f2_is_oil else 0
            resp["pipes"]["fuel2_is_oil"] = f2_is_oil
            resp["pipes"]["fuel2_oil_lph"] = round(f2_oil_lph, 2) if f2_is_oil else None
            resp["pipes"]["fuel2_dia_mm"] = round(pipes2.ng_pipe_inner_dia_mm, 2) if not f2_is_oil else f2_oil_nb
            resp["pipes"]["fuel2_nb"] = pipes2.ng_pipe_nb if not f2_is_oil else f2_oil_nb
            if not f2_is_oil and equip2:
                resp["pipes"]["fuel2_gas_train_flow"] = round(equip2["ng_gas_train"]["max_flow"], 0)
                resp["pipes"]["fuel2_gas_train_model"] = f'{equip2["ng_gas_train"]["inlet_nb"]} x {equip2["ng_gas_train"]["outlet_nb"]}'

        return resp
    except Exception as e:
        import traceback
        return {"error": str(e), "trace": traceback.format_exc()}


class RegenCalcRequest(BaseModel):
    material_weight_kg: float
    Ti: float
    Tf: float
    Cp: float = 0.48
    cycle_time_hr: float = 2.0
    efficiency: float = 0.65
    num_pairs_override: int = 0
    markup: float = 1.80


@app.post("/api/regen-calculate")
def regen_calculate(req: RegenCalcRequest):
    try:
        from calculations.regen import RegenInputs, calculate_regen
        from bom.regen_builder import build_regen_df, select_model, get_supplementary_data

        result = calculate_regen(RegenInputs(
            material_weight_kg=req.material_weight_kg,
            Ti=req.Ti,
            Tf=req.Tf,
            Cp=req.Cp,
            cycle_time_hr=req.cycle_time_hr,
            efficiency=req.efficiency,
            num_pairs_override=req.num_pairs_override,
        ))

        # Select the appropriate KW model (smallest >= required_kw per pair)
        kw_per_pair = result.required_kw / max(1, result.num_pairs)
        model_kw    = select_model(kw_per_pair)
        model_markup = req.markup if req.markup != 1.80 else None  # None → use model default

        bom_df = build_regen_df(model_kw, model_markup, num_pairs=result.num_pairs)
        supplementary = get_supplementary_data(model_kw)

        # Augment supplementary with full sizing + nozzle + legacy rates from DB
        try:
            with sqlite3.connect(DB_PATH) as _c:
                # ALL sizing rows (all KW models) — for full Excel-like tables
                sz_cols = [d[0] for d in _c.execute("SELECT * FROM regen_sizing LIMIT 0").description]
                sz_all  = [dict(zip(sz_cols, r)) for r in _c.execute("SELECT * FROM regen_sizing ORDER BY kw").fetchall()]
                supplementary['burner_sizing']['all_sizing'] = sz_all
                # Single selected-KW row (for inline detail)
                sz_row = next((r for r in sz_all if r['kw'] == model_kw), None)
                if sz_row:
                    supplementary['burner_sizing']['dimensions'] = {
                        k: sz_row.get(k) for k in [
                            'shell_thick','retainer_thick','refractory_thick',
                            'dim_L','dim_H','dim_W','bottom_h',
                            'vol_total','vol_effective','vol_refractory',
                            'density_castable','wt_refractory_insulation',
                            'loose_density_balls','vol_available_balls','balls_filling_pct',
                        ]
                    }
                    supplementary['burner_sizing']['weight_detail'] = {
                        k: sz_row.get(k) for k in [
                            'bb_dia_inner','bb_dia_outer','bb_depth','wt_burner_block',
                            'burner_length','burner_dia',
                            'wt_burner_shell','wt_burner_refrac_detail','wt_burner_total',
                            'wt_shell','wt_ss_plate','wt_ceramic_balls_burner',
                            'wt_regen_total','wt_grand_total','bloom_approx_wt',
                        ]
                    }
                # Legacy material rates (from Excel "Burner Sizing and costing" sheet)
                mr_rows = _c.execute("SELECT material, wastage, material_cost, labor_cost FROM regen_material_rates").fetchall()
                legacy_rates = {r[0]: {'material': r[0], 'wastage': r[1], 'mat_cost': r[2], 'labour_cost': r[3] or 0} for r in mr_rows}
                if legacy_rates:
                    supplementary['burner_sizing']['legacy_material_rates'] = list(legacy_rates.values())
                    # Rebuild cost_detail using legacy rates so formula matches Excel exactly
                    from bom.regen_builder import _BURNER_WEIGHTS
                    w = _BURNER_WEIGHTS[model_kw]
                    def _legacy_rate(mat_key):
                        r = legacy_rates.get(mat_key, {})
                        mc, lc, wa = r.get('mat_cost', 0) or 0, r.get('labour_cost', 0) or 0, r.get('wastage', 0) or 0
                        return round((mc + lc) * (1 + wa), 4), mc, lc, wa
                    ms_rate,   ms_m,  ms_l,  ms_w  = _legacy_rate('MS')
                    ss_rate,   ss_m,  ss_l,  ss_w  = _legacy_rate('SS')
                    rf_rate,   rf_m,  rf_l,  rf_w  = _legacy_rate('Refractory')
                    cb_rate,   cb_m,  cb_l,  cb_w  = _legacy_rate('Ceramic Balls')
                    supplementary['burner_sizing']['cost_detail'] = [
                        dict(component='Burner Body',  material='MS',           weight_kg=w['burner_ms'],    mat_cost=ms_m, labour_cost=ms_l, wastage=ms_w, rate=ms_rate, cost=round(w['burner_ms']    * ms_rate, 2)),
                        dict(component='Burner Body',  material='Refractory',   weight_kg=w['burner_refrac'],mat_cost=rf_m, labour_cost=rf_l, wastage=rf_w, rate=rf_rate, cost=round(w['burner_refrac'] * rf_rate, 2)),
                        dict(component='Regenerator',  material='MS',           weight_kg=w['regen_ms'],     mat_cost=ms_m, labour_cost=ms_l, wastage=ms_w, rate=ms_rate, cost=round(w['regen_ms']     * ms_rate, 2)),
                        dict(component='Regenerator',  material='SS',           weight_kg=w['regen_ss'],     mat_cost=ss_m, labour_cost=ss_l, wastage=ss_w, rate=ss_rate, cost=round(w['regen_ss']     * ss_rate, 2)),
                        dict(component='Regenerator',  material='Refractory',   weight_kg=w['regen_refrac'], mat_cost=rf_m, labour_cost=rf_l, wastage=rf_w, rate=rf_rate, cost=round(w['regen_refrac'] * rf_rate, 2)),
                        dict(component='Regenerator',  material='Ceramic Balls',weight_kg=w['regen_ceramic'],mat_cost=cb_m, labour_cost=cb_l, wastage=cb_w, rate=cb_rate, cost=round(w['regen_ceramic'] * cb_rate, 2)),
                        dict(component='Burner Block', material='Refractory',   weight_kg=w['block_refrac'], mat_cost=rf_m, labour_cost=rf_l, wastage=rf_w, rate=rf_rate, cost=round(w['block_refrac'] * rf_rate, 2)),
                    ]
                    supplementary['burner_sizing']['total_unit_cost'] = round(sum(d['cost'] for d in supplementary['burner_sizing']['cost_detail']), 2)
                    supplementary['burner_sizing']['total_pair_cost'] = round(supplementary['burner_sizing']['total_unit_cost'] * 2, 2)
                # Nozzle sizing (all burners)
                nz_cols = [d[0] for d in _c.execute("SELECT * FROM regen_nozzle_sizing LIMIT 0").description]
                nz_rows = _c.execute("SELECT * FROM regen_nozzle_sizing ORDER BY power_kw").fetchall()
                supplementary['nozzle_sizing'] = [dict(zip(nz_cols, r)) for r in nz_rows]
                # All pipe sizes (all KW + all gas types) for full Excel-like table
                ps_cols = [d[0] for d in _c.execute("SELECT * FROM regen_pipe_sizes LIMIT 0").description]
                ps_rows = _c.execute("SELECT * FROM regen_pipe_sizes ORDER BY gas_type, burner_size_kw").fetchall()
                supplementary['all_pipe_sizes'] = [dict(zip(ps_cols, r)) for r in ps_rows]
        except Exception:
            pass  # DB not ready yet — supplementary still has the hardcoded data

        total_cost    = float(bom_df["TOTAL COST"].sum())
        total_selling = float(bom_df["TOTAL SELLING"].sum())

        return {
            "calculations": {
                "material_weight_kg": req.material_weight_kg,
                "Ti": req.Ti,
                "Tf": req.Tf,
                "Cp": req.Cp,
                "delta_T": round(result.delta_T, 2),
                "heat_required_kj": round(result.heat_required_kj, 2),
                "heat_required_kcal": round(result.heat_required_kcal, 2),
                "cycle_time_hr": req.cycle_time_hr,
                "efficiency": req.efficiency,
                "required_kw": round(result.required_kw, 2),
                "num_pairs": result.num_pairs,
                "model_kw": model_kw,
                "total_kw": model_kw * result.num_pairs,
            },
            "bom": bom_df.to_dict(orient="records"),
            "cost_summary": {
                "total_cost": round(total_cost, 2),
                "total_selling": round(total_selling, 2),
                "markup": req.markup,
            },
            "supplementary": supplementary,
        }
    except Exception as e:
        import traceback
        return {"error": str(e), "trace": traceback.format_exc()}


@app.post("/api/hlph-calculate")
def hlph_calculate(req: VLPHCalcRequest):
    try:
        from calculations.burner import BurnerInputs, calculate_burner
        from calculations.pipes import PipeInputs, calculate_pipe_sizes, select_oil_pipe_nb
        from bom.selectors.selection_engine import select_equipment
        from bom.hlph_builder import build_hlph_df, build_hlph_manual_df

        FUEL_NAMES = {
            "ng": "Natural Gas", "lpg": "LPG", "cog": "COG", "bg": "BFG", "rlng": "RLNG", "mg": "Mixed Gas",
            "hsd": "Diesel (HSD)", "ldo": "LDO", "hdo": "HDO", "fo": "Furnace Oil",
            "sko": "Kerosene (SKO)", "cfo": "CFO", "lshs": "LSHS",
        }
        OIL_FUELS = {"hsd", "ldo", "hdo", "fo", "sko", "cfo", "lshs"}

        f1_cv = req.fuel1_cv if req.fuel1_cv > 0 else req.fuel_cv

        burner_inputs = BurnerInputs(
            Ti=req.Ti, Tf=req.Tf,
            refractory_weight=req.refractory_weight,
            fuel_cv=f1_cv,
            time_taken_hr=req.time_taken_hr,
            refractory_heat_factor=req.refractory_heat_factor,
            efficiency=req.efficiency,
        )
        br = calculate_burner(burner_inputs)
        ng_flow = br.extra_firing_rate_nm3hr
        air_flow = br.air_qty_nm3hr

        pipes1 = calculate_pipe_sizes(PipeInputs(ng_flow_nm3hr=ng_flow, air_flow_nm3hr=air_flow))

        burner_pressure_wg = 36 if req.blower_pressure == "40" else 24
        is_dual = req.fuel2_type != "none" and req.fuel2_cv > 0

        equip1 = select_equipment(
            ng_flow_nm3hr=ng_flow, air_flow_nm3hr=air_flow,
            is_dual_fuel=is_dual, fuel_cv=f1_cv,
            blower_pressure=req.blower_pressure, fuel_type=req.fuel1_type,
            hpu_variant=req.hpu_variant, burner_pressure_wg=burner_pressure_wg,
            butterfly_valve_vendor=req.butterfly_valve_vendor,
            shutoff_valve_vendor=req.shutoff_valve_vendor,
            control_mode=req.control_mode, auto_control_type=req.auto_control_type,
        )

        f1_is_oil = req.fuel1_type in OIL_FUELS
        f1_oil_lph = equip1["burner"].get("equivalent_lph", 0) if f1_is_oil else 0

        # Fuel 2
        br2, equip2, ng_flow2 = None, None, 0
        if is_dual:
            br2 = calculate_burner(BurnerInputs(
                Ti=req.Ti, Tf=req.Tf,
                refractory_weight=req.refractory_weight,
                fuel_cv=req.fuel2_cv,
                time_taken_hr=req.time_taken_hr,
                refractory_heat_factor=req.refractory_heat_factor,
                efficiency=req.efficiency,
            ))
            ng_flow2 = br2.extra_firing_rate_nm3hr
            equip2 = select_equipment(
                ng_flow_nm3hr=ng_flow2, air_flow_nm3hr=br2.air_qty_nm3hr,
                is_dual_fuel=is_dual, fuel_cv=req.fuel2_cv,
                blower_pressure=req.blower_pressure, fuel_type=req.fuel2_type,
                hpu_variant=req.hpu_variant, burner_pressure_wg=burner_pressure_wg,
                butterfly_valve_vendor=req.butterfly_valve_vendor,
                shutoff_valve_vendor=req.shutoff_valve_vendor,
                control_mode=req.control_mode, auto_control_type=req.auto_control_type,
            )

        if req.control_mode == "manual":
            bom_df = build_hlph_manual_df(
                equipment=equip1, ladle_tons=req.ladle_tons,
                fuel1_type=req.fuel1_type,
                pressure_gauge_vendor=req.pressure_gauge_vendor,
                pilot_burner=req.pilot_burner,
                pipeline_weight_kg=req.pipeline_weight_kg,
                include_pilot=req.manual_pilot_burner == "yes",
                pilot_line_fuel=req.pilot_line_fuel,
            )
        else:
            bom_df = build_hlph_df(
                equipment=equip1, ladle_tons=req.ladle_tons,
                fuel1_type=req.fuel1_type, fuel2_type=req.fuel2_type,
                equipment2=equip2,
                control_mode=req.control_mode, auto_control_type=req.auto_control_type,
                control_valve_vendor=req.control_valve_vendor,
                butterfly_valve_vendor=req.butterfly_valve_vendor,
                shutoff_valve_vendor=req.shutoff_valve_vendor,
                pressure_gauge_vendor=req.pressure_gauge_vendor,
                pilot_burner=req.pilot_burner, pilot_line_fuel=req.pilot_line_fuel,
                pipeline_weight_kg=req.pipeline_weight_kg,
                purging_line=req.purging_line,
            )

        detail = bom_df[bom_df["MEDIA"] != ""].copy()
        bought_out_total = float(bom_df.loc[bom_df["ITEM NAME"] == "BOUGHT OUT ITEMS", "TOTAL"].values[0]) if "BOUGHT OUT ITEMS" in bom_df["ITEM NAME"].values else 0
        encon_total = float(bom_df.loc[bom_df["ITEM NAME"] == "ENCON ITEMS", "TOTAL"].values[0]) if "ENCON ITEMS" in bom_df["ITEM NAME"].values else 0
        grand_total = float(bom_df.loc[bom_df["ITEM NAME"] == "GRAND TOTAL", "TOTAL"].values[0]) if "GRAND TOTAL" in bom_df["ITEM NAME"].values else 0

        cfm = air_flow / 1.7
        blower_hp_calc = cfm * int(req.blower_pressure) / 3200

        resp = {
            "calculations": {
                "Ti": req.Ti, "Tf": req.Tf,
                "refractory_weight": req.refractory_weight,
                "fuel_cv": f1_cv,
                "fuel1_type": req.fuel1_type,
                "fuel1_name": FUEL_NAMES.get(req.fuel1_type, req.fuel1_type),
                "fuel1_cv": f1_cv,
                "is_dual": is_dual,
                "time_taken_hr": req.time_taken_hr,
                "avg_temp_rise": round(br.avg_temp_rise, 2),
                "firing_rate_kcal": round(br.firing_rate_kcal, 2),
                "heat_load_kcal": round(br.heat_load_kcal, 2),
                "fuel_consumption_nm3": round(br.fuel_consumption_nm3, 2),
                "calculated_firing_rate_nm3hr": round(br.calculated_firing_rate_nm3hr, 2),
                "extra_firing_rate_nm3hr": round(ng_flow, 2),
                "equivalent_lph": round(equip1["burner"].get("equivalent_lph", 0), 2),
                "fuel_density": equip1["burner"].get("fuel_density", 0),
                "final_firing_rate_mw": round(br.final_firing_rate_mw, 2),
                "air_qty_nm3hr": round(air_flow, 2),
                "cfm": round(cfm, 2),
                "blower_hp_calc": round(blower_hp_calc, 2),
            },
            "pipes": {
                "fuel1_label": FUEL_NAMES.get(req.fuel1_type, "Fuel 1"),
                "fuel1_is_oil": f1_is_oil,
                "fuel1_oil_lph": round(f1_oil_lph, 2) if f1_is_oil else None,
                "ng_flow": round(ng_flow, 2),
                "ng_nb": pipes1.ng_pipe_nb,
                "air_flow": round(air_flow, 2),
                "air_nb": pipes1.air_pipe_nb,
                "gas_train_flow": round(equip1["ng_gas_train"]["max_flow"], 0) if equip1.get("ng_gas_train") else 0,
                "gas_train_model": f'{equip1["ng_gas_train"]["inlet_nb"]} x {equip1["ng_gas_train"]["outlet_nb"]}' if equip1.get("ng_gas_train") else "",
            },
            "equipment": {
                "burner_model": equip1["burner"]["model"],
                "blower_model": equip1["blower"]["model"],
                "blower_hp": equip1["blower"]["hp"],
                "blower_airflow": equip1["blower"]["airflow_nm3hr"],
                "ng_gas_train": f'{equip1["ng_gas_train"]["inlet_nb"]} x {equip1["ng_gas_train"]["outlet_nb"]} NB' if equip1.get("ng_gas_train") else "",
                "hpu": f'{equip1["hpu"]["model"]} — {equip1["hpu"]["unit_kw"]} KW' if equip1.get("hpu") else None,
            },
            "bom": detail.to_dict(orient="records"),
            "cost_summary": {
                "bought_out_total": round(bought_out_total, 2),
                "encon_total": round(encon_total, 2),
                "grand_total": round(grand_total, 2),
            },
        }
        return resp
    except Exception as e:
        import traceback
        return {"error": str(e), "trace": traceback.format_exc()}


# ── Box Type Furnace ────────────────────────────────────────────────────────

@app.get("/btf", response_class=HTMLResponse)
def btf_costing_form():
    html_path = os.path.join(BASE_DIR, "btf_costing.html")
    with open(html_path, "r", encoding="utf-8") as f:
        return f.read()


class BTFCalcRequest(BaseModel):
    combustion_mode: str = "onoff"
    markup: float = 1.8


@app.post("/api/btf-calculate")
def btf_calculate(req: BTFCalcRequest):
    try:
        from bom.btf_builder import build_btf_df, get_supplementary
        import json
        df, summary = build_btf_df(combustion_mode=req.combustion_mode, markup=req.markup)
        bom = json.loads(df.to_json(orient="records"))
        return {
            "bom": bom,
            "cost_summary": summary,
            "supplementary": get_supplementary(),
        }
    except Exception as e:
        import traceback
        return {"error": str(e), "detail": traceback.format_exc()}


# ── SNSF BRF (Billet Reheating Furnace) ─────────────────────────────────────

@app.get("/snsf-brf", response_class=HTMLResponse)
def snsf_brf_costing_form():
    html_path = os.path.join(BASE_DIR, "snsf_brf_costing.html")
    with open(html_path, "r", encoding="utf-8") as f:
        return f.read()


class SNSFBRFCalcRequest(BaseModel):
    include_ng_optional: bool = False
    include_client_scope: bool = False


@app.post("/api/snsf-brf-calculate")
def snsf_brf_calculate(req: SNSFBRFCalcRequest):
    try:
        from bom.snsf_brf_builder import build_snsf_brf_df, get_supplementary
        import json
        df, summary = build_snsf_brf_df(
            include_ng_optional=req.include_ng_optional,
            include_client_scope=req.include_client_scope,
        )
        bom = json.loads(df.to_json(orient="records"))
        return {
            "bom": bom,
            "cost_summary": summary,
            "supplementary": get_supplementary(),
        }
    except Exception as e:
        import traceback
        return {"error": str(e), "detail": traceback.format_exc()}


@app.post("/api/generate-quote")
async def generate_quote(req: QuoteRequest):
    try:
        from engine.quote_engine import calculate_quote
        from engine.quote_writer import generate_quote_docx

        seq = next_quote_seq()
        form_data = {
            "quote_seq": seq,
            "customer": {
                "company_name":  req.company_name,
                "address":       ", ".join(filter(None, [req.company_address, req.company_city, req.company_state, req.company_pin])),
                "poc_name":      req.poc_name,
                "poc_designation": req.poc_designation,
                "mobile_no":     req.mobile_no,
                "email":         req.email,
                "project_name":  req.project_name,
                "ref_no":        req.ref_no,
                "gstin":         req.company_gstin,
            },
            "items": [item.dict() for item in req.items],
            "gst_percent": req.gst_percent,
            "freight": req.freight,
            "valid_days": req.valid_days,
        }

        quote_data = calculate_quote(form_data)
        filename = f"Quote_{quote_data['quote_no'].replace('/', '_')}.docx"
        output_path = os.path.join(QUOTES_FOLDER, filename)
        generate_quote_docx(quote_data, output_path)

        return {
            "success": True,
            "quote_no": quote_data["quote_no"],
            "download_url": f"/api/download-quote/{filename}",
            "summary": {
                "subtotal": quote_data["subtotal"],
                "gst": quote_data["gst_amount"],
                "freight": quote_data["freight"],
                "total": quote_data["grand_total"],
            }
        }
    except Exception as e:
        import traceback
        return {"error": str(e), "trace": traceback.format_exc()}


@app.get("/api/price-master/items")
def price_master_items():
    """Return all items grouped by category for the price master page."""
    try:
        conn = sqlite3.connect(DB_PATH)
        rows = conn.execute(
            "SELECT category, item, unit, price, previous_price FROM component_price_master ORDER BY category, item"
        ).fetchall()
        conn.close()
        result = {}
        for cat, item, unit, price, prev in rows:
            result.setdefault(cat, []).append({
                "item": item, "unit": unit or "nos",
                "price": price, "previous_price": prev
            })
        return {"categories": result, "total": len(rows)}
    except Exception as e:
        return {"error": str(e), "categories": {}, "total": 0}


@app.get("/api/price-master/coverage")
def price_master_coverage():
    """Check which BOM items have prices in the DB and which are missing."""
    from bom.vlph_builder import LEGACY_ITEM_SEQUENCE
    from bom.static_items import static_items
    needed = set(LEGACY_ITEM_SEQUENCE) | {item for _, item, _, _ in static_items()}
    needed -= {"RATIO CONTROLLER"}  # excluded from BOM
    try:
        conn = sqlite3.connect(DB_PATH)
        have = {r[0] for r in conn.execute("SELECT item FROM component_price_master").fetchall()}
        conn.close()
    except Exception:
        have = set()
    covered = needed & have
    missing = needed - have
    return {
        "total_in_db":  len(have),
        "total_bom":    len(needed),
        "covered_count": len(covered),
        "missing_count": len(missing),
        "covered":  sorted(covered),
        "missing":  sorted(missing),
        "extra":    sorted(have - needed),
    }


@app.post("/api/upload-pricelist")
async def upload_pricelist(file: UploadFile = File(...)):
    """
    Parses the full ENCON Pricelist WorkBook.
    Updates all master tables: component_price_master, hpu_master,
    burner_pricelist_master, blower_pricelist_master, horizontal_master,
    vertical_master, recuperator_master, gail_gas_burner_master,
    rad_heat_master, rad_heat_tata_master, and burner parts tables.
    """
    try:
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        with open(file_path, "wb") as buf:
            shutil.copyfileobj(file.file, buf)

        conn = sqlite3.connect(DB_PATH)
        results = _parse_pricelist_all(file_path, conn)
        clean_duplicate_rates(conn)
        conn.close()

        updated = {t: r for t, r in results.items() if "rows" in r}
        skipped = {t: r["skipped"] for t, r in results.items() if "skipped" in r}
        errors  = {t: r["error"]   for t, r in results.items() if "error"   in r}

        total_rows = sum(r["rows"] for r in updated.values())

        return {
            "success": True,
            "total_rows_loaded": total_rows,
            "tables_updated": {t: r["rows"] for t, r in updated.items()},
            "tables_skipped": skipped,
            "errors": errors,
        }

    except Exception as e:
        import traceback
        return {"error": str(e), "trace": traceback.format_exc()}


_STOCK_CACHE: dict = {}   # keyed by (file_path, mtime)

@app.get("/api/stock-rates")
def get_stock_rates():
    """
    Parse ENCON Stock Summary Excel and return all sections/items with purchase rates.
    Result is cached in memory; re-parsed only if the file changes.
    """
    import glob as glob_mod
    import openpyxl

    # Find stock file
    candidates = (
        glob_mod.glob(os.path.join(UPLOAD_FOLDER, "Stock*.xlsx")) +
        glob_mod.glob(os.path.join(BASE_DIR, "Stock*.xlsx"))
    )
    if not candidates:
        return {"error": "Stock file not found. Upload a file named Stock*.xlsx"}

    stock_path = candidates[0]
    mtime = os.path.getmtime(stock_path)
    cache_key = (stock_path, mtime)
    if cache_key in _STOCK_CACHE:
        return _STOCK_CACHE[cache_key]

    try:
        wb = openpyxl.load_workbook(stock_path, read_only=True, data_only=True)
        ws = wb["Stock Summary"]
    except Exception as e:
        return {"error": f"Cannot open stock file: {e}"}

    sections = []
    cur_section = "Raw Material"
    cur_items = []

    for r in range(9, 5000):
        v = [ws.cell(r, c).value for c in range(1, 6)]
        if all(x is None for x in v):
            break
        # New section header: col A = None, col B = text, col D ≠ 'TOTAL'
        if v[0] is None and v[1] and str(v[3] or "").strip() != "TOTAL":
            if cur_items:
                sections.append({"section": cur_section, "items": cur_items})
            cur_section = str(v[1]).strip()
            cur_items = []
        elif v[0] is not None and isinstance(v[0], (int, float)):
            name = str(v[1]).strip() if v[1] else ""
            rate = v[3]
            cur_items.append({
                "sl":   int(v[0]),
                "name": name,
                "rate": round(float(rate), 4) if isinstance(rate, (int, float)) else None,
            })

    if cur_items:
        sections.append({"section": cur_section, "items": cur_items})

    total_items = sum(len(s["items"]) for s in sections)
    result = {
        "file": os.path.basename(stock_path),
        "sections": sections,
        "total_items": total_items,
    }
    _STOCK_CACHE[cache_key] = result
    return result


# ── Stock → component_price_master mapping ────────────────────────────────────
# Maps pricelist item name (or prefix) → stock item name (exact)
_STOCK_PRICE_MAP = {
    # MS structural
    "M.S. Angle 65,50":       "M.S Angle 50x50x6 mm",
    "M.S.Chanel":             "M.S Channel 100x50",
    "M.S. Chanel":            "M.S Channel 100x50",
    "M.S. Angle 100,100":     "M.S Angle 100X100X10",
    "M.S. Angle 50*6":        "M.S Angle 50x50x6 mm",
    "M.S. Flat":              None,   # no good match
    "M.S. Round":             "M.S Round Dia 16 MM",
    # MS plates / sheets
    "M.S. Plate 5mm":         "M.S PLATE 1500X6300X5 MM",
    "M.S. Plate 8mm":         "M.S Plate 1500X6300X10 MM",
    "M.S. Plate 16mm* 5mm":   "M.S PLATE 1500X6300X5 MM",
    "M.S. Plate 16mm*5mm":    "M.S PLATE 1500X6300X5 MM",
    "M.S. Plate 16mm*10mm":   "M.S Plate 1500X6300X10 MM",
    "M.S. Sheet 2mm":         None,
    "M.S. Sheet  2mm":        None,
    "M.S. Sheet 3mm":         None,
    "M.S. Sheet  3mm":        None,
    "M.S. Sheet 4mm":         None,
    "M.S. Sheet 5mm":         "M.S Chequered Plate 1250X5000X5 MM",
    "M.S. Sheet 8mm":         "M.S Plate 1510x6310x16mm",
    # MS tubes
    'M.S. Tube "B" Class 1.5 in': "M.S ERW Pipe 40 NB",
    'M.S. Tube "C" Class 1.5 in': "M.S ERW Pipe 40 NB",
    "M.S. Tube B Class 1.5 in":   "M.S ERW Pipe 40 NB",
    "M.S. Tube C Class 1.5 in":   "M.S ERW Pipe 40 NB",
    # SS
    "S.S. Sheet 3mm":         "S.S Plate 1500X3000X3 MM",
    "SS Pipe 304 60 X 3mm":   "S.S 304 ERW PIPE 100 MM OD",
    "SS Pipe 304 60x3mm (per mtr)": "S.S 304 ERW PIPE 100 MM OD",
    "SS Pipe 304 76 X 3mm":   "S.S ERW PIPE OD 65 X ID 58 / 57 MM",
    "SS Pipe 304 76x3mm (per mtr)": "S.S ERW PIPE OD 65 X ID 58 / 57 MM",
    "SS Pipe 304 100 X 3mm":  "S.S 304 ERW Pipe 100 NB",
    "SS Pipe 304 100x3mm (per mtr)": "S.S 304 ERW Pipe 100 NB",
    # Refractories
    "Ceramic Fiber":          "Ceramic Fiber 128 Kg/m3",
    "Whyteheat K":            "Whytheat-A",
}

@app.post("/api/stock/sync")
def sync_stock_to_pricelist():
    """
    Reads the stock file, matches items to component_price_master via
    _STOCK_PRICE_MAP, and updates matched rows.  Returns a summary.
    """
    import glob as glob_mod
    import openpyxl

    candidates = (
        glob_mod.glob(os.path.join(UPLOAD_FOLDER, "Stock*.xlsx")) +
        glob_mod.glob(os.path.join(BASE_DIR, "Stock*.xlsx"))
    )
    if not candidates:
        return {"error": "Stock file not found"}
    stock_path = candidates[0]

    try:
        wb = openpyxl.load_workbook(stock_path, read_only=True, data_only=True)
        ws = wb["Stock Summary"]
    except Exception as e:
        return {"error": str(e)}

    # Build {name_lower: rate} lookup from stock
    stock_lookup = {}
    for r in range(9, 5000):
        sl   = ws.cell(r, 1).value
        name = ws.cell(r, 2).value
        rate = ws.cell(r, 4).value
        if sl is None and name is None:
            break
        if isinstance(sl, (int, float)) and name and isinstance(rate, (int, float)):
            stock_lookup[str(name).strip().lower()] = (str(name).strip(), float(rate))
    wb.close()

    conn = sqlite3.connect(DB_PATH)
    updated, skipped = [], []

    # Fetch all pricelist items
    pl_items = conn.execute("SELECT rowid, item, price FROM component_price_master").fetchall()
    for rowid, item, old_price in pl_items:
        stock_name = _STOCK_PRICE_MAP.get(item)
        if stock_name is None:
            if item in _STOCK_PRICE_MAP:
                skipped.append({"item": item, "reason": "manual_skip"})
            continue
        hit = stock_lookup.get(stock_name.lower())
        if not hit:
            skipped.append({"item": item, "stock_name": stock_name, "reason": "not_in_stock"})
            continue
        new_price = round(hit[1], 2)
        if new_price != old_price:
            conn.execute("UPDATE component_price_master SET price=? WHERE rowid=?", (new_price, rowid))
            updated.append({"item": item, "stock_name": hit[0], "old": old_price, "new": new_price})

    conn.commit()

    # ── Cascade: re-run all parsers so every computed amount/selling price
    #    reflects the updated rates (recuperator, HPU, blower, LPH, burner…)
    cascade_ok = False
    cascade_tables = {}
    xl_path = _find_latest_pricebook()
    if xl_path and updated:
        try:
            cascade_tables = _cascade_recalculate(xl_path, conn)
            conn.commit()
            cascade_ok = True
        except Exception as ce:
            cascade_tables = {"error": str(ce)}

    conn.close()
    _STOCK_CACHE.clear()
    return {
        "file": os.path.basename(stock_path),
        "updated": updated,
        "skipped_count": len(skipped),
        "updated_count": len(updated),
        "cascade": cascade_ok,
        "cascade_tables": {k: v for k, v in cascade_tables.items() if isinstance(v, dict)},
    }


@app.post("/api/stock/upload")
async def upload_stock_file(file: UploadFile = File(...)):
    """Accept a new Stock*.xlsx upload and save to BASE_DIR, then auto-sync."""
    import shutil
    # Validate name
    if not file.filename.startswith("Stock") or not file.filename.endswith(".xlsx"):
        return {"error": "File must be named Stock*.xlsx"}
    dest = os.path.join(BASE_DIR, file.filename)
    with open(dest, "wb") as f:
        shutil.copyfileobj(file.file, f)
    _STOCK_CACHE.clear()
    # Auto-sync after upload
    sync_result = sync_stock_to_pricelist()
    return {"uploaded": file.filename, "sync": sync_result}


class ExcelExportRequest(BaseModel):
    equipment_type: str          # "VLPH" | "HLPH" | "Regen"
    customer: dict = {}
    calculations: dict = {}
    bom: list = []
    cost_summary: dict = {}
    pipes: dict = {}
    equipment: dict = {}


@app.post("/api/export-excel")
def export_excel(req: ExcelExportRequest):
    import io
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from fastapi.responses import StreamingResponse

    NAVY   = "1A3A5C"
    LIGHT  = "EFF6FF"
    WHITE  = "FFFFFF"
    GREY   = "F8FAFC"
    GREEN  = "065F46"
    GREEN_BG = "F0FDF4"

    def thin():
        s = Side(style="thin", color="E2E8F0")
        return Border(left=s, right=s, top=s, bottom=s)

    wb = Workbook()

    def hdr(ws, row, col, val, bg=NAVY, fg=WHITE, bold=True, size=11, span=None):
        c = ws.cell(row=row, column=col, value=val)
        c.font = Font(bold=bold, color=fg, size=size, name="Calibri")
        c.fill = PatternFill("solid", fgColor=bg)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = thin()
        return c

    def cell(ws, row, col, val, bold=False, align="left", bg=WHITE, fg="1E293B", num_fmt=None):
        c = ws.cell(row=row, column=col, value=val)
        c.font = Font(bold=bold, color=fg, size=10, name="Calibri")
        c.fill = PatternFill("solid", fgColor=bg)
        c.alignment = Alignment(horizontal=align, vertical="center")
        c.border = thin()
        if num_fmt:
            c.number_format = num_fmt
        return c

    def section_hdr(ws, r, ncols, text, bg=NAVY, fg=WHITE, size=10):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
        c = ws.cell(row=r, column=1, value=text)
        c.font = Font(bold=True, color=fg, size=size, name="Calibri")
        c.fill = PatternFill("solid", fgColor=bg)
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[r].height = 20
        return c

    def title_row(ws, r, ncols, text, size=14):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
        c = ws.cell(row=r, column=1, value=text)
        c.font = Font(bold=True, color=WHITE, size=size, name="Calibri")
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[r].height = 30
        return c

    # ════════════════════════════════════════════════════════════════════════
    #  REGEN — 5 sheets matching HTML tabs exactly
    # ════════════════════════════════════════════════════════════════════════
    if req.equipment_type == "Regen":
        supp = (req.equipment or {}).get("supplementary", {}) if req.equipment else {}
        bs   = supp.get("burner_sizing", {}) if supp else {}
        calc = req.calculations
        cs   = req.cost_summary

        # ── Sheet 1: Process Calcs ────────────────────────────────────────
        ws1 = wb.active
        ws1.title = "Process Calcs"
        for col, w in zip("ABCDE", [28, 22, 22, 14, 18]):
            ws1.column_dimensions[col].width = w

        r1 = 1
        title_row(ws1, r1, 5, "ENCON — Regenerative Burner System — Process Calcs")
        r1 += 1
        ws1.merge_cells(f"A{r1}:E{r1}")
        d = ws1.cell(row=r1, column=1, value=f"Generated: {datetime.now().strftime('%d %b %Y, %H:%M')}")
        d.font = Font(color="64748B", size=9, name="Calibri")
        d.fill = PatternFill("solid", fgColor=LIGHT)
        d.alignment = Alignment(horizontal="right", vertical="center")
        r1 += 2

        # Customer
        if req.customer:
            section_hdr(ws1, r1, 5, "CUSTOMER DETAILS")
            r1 += 1
            for label, val in [
                ("Company",     req.customer.get("company_name","")),
                ("Contact",     req.customer.get("poc_name","")),
                ("Designation", req.customer.get("poc_designation","")),
                ("Mobile",      req.customer.get("mobile_no","")),
                ("Email",       req.customer.get("email","")),
                ("Project",     req.customer.get("project_name","")),
                ("Ref No.",     req.customer.get("ref_no","")),
            ]:
                if val:
                    cell(ws1, r1, 1, label, bold=True, bg=GREY)
                    ws1.merge_cells(f"B{r1}:E{r1}")
                    cell(ws1, r1, 2, val)
                    r1 += 1
            r1 += 1

        # Process Parameters
        section_hdr(ws1, r1, 5, "PROCESS PARAMETERS")
        r1 += 1
        proc_params = [
            ("Material Weight",   calc.get("material_weight_kg",""), "kg"),
            ("Initial Temp (Ti)", calc.get("Ti",""),                 "°C"),
            ("Final Temp (Tf)",   calc.get("Tf",""),                 "°C"),
            ("Temp Rise (ΔT)",    calc.get("delta_T",""),            "°C"),
            ("Specific Heat Cp",  calc.get("Cp",""),                 "kJ/kg·°C"),
            ("Cycle Time",        calc.get("cycle_time_hr",""),      "hr"),
            ("Efficiency",        f"{round(calc.get('efficiency',0)*100)}%", ""),
            ("Heat Required",     calc.get("heat_required_kj",""),   "kJ"),
            ("Required Power",    calc.get("required_kw",""),        "kW"),
            ("No. of Pairs",      calc.get("num_pairs",""),          f"× {calc.get('model_kw',1000)} KW"),
            ("Total KW",          calc.get("total_kw",""),           "KW"),
        ]
        for i, (label, val, unit) in enumerate(proc_params):
            bg = GREY if i % 2 == 0 else WHITE
            cell(ws1, r1, 1, label, bold=True, bg=bg)
            cell(ws1, r1, 2, val, bg=bg, align="right")
            cell(ws1, r1, 3, unit, bg=bg)
            ws1.merge_cells(f"D{r1}:E{r1}")
            cell(ws1, r1, 4, "", bg=bg)
            r1 += 1
        r1 += 1

        # Cost Summary
        section_hdr(ws1, r1, 5, "COST SUMMARY")
        r1 += 1
        for label, val, is_total in [
            ("Total Cost Price",  cs.get("total_cost", 0),    False),
            ("Markup",            cs.get("markup", 0),         False),
            ("Total Selling Price", cs.get("total_selling", 0), True),
        ]:
            bg  = GREEN_BG if is_total else GREY
            fg_ = GREEN    if is_total else "1E293B"
            cell(ws1, r1, 1, label, bold=is_total, bg=bg, fg=fg_)
            ws1.merge_cells(f"B{r1}:D{r1}")
            cell(ws1, r1, 2, "", bg=bg)
            c = ws1.cell(row=r1, column=5, value=val)
            c.font  = Font(bold=is_total, color=fg_, size=11 if is_total else 10, name="Calibri")
            c.fill  = PatternFill("solid", fgColor=bg)
            c.alignment = Alignment(horizontal="right", vertical="center")
            c.number_format = '#,##0.00'
            c.border = thin()
            ws1.row_dimensions[r1].height = 22 if is_total else 18
            r1 += 1

        # ── Sheet 2: Burner Sizing and Costing ───────────────────────────
        ws2 = wb.create_sheet("Burner Sizing and Costing")
        ALL_COLS_BS = list("ABCDEFGHIJKLMNOP")
        for i, w in enumerate([14,14,14,14,12,12,12,12,14,14,14,14,14,14,14,14]):
            ws2.column_dimensions[ALL_COLS_BS[i]].width = w

        r2 = 1
        title_row(ws2, r2, 16, f"Burner + Regenerator Sizing — All KW Rows (Selected: {bs.get('kw','')} KW)")
        r2 += 2

        all_sz  = bs.get("all_sizing", [])
        sel_kw  = bs.get("kw")
        nozzles = supp.get("nozzle_sizing", [])

        # — Regenerator Dimensions table —
        section_hdr(ws2, r2, 16, "REGENERATOR DIMENSIONS (all KW models)")
        r2 += 1
        dim_hdrs = ["KW","Shell (mm)","Retainer (mm)","Refrac. (mm)","L (m)","H (m)","W (m)","Bottom H (m)",
                    "Vol Total (m³)","Vol Eff. (m³)","Vol Refrac. (m³)","Density Castable",
                    "Wt Refrac. (kg)","Loose Density","Vol Balls (m³)","Balls Fill %"]
        for ci, lbl in enumerate(dim_hdrs, 1):
            hdr(ws2, r2, ci, lbl, size=8)
        r2 += 1
        for row_s in all_sz:
            is_sel = row_s.get("kw") == sel_kw
            bg = LIGHT if is_sel else (GREY if r2 % 2 == 0 else WHITE)
            vals = [
                row_s.get("kw",""), row_s.get("shell_thick",""), row_s.get("retainer_thick",""),
                row_s.get("refractory_thick",""), row_s.get("dim_L",""), row_s.get("dim_H",""),
                row_s.get("dim_W",""), row_s.get("bottom_h",""),
                round(row_s["vol_total"],4) if row_s.get("vol_total") is not None else "",
                round(row_s["vol_effective"],4) if row_s.get("vol_effective") is not None else "",
                round(row_s["vol_refractory"],4) if row_s.get("vol_refractory") is not None else "",
                row_s.get("density_castable",""),
                round(row_s["wt_refractory_insulation"],1) if row_s.get("wt_refractory_insulation") is not None else "",
                row_s.get("loose_density_balls",""),
                round(row_s["vol_available_balls"],4) if row_s.get("vol_available_balls") is not None else "",
                f"{row_s['balls_filling_pct']}%" if row_s.get("balls_filling_pct") is not None else "",
            ]
            for ci, v in enumerate(vals, 1):
                cell(ws2, r2, ci, v, bold=is_sel, bg=bg, align="center")
            r2 += 1
        r2 += 1

        # — Weight Breakdown table —
        section_hdr(ws2, r2, 12, "BURNER SIZE — WEIGHT BREAKDOWN (all KW models)")
        r2 += 1
        wt_cols = list("ABCDEFGHIJKL")
        for i, w in enumerate([14,14,14,14,14,14,14,14,14,12,12,12]):
            ws2.column_dimensions[wt_cols[i]].width = w
        wt_hdrs = ["KW","Burner MS (kg)","Burner Refrac. (kg)","Regen MS (kg)","Regen SS (kg)",
                   "Regen Refrac. (kg)","Ceramic Balls (kg)","BB Refrac. (kg)","Grand Total (kg)",
                   "BB Dia Inner (m)","Burner Length (m)","Burner Dia (m)"]
        for ci, lbl in enumerate(wt_hdrs, 1):
            hdr(ws2, r2, ci, lbl, size=8)
        r2 += 1
        for row_s in all_sz:
            is_sel = row_s.get("kw") == sel_kw
            bg = LIGHT if is_sel else (GREY if r2 % 2 == 0 else WHITE)
            def fmt_wt(v): return round(v, 1) if v is not None else ""
            vals = [
                row_s.get("kw",""), fmt_wt(row_s.get("wt_burner_ms")), fmt_wt(row_s.get("wt_burner_refrac")),
                fmt_wt(row_s.get("wt_regen_ms")), fmt_wt(row_s.get("wt_regen_ss")),
                fmt_wt(row_s.get("wt_regen_refrac")), fmt_wt(row_s.get("wt_ceramic_balls")),
                fmt_wt(row_s.get("wt_burner_block_summary")), fmt_wt(row_s.get("wt_total")),
                row_s.get("bb_dia_inner",""), row_s.get("burner_length",""), row_s.get("burner_dia",""),
            ]
            for ci, v in enumerate(vals, 1):
                cell(ws2, r2, ci, v, bold=is_sel, bg=bg, align="center")
            r2 += 1
        r2 += 1

        # — Nozzle Sizing —
        if nozzles:
            section_hdr(ws2, r2, 8, "NOZZLE SIZING")
            r2 += 1
            for ci, lbl in enumerate(["Burner Name","Power (KW)","DN Air In (mm)","Air Speed (m/s)",
                                       "DN Fume Out (mm)","Fume Speed (m/s)","DN NG In (mm)","NG Speed (m/s)"], 1):
                hdr(ws2, r2, ci, lbl, size=9)
            r2 += 1
            for nz in nozzles:
                is_sel = nz.get("power_kw") == sel_kw
                bg = LIGHT if is_sel else (GREY if r2 % 2 == 0 else WHITE)
                for ci, v in enumerate([nz.get("burner_name",""), nz.get("power_kw",""),
                                        nz.get("dn_air_in",""), nz.get("air_speed_ms",""),
                                        nz.get("dn_fume_out",""), nz.get("fume_speed_ms",""),
                                        nz.get("dn_ng_in",""), nz.get("ng_speed_ms","")], 1):
                    cell(ws2, r2, ci, v, bold=is_sel, bg=bg, align="center")
                r2 += 1
            r2 += 1

        # — Costing Consideration: Material Rates —
        legacy_rates = bs.get("legacy_material_rates", [])
        section_hdr(ws2, r2, 5, "COSTING CONSIDERATION — MATERIAL RATES")
        r2 += 1
        for ci, lbl in enumerate(["Material","Wastage","Material Cost (₹/kg)","Labour Cost (₹/kg)","Effective Rate (₹/kg)"], 1):
            hdr(ws2, r2, ci, lbl, size=9)
        r2 += 1
        default_rates = [("MS",0.1,50,25),("SS",0.1,50,25),("Refractory",0.1,56,25),("Ceramic Balls",0.1,125,0)]
        rate_rows = legacy_rates if legacy_rates else [
            {"material":m,"wastage":wa,"mat_cost":mc,"labour_cost":lc} for m,wa,mc,lc in default_rates
        ]
        for i, mr in enumerate(rate_rows):
            bg = GREY if i % 2 == 0 else WHITE
            wa  = mr.get("wastage") or 0
            mc  = mr.get("mat_cost") or 0
            lc  = mr.get("labour_cost") or 0
            eff = round((mc + lc) * (1 + wa), 2)
            for ci, v in enumerate([mr.get("material",""), f"{round(wa*100)}%", mc or "—", lc or "—", eff], 1):
                cell(ws2, r2, ci, v, bg=bg, align="right" if ci > 1 else "left")
            r2 += 1
        r2 += 1

        # — Cost Breakdown all KW —
        section_hdr(ws2, r2, 9, "COST BREAKDOWN — PER UNIT (1 burner + regenerator), ALL KW MODELS")
        r2 += 1
        cd_hdrs = ["KW","Burner MS (₹)","Burner Refrac. (₹)","Regen MS (₹)","Regen SS (₹)",
                   "Regen Refrac. (₹)","Ceramic Balls (₹)","Burner Block (₹)","TOTAL UNIT COST (₹)"]
        for ci, lbl in enumerate(cd_hdrs, 1):
            hdr(ws2, r2, ci, lbl, size=8)
        r2 += 1
        for row_s in all_sz:
            is_sel = row_s.get("kw") == sel_kw
            bg = LIGHT if is_sel else (GREY if r2 % 2 == 0 else WHITE)
            vals = [
                row_s.get("kw",""),
                round(row_s["cost_burner_ms"],2) if row_s.get("cost_burner_ms") is not None else "",
                round(row_s["cost_burner_refrac"],2) if row_s.get("cost_burner_refrac") is not None else "",
                round(row_s["cost_regen_ms"],2) if row_s.get("cost_regen_ms") is not None else "",
                round(row_s["cost_regen_ss"],2) if row_s.get("cost_regen_ss") is not None else "",
                round(row_s["cost_regen_refrac"],2) if row_s.get("cost_regen_refrac") is not None else "",
                round(row_s["cost_ceramic_balls"],2) if row_s.get("cost_ceramic_balls") is not None else "",
                round(row_s["cost_burner_block"],2) if row_s.get("cost_burner_block") is not None else "",
                round(row_s["cost_total"],2) if row_s.get("cost_total") is not None else "",
            ]
            for ci, v in enumerate(vals, 1):
                cell(ws2, r2, ci, v, bold=is_sel, bg=bg, align="center",
                     num_fmt='#,##0.00' if ci > 1 and isinstance(v,(int,float)) else None)
            r2 += 1
        # Totals row
        r2_tot = r2
        ws2.merge_cells(f"A{r2_tot}:H{r2_tot}")
        cell(ws2, r2_tot, 1, f"Selected: {sel_kw} KW — 1 Pair (×2)", bold=True, bg=GREEN_BG, fg=GREEN, align="right")
        cell(ws2, r2_tot, 9, bs.get("total_pair_cost", 0), bold=True, bg=GREEN_BG, fg=GREEN, align="center",
             num_fmt='#,##0.00')

        # ── Sheet 3: Burner Pipe Size ─────────────────────────────────────
        ws3 = wb.create_sheet("Burner Pipe Size")
        PIPE_COLS = list("ABCDEFGHIJKLM")
        for i, w in enumerate([16,14,14,14,12,12,12,12,12,12,12,12,12]):
            ws3.column_dimensions[PIPE_COLS[i]].width = w

        r3 = 1
        title_row(ws3, r3, 13, "Burner Pipe Size — All Gas Types and KW Models")
        r3 += 2

        all_ps  = supp.get("all_pipe_sizes", [])
        ps_sel  = supp.get("pipe_sizes", {})
        sel_ps_kw = bs.get("kw")
        pipe_col_hdrs = ["Burner Size (KW)","Gas Flow (Nm³/hr)","Air Flow (Nm³/hr)","Total Flue (Nm³/hr)",
                         "DN Air (mm)","DN Gas (mm)","DN Flue (mm)",
                         "Area Air (m²)","Area Gas (m²)","Area Flue (m²)",
                         "Vel Gas (m/s)","Vel Air (m/s)","Vel Flue (m/s)"]

        if all_ps:
            # Group by gas type
            gas_types = []
            for row_p in all_ps:
                gt = row_p.get("gas_type","")
                if gt not in gas_types:
                    gas_types.append(gt)
            for gt in gas_types:
                section_hdr(ws3, r3, 13, gt)
                r3 += 1
                for ci, lbl in enumerate(pipe_col_hdrs, 1):
                    hdr(ws3, r3, ci, lbl, size=8)
                r3 += 1
                for row_p in all_ps:
                    if row_p.get("gas_type") != gt:
                        continue
                    is_sel = row_p.get("burner_size_kw") == sel_ps_kw
                    bg = LIGHT if is_sel else (GREY if r3 % 2 == 0 else WHITE)
                    vals = [
                        row_p.get("burner_size_kw",""),
                        round(row_p["gas_flow_nm3hr"],2) if row_p.get("gas_flow_nm3hr") is not None else "",
                        round(row_p["air_flow_nm3hr"],2) if row_p.get("air_flow_nm3hr") is not None else "",
                        round(row_p["flue_flow_nm3hr"],2) if row_p.get("flue_flow_nm3hr") is not None else "",
                        row_p.get("dn_air_mm",""), row_p.get("dn_gas_mm",""), row_p.get("dn_flue_mm",""),
                        round(row_p["area_air_m2"],4) if row_p.get("area_air_m2") is not None else "",
                        round(row_p["area_gas_m2"],4) if row_p.get("area_gas_m2") is not None else "",
                        round(row_p["area_flue_m2"],4) if row_p.get("area_flue_m2") is not None else "",
                        round(row_p["vel_gas_ms"],1) if row_p.get("vel_gas_ms") is not None else "",
                        round(row_p["vel_air_ms"],1) if row_p.get("vel_air_ms") is not None else "",
                        round(row_p["vel_flue_ms"],1) if row_p.get("vel_flue_ms") is not None else "",
                    ]
                    for ci, v in enumerate(vals, 1):
                        cell(ws3, r3, ci, v, bold=is_sel, bg=bg, align="center")
                    r3 += 1
                r3 += 1
        elif ps_sel:
            # Fallback: single KW from pipe_sizes
            section_hdr(ws3, r3, 7, f"Natural Gas (NG) — {ps_sel.get('fuel','')} @ {ps_sel.get('pressure','')}")
            r3 += 1
            for ci, lbl in enumerate(pipe_col_hdrs[:7], 1):
                hdr(ws3, r3, ci, lbl, size=9)
            r3 += 1
            fb_vals = [ps_sel.get("kw",""), ps_sel.get("ng_flow_nm3hr",""), ps_sel.get("air_flow_nm3hr",""),
                       ps_sel.get("flue_flow_nm3hr",""), ps_sel.get("air_line_dn",""),
                       ps_sel.get("gas_line_dn",""), ps_sel.get("flue_line_dn","")]
            for ci, v in enumerate(fb_vals, 1):
                cell(ws3, r3, ci, v, bg=LIGHT, align="center")

        # ── Sheet 4: Blower ───────────────────────────────────────────────
        ws4 = wb.create_sheet("Blower")
        for col, w in zip("ABCDEF", [22, 14, 14, 16, 20, 20]):
            ws4.column_dimensions[col].width = w

        r4 = 1
        bl  = supp.get("blower_selection", {})
        cat = supp.get("blower_catalogue", [])
        title_row(ws4, r4, 6, f"ENCON 40\" WG Blower Selection — {bl.get('kw','')} KW")
        r4 += 2

        section_hdr(ws4, r4, 6, "SELECTED BLOWER")
        r4 += 1
        sel_blower_rows = [
            ("Burner KW",          f"{bl.get('kw','')} KW"),
            ("Selected Model",     bl.get("selected_model","")),
            ("HP",                 bl.get("hp","")),
            ("CFM",                bl.get("cfm","")),
            ("Flow Rate (Nm³/hr)", bl.get("nm3hr","")),
            ("Qty per pair",       bl.get("qty_per_pair","")),
            ("Price w/o Motor ₹",  bl.get("price_without_motor",0)),
            ("Price with Motor ₹", bl.get("price_with_motor",0)),
            ("Costing Price Used ₹", bl.get("costing_price",0)),
        ]
        for i, (lbl, val) in enumerate(sel_blower_rows):
            is_cost = "Costing" in lbl
            bg = GREEN_BG if is_cost else (GREY if i % 2 == 0 else WHITE)
            fg_ = GREEN if is_cost else "1E293B"
            cell(ws4, r4, 1, lbl, bold=is_cost, bg=bg, fg=fg_)
            ws4.merge_cells(f"B{r4}:F{r4}")
            cell(ws4, r4, 2, val, bold=is_cost, bg=bg, fg=fg_, align="right",
                 num_fmt='#,##0.00' if isinstance(val,(int,float)) else None)
            r4 += 1
        r4 += 1

        section_hdr(ws4, r4, 6, "ENCON 40\" WG BLOWER CATALOGUE")
        r4 += 1
        for ci, lbl in enumerate(["Model","HP","CFM","Nm³/hr","Price w/o Motor ₹","Price with Motor ₹"], 1):
            hdr(ws4, r4, ci, lbl, size=9)
        r4 += 1
        for row_b in cat:
            is_sel_b = row_b.get("model","") == bl.get("selected_model","")
            bg = LIGHT if is_sel_b else (GREY if r4 % 2 == 0 else WHITE)
            for ci, v in enumerate([row_b.get("model",""), row_b.get("hp",""), row_b.get("cfm",""),
                                     row_b.get("nm3hr",""), row_b.get("price_without_motor",""),
                                     row_b.get("price_with_motor","")], 1):
                cell(ws4, r4, ci, v, bold=is_sel_b, bg=bg,
                     align="right" if ci > 2 else "left",
                     num_fmt='#,##0.00' if ci > 4 and isinstance(v,(int,float)) else None)
            r4 += 1

        # ── Sheet 5: BOM ──────────────────────────────────────────────────
        ws5 = wb.create_sheet("BOM")
        for col, w in zip("ABCDEFGHI", [6, 16, 28, 26, 8, 14, 14, 14, 14]):
            ws5.column_dimensions[col].width = w

        r5 = 1
        title_row(ws5, r5, 9, f"BILL OF MATERIALS — REGENERATIVE BURNER SYSTEM ({calc.get('model_kw','1000')} KW)")
        r5 += 2

        bom_col_hdrs = ["S.No.","SECTION","ITEM NAME","SPECIFICATION","QTY",
                        "COST/UNIT ₹","TOTAL COST ₹","SELL/UNIT ₹","TOTAL SELLING ₹"]
        for ci, lbl in enumerate(bom_col_hdrs, 1):
            hdr(ws5, r5, ci, lbl, size=9)
        ws5.row_dimensions[r5].height = 22
        r5 += 1
        sno = 0
        for i, row_d in enumerate(req.bom):
            bg = GREY if i % 2 == 0 else WHITE
            sno += 1
            bom_vals = [sno,
                        row_d.get("SECTION",""), row_d.get("ITEM NAME",""), row_d.get("SPECIFICATION",""),
                        row_d.get("QTY",""),
                        row_d.get("COST/UNIT",0), row_d.get("TOTAL COST",0),
                        row_d.get("SELL/UNIT",0), row_d.get("TOTAL SELLING",0)]
            for ci, v in enumerate(bom_vals, 1):
                cell(ws5, r5, ci, v, bg=bg, align="right" if ci >= 5 else ("center" if ci==1 else "left"),
                     num_fmt='#,##0.00' if ci >= 6 and isinstance(v,(int,float)) else None)
            ws5.row_dimensions[r5].height = 18
            r5 += 1

        # Grand total row
        ws5.merge_cells(f"A{r5}:F{r5}")
        cell(ws5, r5, 1, "GRAND TOTAL", bold=True, bg=GREEN_BG, fg=GREEN, align="right")
        cell(ws5, r5, 7, cs.get("total_cost",0), bold=True, bg=GREEN_BG, fg=GREEN,
             align="right", num_fmt='#,##0.00')
        cell(ws5, r5, 8, "", bold=True, bg=GREEN_BG, fg=GREEN)
        cell(ws5, r5, 9, cs.get("total_selling",0), bold=True, bg=GREEN_BG, fg=GREEN,
             align="right", num_fmt='#,##0.00')
        ws5.row_dimensions[r5].height = 22

        # Save and return
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        cname = req.customer.get("company_name","").replace(" ","_") if req.customer else "ENCON"
        fname = f"Regen_Costing_{cname or 'ENCON'}_{datetime.now().strftime('%d%b%Y')}.xlsx"
        return StreamingResponse(buf,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{fname}"'})

    # ════════════════════════════════════════════════════════════════════════
    #  BTF / SNSF BRF path
    # ════════════════════════════════════════════════════════════════════════
    if req.equipment_type in ("BTF", "SNSF BRF"):
        ws = wb.active
        ws.title = req.equipment_type
        ws.column_dimensions["A"].width = 36
        ws.column_dimensions["B"].width = 42
        ws.column_dimensions["C"].width = 14
        ws.column_dimensions["D"].width = 12
        ws.column_dimensions["E"].width = 18
        ws.column_dimensions["F"].width = 20
        ws.column_dimensions["G"].width = 12
        ws.column_dimensions["H"].width = 20

        r1 = 1
        ws.merge_cells(f"A{r1}:H{r1}")
        t = ws.cell(row=r1, column=1, value=f"ENCON — {req.equipment_type} Costing Sheet")
        t.font = Font(bold=True, color=WHITE, size=14, name="Calibri")
        t.fill = PatternFill("solid", fgColor=NAVY)
        t.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[r1].height = 30
        r1 += 1

        ws.merge_cells(f"A{r1}:H{r1}")
        d = ws.cell(row=r1, column=1, value=f"Generated: {datetime.now().strftime('%d %b %Y, %H:%M')}")
        d.font = Font(color="64748B", size=9, name="Calibri")
        d.fill = PatternFill("solid", fgColor=LIGHT)
        d.alignment = Alignment(horizontal="right", vertical="center")
        r1 += 2

        # BOM header
        is_brf = req.equipment_type == "SNSF BRF"
        bom_cols = ["SECTION", "ITEM", "QTY", "UNIT", "UNIT PRICE", "COST PRICE",
                    "MARKUP" if is_brf else "", "SELL PRICE"]
        for ci, col in enumerate(bom_cols, 1):
            if col:
                hdr(ws, r1, ci, col, size=9)
        ws.row_dimensions[r1].height = 22
        r1 += 1

        last_section = ""
        for i, row_d in enumerate(req.bom):
            section = row_d.get("SECTION", "")
            if section != last_section:
                last_section = section
                ws.merge_cells(f"A{r1}:H{r1}")
                cell(ws, r1, 1, section, bold=True, bg=LIGHT, fg=NAVY)
                ws.row_dimensions[r1].height = 20
                r1 += 1

            bg = GREY if i % 2 == 0 else WHITE
            cell(ws, r1, 1, "", bg=bg)
            cell(ws, r1, 2, row_d.get("ITEM", ""), bg=bg)
            cell(ws, r1, 3, row_d.get("QTY", ""), bg=bg, align="right")
            cell(ws, r1, 4, row_d.get("UNIT", ""), bg=bg)
            cell(ws, r1, 5, row_d.get("UNIT PRICE", 0), bg=bg, align="right", num_fmt='#,##0')
            cell(ws, r1, 6, row_d.get("COST PRICE", 0), bg=bg, align="right", num_fmt='#,##0')
            if is_brf:
                cell(ws, r1, 7, f'{row_d.get("MARKUP", "")}x', bg=bg, align="center")
            cell(ws, r1, 8, row_d.get("SELL PRICE", 0), bg=bg, align="right", num_fmt='#,##0')
            ws.row_dimensions[r1].height = 18
            r1 += 1
        r1 += 1

        # Cost Summary
        cs = req.cost_summary
        ws.merge_cells(f"A{r1}:H{r1}")
        hdr(ws, r1, 1, "COST SUMMARY", size=10)
        ws.row_dimensions[r1].height = 20
        r1 += 1

        if req.equipment_type == "BTF":
            summary_rows = [
                ("Structure Cost",     cs.get("structure_cost", 0)),
                ("Combustion Cost",    cs.get("combustion_cost", 0)),
                ("Total Cost Price",   cs.get("total_cost", 0)),
                ("Sell Price",         cs.get("sell_price", 0)),
                ("Designing (10%)",    cs.get("designing_10pct", 0)),
                ("Negotiation (10%)",  cs.get("negotiation_10pct", 0)),
                ("Quoted Price",       cs.get("quoted_price", 0)),
            ]
        else:
            summary_rows = [
                ("Main Scope Cost",    cs.get("main_cost", 0)),
                ("Main Scope Sell",    cs.get("main_sell", 0)),
                ("NG Optional Cost",   cs.get("ng_optional_cost", 0)),
                ("NG Optional Sell",   cs.get("ng_optional_sell", 0)),
                ("Client Scope Cost",  cs.get("client_scope_cost", 0)),
                ("Client Scope Sell",  cs.get("client_scope_sell", 0)),
                ("Grand Total Cost",   cs.get("grand_cost", 0)),
                ("Grand Total Sell",   cs.get("grand_sell", 0)),
            ]

        for label, val in summary_rows:
            is_total = "Grand" in label or "Quoted" in label
            bg = GREEN_BG if is_total else GREY
            fg = GREEN if is_total else "1E293B"
            cell(ws, r1, 1, label, bold=is_total, bg=bg, fg=fg)
            ws.merge_cells(f"B{r1}:G{r1}")
            cell(ws, r1, 2, "", bg=bg)
            c = ws.cell(row=r1, column=8, value=val)
            c.font = Font(bold=is_total, color=fg, size=11 if is_total else 10, name="Calibri")
            c.fill = PatternFill("solid", fgColor=bg)
            c.alignment = Alignment(horizontal="right", vertical="center")
            c.number_format = '₹#,##0'
            c.border = thin()
            ws.row_dimensions[r1].height = 20
            r1 += 1

        # ── Additional sheets for BTF (match legacy 5-sheet workbook) ─────
        supp = (req.equipment or {}).get("supplementary", {})
        if supp and req.equipment_type == "BTF":
            def _kv_sheet(ws2, r2, items, col_widths=(36, 24)):
                """Write key-value rows to a sheet."""
                for ci, w in enumerate(col_widths, 1):
                    ws2.column_dimensions[chr(64+ci)].width = w
                return r2

            # ── Sheet: 10 T RHF ──────────────────────────────────────────
            ws_rhf = wb.create_sheet("10 T RHF", 0)
            ws_rhf.column_dimensions["A"].width = 36
            ws_rhf.column_dimensions["B"].width = 18
            ws_rhf.column_dimensions["C"].width = 14
            ws_rhf.column_dimensions["D"].width = 16
            ws_rhf.column_dimensions["E"].width = 14
            r2 = 1
            # Furnace dims
            fd = supp.get("furnace_dimensions", {})
            ws_rhf.merge_cells(f"A{r2}:E{r2}")
            hdr(ws_rhf, r2, 1, "FURNACE DIMENSIONS", size=10); r2 += 1
            cell(ws_rhf, r2, 1, "Furnace Internal Dim.", bold=True, bg=GREY)
            cell(ws_rhf, r2, 2, f'{fd.get("internal_L_mm","")}', bg=WHITE, align="right")
            cell(ws_rhf, r2, 3, f'{fd.get("internal_W_mm","")}', bg=WHITE, align="right")
            cell(ws_rhf, r2, 4, f'{fd.get("internal_H_mm","")} mm', bg=WHITE, align="right")
            r2 += 1
            for fv in supp.get("furnace_volumes", []):
                cell(ws_rhf, r2, 1, fv["item"], bold=True, bg=GREY)
                cell(ws_rhf, r2, 2, fv["value"], bg=WHITE, align="right")
                cell(ws_rhf, r2, 3, fv["unit"], bg=WHITE); r2 += 1
            r2 += 1
            # MS Structure
            ws_rhf.merge_cells(f"A{r2}:E{r2}")
            hdr(ws_rhf, r2, 1, "MS STRUCTURE", size=10); r2 += 1
            for sec in supp.get("ms_structure", []):
                cell(ws_rhf, r2, 1, sec["section"], bold=True, bg=LIGHT, fg=NAVY); r2 += 1
                for it in sec["items"]:
                    bg2 = GREEN_BG if it.get("highlight") else GREY
                    cell(ws_rhf, r2, 1, it["item"], bold=it.get("highlight",False), bg=bg2)
                    cell(ws_rhf, r2, 2, it["value"], bg=bg2, align="right"); r2 += 1
            r2 += 1
            # Ceramic Fibre
            ws_rhf.merge_cells(f"A{r2}:E{r2}")
            hdr(ws_rhf, r2, 1, "CERAMIC FIBRE", size=10); r2 += 1
            for cf in supp.get("ceramic_fibre", []):
                cell(ws_rhf, r2, 1, cf["wall"], bg=GREY)
                cell(ws_rhf, r2, 2, cf["L"], bg=WHITE, align="right")
                cell(ws_rhf, r2, 3, cf["W"], bg=WHITE, align="right")
                cell(ws_rhf, r2, 4, cf["thk"], bg=WHITE, align="right"); r2 += 1
            cfs = supp.get("ceramic_fibre_summary", {})
            cell(ws_rhf, r2, 1, "Total Rolls", bold=True, bg=GREEN_BG)
            cell(ws_rhf, r2, 2, cfs.get("total_rolls",""), bg=GREEN_BG, align="right"); r2 += 2
            # Door
            ws_rhf.merge_cells(f"A{r2}:E{r2}")
            hdr(ws_rhf, r2, 1, "DOOR", size=10); r2 += 1
            hdr(ws_rhf, r2, 1, "Item", size=9); hdr(ws_rhf, r2, 2, "Qty", size=9); hdr(ws_rhf, r2, 3, "Rate", size=9); hdr(ws_rhf, r2, 4, "Total", size=9); r2 += 1
            door_total = 0
            for d in supp.get("door", []):
                cell(ws_rhf, r2, 1, d["item"], bg=GREY)
                cell(ws_rhf, r2, 2, d["qty"], bg=WHITE, align="right")
                cell(ws_rhf, r2, 3, d["rate"], bg=WHITE, align="right", num_fmt='#,##0')
                cell(ws_rhf, r2, 4, d["total"], bg=WHITE, align="right", num_fmt='#,##0')
                door_total += d["total"]; r2 += 1
            cell(ws_rhf, r2, 1, "Total Door", bold=True, bg=GREEN_BG)
            cell(ws_rhf, r2, 4, door_total, bg=GREEN_BG, align="right", num_fmt='#,##0'); r2 += 2
            # Trolley
            ws_rhf.merge_cells(f"A{r2}:E{r2}")
            hdr(ws_rhf, r2, 1, "TROLLEY / BOGIE", size=10); r2 += 1
            hdr(ws_rhf, r2, 1, "Item", size=9); hdr(ws_rhf, r2, 2, "Qty", size=9); hdr(ws_rhf, r2, 3, "Rate", size=9); hdr(ws_rhf, r2, 4, "Total", size=9); r2 += 1
            tr_total = 0
            for t in supp.get("trolley_bogie", []):
                cell(ws_rhf, r2, 1, t["item"], bg=GREY)
                cell(ws_rhf, r2, 2, t["qty"], bg=WHITE, align="right")
                cell(ws_rhf, r2, 3, t["rate"], bg=WHITE, align="right", num_fmt='#,##0')
                cell(ws_rhf, r2, 4, t["total"], bg=WHITE, align="right", num_fmt='#,##0')
                tr_total += t["total"]; r2 += 1
            cell(ws_rhf, r2, 1, "Total Trolley/Bogie", bold=True, bg=GREEN_BG)
            cell(ws_rhf, r2, 4, tr_total, bg=GREEN_BG, align="right", num_fmt='#,##0'); r2 += 2
            # Refractory
            ws_rhf.merge_cells(f"A{r2}:E{r2}")
            hdr(ws_rhf, r2, 1, "REFRACTORY", size=10); r2 += 1
            ref = supp.get("refractory", {})
            for sec_name, sec_key in [("Walls","walls"),("Bogie","bogie")]:
                cell(ws_rhf, r2, 1, sec_name, bold=True, bg=LIGHT, fg=NAVY); r2 += 1
                sec_total = 0
                for br in ref.get(sec_key, []):
                    cell(ws_rhf, r2, 1, br["type"], bg=GREY)
                    cell(ws_rhf, r2, 2, br["qty"], bg=WHITE, align="right")
                    cell(ws_rhf, r2, 3, br.get("cost_per",""), bg=WHITE, align="right")
                    cell(ws_rhf, r2, 4, br["total"], bg=WHITE, align="right", num_fmt='#,##0')
                    sec_total += br["total"]; r2 += 1
                cell(ws_rhf, r2, 1, f"Subtotal {sec_name}", bold=True, bg=GREEN_BG)
                cell(ws_rhf, r2, 4, sec_total, bg=GREEN_BG, align="right", num_fmt='#,##0'); r2 += 1
            for sec_name, sec_key in [("Misc (Mortar)","misc"),("Arch Bricks","arch")]:
                cell(ws_rhf, r2, 1, sec_name, bold=True, bg=LIGHT, fg=NAVY); r2 += 1
                for br in ref.get(sec_key, []):
                    cell(ws_rhf, r2, 1, br["type"], bg=GREY)
                    cell(ws_rhf, r2, 2, br["qty"], bg=WHITE, align="right")
                    cell(ws_rhf, r2, 3, br.get("rate", br.get("cost_per","")), bg=WHITE, align="right")
                    cell(ws_rhf, r2, 4, br["total"], bg=WHITE, align="right", num_fmt='#,##0'); r2 += 1

            # ── Sheet: Calculation ────────────────────────────────────────
            ws_calc = wb.create_sheet("Calculation", 1)
            ws_calc.column_dimensions["A"].width = 28
            ws_calc.column_dimensions["B"].width = 20
            ws_calc.column_dimensions["C"].width = 14
            r2 = 1
            fp = supp.get("furnace_params", {})
            ws_calc.merge_cells(f"A{r2}:C{r2}")
            hdr(ws_calc, r2, 1, "HEAT LOAD", size=10); r2 += 1
            for h in supp.get("heat_load", []):
                is_t = h["item"] in ("Total", "Gross")
                bg2 = GREEN_BG if is_t else GREY
                cell(ws_calc, r2, 1, h["item"], bold=is_t, bg=bg2)
                cell(ws_calc, r2, 2, h["value"], bg=bg2, align="right", num_fmt='#,##0')
                cell(ws_calc, r2, 3, h["unit"], bg=bg2); r2 += 1
            r2 += 1
            ws_calc.merge_cells(f"A{r2}:C{r2}")
            hdr(ws_calc, r2, 1, "FURNACE PARAMETERS", size=10); r2 += 1
            for label, val in [
                ("Capacity", f'{fp.get("furnace_capacity_tph","")} Ton/hr'),
                ("Fuel Consumption", f'{fp.get("fuel_consumption_nm3hr","")} Nm3/hr'),
                ("Air Flow", f'{fp.get("air_flow_nm3hr","")} Nm3/hr'),
                ("CFM", f'{fp.get("cfm","")}'),
                ("Blower HP", f'{fp.get("blower_hp_calc","")} → {fp.get("blower_hp_selected","")} HP'),
                ("Burners", f'{fp.get("no_of_burners","")} ({fp.get("no_of_zones","")} zones)'),
                ("Rating/Zone", f'{fp.get("rating_per_zone_kcal","")} kcal ({fp.get("rating_per_zone_kw","")} kW)'),
            ]:
                cell(ws_calc, r2, 1, label, bold=True, bg=GREY)
                cell(ws_calc, r2, 2, val, bg=WHITE); r2 += 1
            r2 += 1
            ws_calc.merge_cells(f"A{r2}:C{r2}")
            hdr(ws_calc, r2, 1, "PIPE SIZING", size=10); r2 += 1
            ps = supp.get("pipe_sizing", {})
            for label, key in [("Air Zone 1","air_zone1"),("Air Zone 2","air_zone2"),("Gas Zone 1","gas_zone1"),("Gas Zone 2","gas_zone2")]:
                p = ps.get(key, {})
                cell(ws_calc, r2, 1, label, bold=True, bg=GREY)
                cell(ws_calc, r2, 2, f'{p.get("flow_nm3hr","")} Nm3/hr @ {p.get("velocity_ms","")} m/s', bg=WHITE)
                cell(ws_calc, r2, 3, f'd = {p.get("inner_dia_mm","")} mm', bg=WHITE); r2 += 1
            r2 += 1
            ws_calc.merge_cells(f"A{r2}:C{r2}")
            hdr(ws_calc, r2, 1, "GAS CONSUMPTION", size=10); r2 += 1
            for g in supp.get("gas_consumption", []):
                bg2 = GREEN_BG if "Guarantee" in g["item"] else GREY
                cell(ws_calc, r2, 1, g["item"], bold="Guarantee" in g["item"], bg=bg2)
                cell(ws_calc, r2, 2, g["value"], bg=bg2); r2 += 1
            r2 += 1
            ws_calc.merge_cells(f"A{r2}:C{r2}")
            hdr(ws_calc, r2, 1, "GEAR BOX", size=10); r2 += 1
            for g in supp.get("gear_box", []):
                cell(ws_calc, r2, 1, g["item"], bold=True, bg=GREY)
                cell(ws_calc, r2, 2, g["value"], bg=WHITE); r2 += 1

            # ── Sheet: Recuperator ────────────────────────────────────────
            ws_rec = wb.create_sheet("Recuperator", 2)
            ws_rec.column_dimensions["A"].width = 36
            ws_rec.column_dimensions["B"].width = 18
            ws_rec.column_dimensions["C"].width = 14
            r2 = 1
            rcp = supp.get("recuperator", {})
            ws_rec.merge_cells(f"A{r2}:C{r2}")
            hdr(ws_rec, r2, 1, "RECUPERATOR CALCULATION", size=10); r2 += 1
            rec_rows = [
                ("Total Flue Gas", f'{rcp.get("total_flue_gas_nm3hr","")}', "Nm3/Hr"),
                ("Total Mass of Flue Gas", f'{rcp.get("total_mass_flue_gas_kghr","")}', "Kg/Hr"),
                ("Specific Heat of Flue Gas", f'{rcp.get("specific_heat_flue_gas","")}', "Kcal/Kg-C"),
                ("Initial Temp of Flue Gas", f'{rcp.get("initial_temp_flue_gas_c","")}', "°C"),
                ("Final Temp of Flue Gas", f'{rcp.get("final_temp_flue_gas_c","")}', "°C"),
                ("Heat Transfer Coefficient", f'{rcp.get("heat_transfer_coeff","")}', "Kcal/m2-C"),
                ("Volume of Combustion Air", f'{rcp.get("air_volume_nm3hr","")}', "Nm3/Hr"),
                ("Initial Temp of Air", f'{rcp.get("initial_air_temp_c","")}', "°C"),
                ("Final Temp of Air", f'{rcp.get("final_air_temp_c","")}', "°C"),
                ("Specific Heat of Air", f'{rcp.get("specific_heat_air","")}', "Kcal/Kg-C"),
                ("Heat Required By Combustion Air", f'{rcp.get("heat_required_kcal","")}', "Kcal"),
                ("Logarithmic Mean Temp Diff", f'{rcp.get("lmtd_c","")}', "°C"),
                ("Surface Area of Recuperator", f'{rcp.get("surface_area_m2","")}', "mtr2"),
                ("Pipes in Row × Column", f'{rcp.get("pipes_per_row","")} × {rcp.get("pipes_per_column","")}', f'= {rcp.get("pipes_total","")}'),
            ]
            for label, val, unit in rec_rows:
                cell(ws_rec, r2, 1, label, bold=True, bg=GREY)
                cell(ws_rec, r2, 2, val, bg=WHITE, align="right")
                cell(ws_rec, r2, 3, unit, bg=WHITE); r2 += 1
            r2 += 1
            ws_rec.merge_cells(f"A{r2}:C{r2}")
            hdr(ws_rec, r2, 1, "HOT BANK (SS 304)", size=10); r2 += 1
            for label, val, unit in [
                ("Pipe Diameter", rcp.get("pipe_dia_mm",""), "mm"),
                ("Pipe Length", rcp.get("pipe_length_m",""), "mtr"),
                ("Pipe Thickness", rcp.get("pipe_thickness_mm",""), "mm"),
                ("Weight per Pipe", rcp.get("pipe_weight_kg_m",""), "Kg/Mtr"),
                ("Total Weight Hot Bank", rcp.get("hot_bank_weight_kg",""), "Kg"),
            ]:
                cell(ws_rec, r2, 1, label, bg=GREY)
                cell(ws_rec, r2, 2, val, bg=WHITE, align="right")
                cell(ws_rec, r2, 3, unit, bg=WHITE); r2 += 1
            r2 += 1
            ws_rec.merge_cells(f"A{r2}:C{r2}")
            hdr(ws_rec, r2, 1, f'COLD BANK ({rcp.get("cold_bank_material","")})', size=10); r2 += 1
            for label, val, unit in [
                ("Pipe Diameter", rcp.get("cold_bank_dia_mm",""), "mm"),
                ("Pipe Length", rcp.get("cold_bank_length_m",""), "mtr"),
                ("Pipe Thickness", rcp.get("cold_bank_thickness_mm",""), "mm"),
                ("Weight per Pipe", rcp.get("cold_bank_weight_kg_m",""), "Kg/Mtr"),
                ("Total Weight Cold Bank", rcp.get("cold_bank_total_wt_kg",""), "Kg"),
            ]:
                cell(ws_rec, r2, 1, label, bg=GREY)
                cell(ws_rec, r2, 2, val, bg=WHITE, align="right")
                cell(ws_rec, r2, 3, unit, bg=WHITE); r2 += 1
            r2 += 1
            ws_rec.merge_cells(f"A{r2}:C{r2}")
            hdr(ws_rec, r2, 1, "COST OF RECUPERATOR", size=10); r2 += 1
            cost_items = [
                ("Cost of All Pipes", rcp.get("cost_all_pipes",0)),
                ("MS Outer Shell", rcp.get("cost_ms_outer_shell",0)),
                ("MS Combustion Air Inlet", rcp.get("cost_ms_combustion_air_inlet",0)),
                ("MS Channel 150×75×10", rcp.get("cost_ms_channel_150x75",0)),
                ("Angle 65×25 mtr", rcp.get("cost_angle_65",0)),
                ("Angle 75×10 mtr", rcp.get("cost_angle_75",0)),
                ("Angle 50×15 mtr", rcp.get("cost_angle_50",0)),
            ]
            for label, val in cost_items:
                cell(ws_rec, r2, 1, label, bg=GREY)
                cell(ws_rec, r2, 2, val, bg=WHITE, align="right", num_fmt='#,##0')
                cell(ws_rec, r2, 3, "Rs", bg=WHITE); r2 += 1
            cell(ws_rec, r2, 1, "Total Material Cost", bold=True, bg=GREEN_BG)
            cell(ws_rec, r2, 2, rcp.get("cost_total_material",0), bg=GREEN_BG, align="right", num_fmt='#,##0'); r2 += 1
            for label, val in [
                ("Bending of Pipes", rcp.get("cost_pipe_bending",0)),
                ("Welding Rod", rcp.get("cost_welding_rod",0)),
                ("Hole Fabrication", rcp.get("cost_hole_fabrication",0)),
                ("Thermocouple with TT", rcp.get("cost_thermocouple_tt",0)),
            ]:
                cell(ws_rec, r2, 1, label, bg=GREY)
                cell(ws_rec, r2, 2, val, bg=WHITE, align="right", num_fmt='#,##0')
                cell(ws_rec, r2, 3, "Rs", bg=WHITE); r2 += 1
            cell(ws_rec, r2, 1, "Total Cost of Recuperator", bold=True, bg=GREEN_BG, fg=GREEN)
            cell(ws_rec, r2, 2, rcp.get("cost_total_recuperator",0), bg=GREEN_BG, align="right", num_fmt='₹#,##0')

            # ── Sheet: Combustion ─────────────────────────────────────
            comb_mode = cs.get("combustion_mode", "onoff")
            comb_title = "Combustion-Massflow" if comb_mode == "massflow" else "Combustion-ONOFF"
            ws_comb = wb.create_sheet(comb_title, 3)
            ws_comb.column_dimensions["A"].width = 8
            ws_comb.column_dimensions["B"].width = 40
            ws_comb.column_dimensions["C"].width = 10
            ws_comb.column_dimensions["D"].width = 18
            r2 = 1
            ws_comb.merge_cells(f"A{r2}:D{r2}")
            hdr(ws_comb, r2, 1, f"COMBUSTION SYSTEM ({comb_title})", size=10); r2 += 1
            hdr(ws_comb, r2, 1, "S.No.", size=9); hdr(ws_comb, r2, 2, "Control System", size=9)
            hdr(ws_comb, r2, 3, "Qty", size=9); hdr(ws_comb, r2, 4, "Total Price", size=9); r2 += 1
            comb_items = [b for b in req.bom if b.get("SECTION") == "Combustion System"]
            comb_total = 0
            for i, ci in enumerate(comb_items, 1):
                cell(ws_comb, r2, 1, i, bg=GREY, align="right")
                cell(ws_comb, r2, 2, ci.get("ITEM",""), bg=WHITE)
                cell(ws_comb, r2, 3, ci.get("QTY",""), bg=WHITE, align="right")
                cell(ws_comb, r2, 4, ci.get("COST PRICE",0), bg=WHITE, align="right", num_fmt='#,##0')
                comb_total += ci.get("COST PRICE",0); r2 += 1
            cell(ws_comb, r2, 2, "TOTAL", bold=True, bg=GREEN_BG)
            cell(ws_comb, r2, 4, comb_total, bold=True, bg=GREEN_BG, align="right", num_fmt='#,##0')

            # Rename BOM sheet to legacy costing name
            costing_title = "Costing with Mass flow" if comb_mode == "massflow" else "Costing with Pulse Firing"
            ws.title = costing_title

        # ── Additional sheets for SNSF BRF (match legacy 7-sheet workbook) ──
        if supp and req.equipment_type == "SNSF BRF":
            # Sheet: Furnace (2) — full 110 rows from legacy
            ws2 = wb.create_sheet("Furnace (2)", 0)
            for ci, w in enumerate([4, 44, 18, 14, 4, 44, 18, 14, 4, 44], 1):
                ws2.column_dimensions[chr(64+ci)].width = w
            SECTION_KW = ['REHEATING','1.','2.','FLUE DUCT','STRUCTURAL','CASTING','PIPE CALC','PREHEATING']
            r2 = 0
            for row_data in supp.get("furnace_full", []):
                r2 += 1
                first = str(row_data[0] if row_data else "").strip()
                is_sec = any(first.upper().startswith(k) for k in SECTION_KW)
                is_hl = first.startswith("Total") or first.startswith("EFFECTIVE") or first.startswith("Overall")
                if is_sec:
                    bg2 = LIGHT
                elif is_hl:
                    bg2 = GREEN_BG
                else:
                    bg2 = GREY if r2 % 2 == 0 else WHITE
                for ci, v in enumerate(row_data):
                    val = v if v != '' else None
                    is_num = isinstance(val, (int, float))
                    cell(ws2, r2, ci+1, val,
                         bold=is_sec or is_hl,
                         bg=bg2,
                         fg=NAVY if is_sec else ("375623" if is_hl else "1E293B"),
                         align="right" if is_num else "left",
                         num_fmt='#,##0.##' if is_num else None)

            # Sheet: Refractory
            ws3 = wb.create_sheet("Refractory", 1)
            ws3.column_dimensions["A"].width = 44; ws3.column_dimensions["B"].width = 12; ws3.column_dimensions["C"].width = 14; ws3.column_dimensions["D"].width = 12; ws3.column_dimensions["E"].width = 16
            r2 = 1; ws3.merge_cells(f"A{r2}:E{r2}"); hdr(ws3, r2, 1, "REFRACTORY WITH MAHA KOSAL/BHILWARA/TRL", size=10); r2 += 1
            hdr(ws3, r2, 1, "Item", size=9); hdr(ws3, r2, 2, "Qty", size=9); hdr(ws3, r2, 3, "Weight", size=9); hdr(ws3, r2, 4, "Rate", size=9); hdr(ws3, r2, 5, "Cost", size=9); r2 += 1
            ref_total = 0
            for ri in supp.get("refractory_items", []):
                cell(ws3, r2, 1, ri["item"], bg=GREY); cell(ws3, r2, 2, ri["qty"], bg=WHITE, align="right")
                cell(ws3, r2, 3, ri["weight_kg"], bg=WHITE, align="right", num_fmt='#,##0')
                cell(ws3, r2, 4, ri["rate"], bg=WHITE, align="right"); cell(ws3, r2, 5, ri["cost"], bg=WHITE, align="right", num_fmt='#,##0')
                ref_total += ri["cost"]; r2 += 1
            for ri in supp.get("refractory_extra", []):
                cell(ws3, r2, 1, ri["item"], bg=GREY); cell(ws3, r2, 2, ri["qty"], bg=WHITE, align="right")
                cell(ws3, r2, 3, ri["wt"], bg=WHITE, align="right", num_fmt='#,##0')
                cell(ws3, r2, 4, ri["rate"], bg=WHITE, align="right"); cell(ws3, r2, 5, ri["cost"], bg=WHITE, align="right", num_fmt='#,##0')
                ref_total += ri["cost"]; r2 += 1
            cell(ws3, r2, 1, "Total Refractory", bold=True, bg=GREEN_BG); cell(ws3, r2, 5, ref_total, bold=True, bg=GREEN_BG, align="right", num_fmt='#,##0')

            # Sheet: Mild Steel
            ws4 = wb.create_sheet("Mild Steel", 2)
            ws4.column_dimensions["A"].width = 44; ws4.column_dimensions["B"].width = 14; ws4.column_dimensions["C"].width = 12; ws4.column_dimensions["D"].width = 14; ws4.column_dimensions["E"].width = 16
            r2 = 1; ws4.merge_cells(f"A{r2}:E{r2}"); hdr(ws4, r2, 1, "MILD STEEL", size=10); r2 += 1
            hdr(ws4, r2, 1, "Description", size=9); hdr(ws4, r2, 2, "Qty", size=9); hdr(ws4, r2, 3, "Wt/qty", size=9); hdr(ws4, r2, 4, "Total Wt", size=9); hdr(ws4, r2, 5, "Cost", size=9); r2 += 1
            for mi in supp.get("mild_steel", []):
                cell(ws4, r2, 1, mi["item"], bg=GREY); cell(ws4, r2, 2, mi["qty_m"], bg=WHITE, align="right")
                cell(ws4, r2, 3, mi["wt_per_m"], bg=WHITE, align="right"); cell(ws4, r2, 4, mi["total_wt"], bg=WHITE, align="right", num_fmt='#,##0')
                cell(ws4, r2, 5, mi["cost"], bg=WHITE, align="right", num_fmt='#,##0'); r2 += 1
            cell(ws4, r2, 1, "Total Structure", bold=True, bg=GREEN_BG)
            cell(ws4, r2, 4, supp.get("mild_steel_total_wt", 0), bg=GREEN_BG, align="right", num_fmt='#,##0')
            cell(ws4, r2, 5, supp.get("mild_steel_total_cost", 0), bg=GREEN_BG, align="right", num_fmt='#,##0'); r2 += 2
            ws4.merge_cells(f"A{r2}:E{r2}"); hdr(ws4, r2, 1, "PIPELINE & CASTING", size=10); r2 += 1
            for pi in supp.get("pipeline_casting", []):
                cell(ws4, r2, 1, pi["item"], bg=GREY); cell(ws4, r2, 3, pi["wt"], bg=WHITE, align="right")
                cell(ws4, r2, 4, pi["rate"], bg=WHITE, align="right"); cell(ws4, r2, 5, pi["cost"], bg=WHITE, align="right", num_fmt='#,##0'); r2 += 1

            # Sheet: Mass Flow Control
            ws5 = wb.create_sheet("Mass Flow Control", 3)
            ws5.column_dimensions["A"].width = 44; ws5.column_dimensions["B"].width = 10; ws5.column_dimensions["C"].width = 16
            r2 = 1; ws5.merge_cells(f"A{r2}:C{r2}"); hdr(ws5, r2, 1, "MASS FLOW CONTROL SYSTEM", size=10); r2 += 1
            mfc_total = 0
            for z in supp.get("mass_flow_control", []):
                cell(ws5, r2, 1, z["zone"], bold=True, bg=LIGHT, fg=NAVY); r2 += 1
                for it in z["items"]:
                    cell(ws5, r2, 1, it["item"], bg=GREY); cell(ws5, r2, 2, it["qty"], bg=WHITE, align="right")
                    t = it["qty"] * it["price"]; cell(ws5, r2, 3, t, bg=WHITE, align="right", num_fmt='#,##0'); mfc_total += t; r2 += 1
            cell(ws5, r2, 1, "TOTAL", bold=True, bg=GREEN_BG); cell(ws5, r2, 3, mfc_total, bold=True, bg=GREEN_BG, align="right", num_fmt='#,##0')

            # Sheet: Combustion
            ws6 = wb.create_sheet("Combustion", 4)
            ws6.column_dimensions["A"].width = 44; ws6.column_dimensions["B"].width = 10; ws6.column_dimensions["C"].width = 18; ws6.column_dimensions["D"].width = 18
            r2 = 1; ws6.merge_cells(f"A{r2}:D{r2}"); hdr(ws6, r2, 1, "COMBUSTION EQUIPMENT", size=10); r2 += 1
            hdr(ws6, r2, 1, "Description", size=9); hdr(ws6, r2, 2, "Qty", size=9); hdr(ws6, r2, 3, "Unit Price", size=9); hdr(ws6, r2, 4, "Cost Price", size=9); r2 += 1
            comb_total2 = 0
            for ci in supp.get("combustion_items", []):
                t = ci["qty"] * ci["price"]
                cell(ws6, r2, 1, ci["item"], bg=GREY); cell(ws6, r2, 2, ci["qty"], bg=WHITE, align="right")
                cell(ws6, r2, 3, ci["price"], bg=WHITE, align="right", num_fmt='#,##0')
                cell(ws6, r2, 4, t, bg=WHITE, align="right", num_fmt='#,##0'); comb_total2 += t; r2 += 1
            cell(ws6, r2, 1, "TOTAL", bold=True, bg=GREEN_BG); cell(ws6, r2, 4, comb_total2, bold=True, bg=GREEN_BG, align="right", num_fmt='#,##0')

            # Sheet: Recuperator1
            ws7 = wb.create_sheet("Recuperator1", 5)
            ws7.column_dimensions["A"].width = 40; ws7.column_dimensions["B"].width = 18; ws7.column_dimensions["C"].width = 14
            r2 = 1; rcp2 = supp.get("recuperator", {}); rcc = supp.get("recuperator_cost", {})
            ws7.merge_cells(f"A{r2}:C{r2}"); hdr(ws7, r2, 1, rcp2.get("title", "Recuperator Calculation"), size=10); r2 += 1
            for label, val, unit in [
                ("Total Flue Gas", rcp2.get("total_flue_gas_nm3hr",""), "Nm3/Hr"),
                ("Total Mass of Flue Gas", rcp2.get("total_mass_flue_gas_kghr",""), "Kg/Hr"),
                ("Initial Temp of Flue Gas", rcp2.get("initial_temp_flue_gas_c",""), "°C"),
                ("Final Temp of Flue Gas", rcp2.get("final_temp_flue_gas_c",""), "°C"),
                ("Air Volume", rcp2.get("air_volume_nm3hr",""), "Nm3/Hr"),
                ("Heat Required", rcp2.get("heat_required_kcal",""), "Kcal"),
                ("LMTD", rcp2.get("lmtd_c",""), "°C"),
                ("Surface Area", rcp2.get("surface_area_m2",""), "mtr2"),
                ("Pipes", f'{rcp2.get("pipes_per_row","")}×{rcp2.get("pipes_per_column","")} = {rcp2.get("pipes_total","")}', ""),
                ("Hot Bank Weight", rcp2.get("hot_bank_weight_kg",""), "Kg"),
                ("Cold Bank Weight", rcc.get("cold_bank_total_wt_kg",""), "Kg"),
            ]:
                cell(ws7, r2, 1, label, bold=True, bg=GREY); cell(ws7, r2, 2, val, bg=WHITE, align="right"); cell(ws7, r2, 3, unit, bg=WHITE); r2 += 1
            r2 += 1; ws7.merge_cells(f"A{r2}:C{r2}"); hdr(ws7, r2, 1, "COST OF RECUPERATOR", size=10); r2 += 1
            for label, val in [
                ("Cost of All Pipes", rcc.get("cost_all_pipes",0)),("MS Outer Shell", rcc.get("cost_ms_outer_shell",0)),
                ("MS Combustion Air Inlet", rcc.get("cost_ms_combustion_air_inlet",0)),
                ("MS Channel 150×75", rcc.get("cost_ms_channel_150x75",0)),
                ("Angle 65×25", rcc.get("cost_angle_65",0)),("Angle 75×10", rcc.get("cost_angle_75",0)),("Angle 50×15", rcc.get("cost_angle_50",0)),
            ]:
                cell(ws7, r2, 1, label, bg=GREY); cell(ws7, r2, 2, val, bg=WHITE, align="right", num_fmt='#,##0'); cell(ws7, r2, 3, "Rs", bg=WHITE); r2 += 1
            cell(ws7, r2, 1, "Total Material Cost", bold=True, bg=GREEN_BG)
            cell(ws7, r2, 2, rcc.get("cost_total_material",0), bg=GREEN_BG, align="right", num_fmt='#,##0'); r2 += 1
            for label, val in [("Bending",rcc.get("cost_pipe_bending",0)),("Welding Rod",rcc.get("cost_welding_rod",0)),
                ("Hole Fabrication",rcc.get("cost_hole_fabrication",0)),("Thermocouple",rcc.get("cost_thermocouple_tt",0))]:
                cell(ws7, r2, 1, label, bg=GREY); cell(ws7, r2, 2, val, bg=WHITE, align="right", num_fmt='#,##0'); r2 += 1
            cell(ws7, r2, 1, "Total Cost of Recuperator", bold=True, bg=GREEN_BG, fg=GREEN)
            cell(ws7, r2, 2, rcc.get("cost_total_recuperator",0), bg=GREEN_BG, align="right", num_fmt='₹#,##0')

            # Rename BOM sheet
            ws.title = "Breakup"

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        cname = req.customer.get("company_name", "").replace(" ", "_") or "ENCON"
        date_str = datetime.now().strftime("%d%b%Y")
        fname = f"{req.equipment_type}_Costing_{cname}_{date_str}.xlsx"
        return StreamingResponse(
            buf,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{fname}"'})

    # ════════════════════════════════════════════════════════════════════════
    #  VLPH / HLPH path
    # ════════════════════════════════════════════════════════════════════════
    ws = wb.active
    ws.title = req.equipment_type
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 18
    ws.column_dimensions["F"].width = 18
    ws.column_dimensions["G"].width = 18
    ws.column_dimensions["H"].width = 18

    r = 1
    # ── Title ──────────────────────────────────────────────────────────────
    ws.merge_cells(f"A{r}:H{r}")
    t = ws.cell(row=r, column=1, value=f"ENCON — {req.equipment_type} Costing Sheet")
    t.font = Font(bold=True, color=WHITE, size=14, name="Calibri")
    t.fill = PatternFill("solid", fgColor=NAVY)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[r].height = 30
    r += 1

    # Date
    ws.merge_cells(f"A{r}:H{r}")
    d = ws.cell(row=r, column=1, value=f"Generated: {datetime.now().strftime('%d %b %Y, %H:%M')}")
    d.font = Font(color="64748B", size=9, name="Calibri")
    d.fill = PatternFill("solid", fgColor=LIGHT)
    d.alignment = Alignment(horizontal="right", vertical="center")
    r += 2

    # ── Customer ──────────────────────────────────────────────────────────
    if req.customer:
        ws.merge_cells(f"A{r}:H{r}")
        hdr(ws, r, 1, "CUSTOMER DETAILS", size=10)
        ws.row_dimensions[r].height = 20
        r += 1
        cust_fields = [
            ("Company",     req.customer.get("company_name","")),
            ("Contact",     req.customer.get("poc_name","")),
            ("Designation", req.customer.get("poc_designation","")),
            ("Mobile",      req.customer.get("mobile_no","")),
            ("Email",       req.customer.get("email","")),
            ("Project",     req.customer.get("project_name","")),
            ("Ref No.",     req.customer.get("ref_no","")),
        ]
        for label, val in cust_fields:
            if val:
                cell(ws, r, 1, label, bold=True, bg=GREY)
                ws.merge_cells(f"B{r}:H{r}")
                cell(ws, r, 2, val, bg=WHITE)
                r += 1
        r += 1

    # ── Process Parameters ────────────────────────────────────────────────
    ws.merge_cells(f"A{r}:H{r}")
    hdr(ws, r, 1, "PROCESS PARAMETERS", size=10)
    ws.row_dimensions[r].height = 20
    r += 1

    calc = req.calculations
    params = [
        ("Initial Temp (Ti)",       f"{calc.get('Ti','')} °C"),
        ("Final Temp (Tf)",         f"{calc.get('Tf','')} °C"),
        ("Refractory Weight",       f"{calc.get('refractory_weight','')} kg"),
        ("Fuel CV",                 f"{calc.get('fuel_cv','')} kcal/Nm³"),
        ("Time Taken",              f"{calc.get('time_taken_hr','')} hr"),
        ("Heat Load",               f"{calc.get('heat_load_kcal','')} kcal"),
        ("Firing Rate",             f"{calc.get('firing_rate_kcal','')} kcal/hr"),
        ("Fuel Consumption",        f"{calc.get('fuel_consumption_nm3','')} Nm³"),
        ("Calc. Firing Rate",       f"{calc.get('calculated_firing_rate_nm3hr','')} Nm³/hr"),
        ("Design Firing Rate",      f"{calc.get('extra_firing_rate_nm3hr','')} Nm³/hr"),
    ]
    for i, (label, val) in enumerate(params):
        bg = GREY if i % 2 == 0 else WHITE
        cell(ws, r, 1, label, bold=True, bg=bg)
        ws.merge_cells(f"B{r}:H{r}")
        cell(ws, r, 2, val, bg=bg)
        r += 1
    r += 1

    # ── BOM Table ─────────────────────────────────────────────────────────
    ws.merge_cells(f"A{r}:H{r}")
    hdr(ws, r, 1, "BILL OF MATERIALS", size=10)
    ws.row_dimensions[r].height = 20
    r += 1

    bom_cols = ["MEDIA", "ITEM NAME", "REFERENCE", "QTY", "UNIT PRICE", "TOTAL", "", ""]
    for ci, col in enumerate(bom_cols, 1):
        hdr(ws, r, ci, col, size=9)
    ws.row_dimensions[r].height = 22
    r += 1

    for i, row_d in enumerate(req.bom):
        bg = GREY if i % 2 == 0 else WHITE
        vals = list(row_d.values())
        for ci, v in enumerate(vals[:8], 1):
            num = ci >= 4
            cell(ws, r, ci, v, bg=bg, align="right" if num else "left",
                 num_fmt='#,##0.00' if isinstance(v, (int, float)) and num else None)
        ws.row_dimensions[r].height = 18
        r += 1
    r += 1

    # ── Cost Summary ──────────────────────────────────────────────────────
    cs = req.cost_summary
    ws.merge_cells(f"A{r}:H{r}")
    hdr(ws, r, 1, "COST SUMMARY", size=10)
    ws.row_dimensions[r].height = 20
    r += 1

    summary_rows = [
        ("Bought Out Total", cs.get("bought_out_total", 0)),
        ("ENCON Total",      cs.get("encon_total", 0)),
        ("Grand Total",      cs.get("grand_total", 0)),
    ]
    for label, val in summary_rows:
        is_total = label == "Grand Total"
        bg = GREEN_BG if is_total else GREY
        fg = GREEN if is_total else "1E293B"
        cell(ws, r, 1, label, bold=is_total, bg=bg, fg=fg)
        ws.merge_cells(f"B{r}:G{r}")
        cell(ws, r, 2, "", bg=bg)
        c = ws.cell(row=r, column=8, value=val)
        c.font = Font(bold=is_total, color=fg, size=11 if is_total else 10, name="Calibri")
        c.fill = PatternFill("solid", fgColor=bg)
        c.alignment = Alignment(horizontal="right", vertical="center")
        c.number_format = '₹#,##0.00'
        c.border = thin()
        ws.row_dimensions[r].height = 20
        r += 1

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    equip = req.equipment_type
    cname = req.customer.get("company_name", "").replace(" ", "_") or "ENCON"
    date_str = datetime.now().strftime("%d%b%Y")
    fname = f"{equip}_Costing_{cname}_{date_str}.xlsx"

    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{fname}"'}
    )


@app.get("/api/download-quote/{filename}")
def download_quote(filename: str):
    file_path = os.path.join(QUOTES_FOLDER, filename)
    if not os.path.exists(file_path):
        return {"error": "File not found"}
    return FileResponse(path=file_path, filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.post("/upload-excel/")
async def upload_excel(request: Request, file: UploadFile = File(...)):
    try:
        filename = file.filename
        table_name = filename.replace(".xlsx", "").strip().lower()
        if VALID_TABLES and table_name not in VALID_TABLES:
            return {"error": f"Invalid table name: {table_name}"}
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        df = pd.read_excel(file_path)
        if df.empty:
            return {"error": "Excel file is empty"}
        df.columns = [col.strip().lower() for col in df.columns]
        client_ip = request.client.host if request.client else "unknown"
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        conn = sqlite3.connect(DB_PATH)
        df.to_sql(table_name, conn, if_exists="replace", index=False)
        conn.execute("""INSERT INTO table_update_log (table_name, updated_at, uploaded_by)
            VALUES (?, ?, ?) ON CONFLICT(table_name) DO UPDATE SET
            updated_at = excluded.updated_at, uploaded_by = excluded.uploaded_by""",
            (table_name, now, client_ip))
        conn.commit()
        conn.close()
        return {"message": f"{table_name} updated successfully!", "updated_at": now}
    except Exception as e:
        return {"error": str(e)}


@app.get("/db/tables")
def get_tables():
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("""SELECT name FROM sqlite_master
            WHERE type='table' AND name NOT LIKE 'sqlite_%' AND name != 'table_update_log'
            ORDER BY name""")
        tables = [row[0] for row in cursor.fetchall()]
        cursor.execute("SELECT table_name, updated_at, uploaded_by FROM table_update_log")
        log = {row[0]: {"updated_at": row[1], "uploaded_by": row[2]} for row in cursor.fetchall()}
        result = []
        for table in tables:
            cursor.execute(f"SELECT COUNT(*) FROM [{table}]")
            count = cursor.fetchone()[0]
            entry = log.get(table, {})
            result.append({"name": table, "rows": count,
                "updated_at": entry.get("updated_at"), "uploaded_by": entry.get("uploaded_by")})
        conn.close()
        return {"tables": result}
    except Exception as e:
        return {"error": str(e)}


@app.get("/db/table/{table_name}")
def get_table_data(table_name: str):
    if VALID_TABLES and table_name not in VALID_TABLES:
        return {"error": "Invalid table name"}
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute(f"SELECT * FROM [{table_name}]")
        rows = cursor.fetchall()
        columns = [desc[0] for desc in cursor.description]
        cursor.execute("SELECT updated_at, uploaded_by FROM table_update_log WHERE table_name = ?", (table_name,))
        log_row = cursor.fetchone()
        conn.close()
        return {"table": table_name, "columns": columns,
            "rows": [list(r) for r in rows], "total": len(rows),
            "updated_at": log_row[0] if log_row else None,
            "uploaded_by": log_row[1] if log_row else None}
    except Exception as e:
        return {"error": str(e)}