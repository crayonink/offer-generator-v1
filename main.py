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
    Ti: float
    Tf: float
    refractory_weight: float
    fuel_cv: float = 8500.0
    time_taken_hr: float
    refractory_heat_factor: float = 0.25
    efficiency: float = 0.52
    ladle_tons: float = 10.0


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
        cascade_counts = {}

        for table, name_col in PARTS_TABLES:
            rows = conn.execute(
                f"SELECT rowid, {name_col}, qty, rate FROM {table} WHERE rate IS NOT NULL"
            ).fetchall()
            updated = 0
            for rowid, part_name, qty, rate in rows:
                # Match: old rate must match AND name must fuzzy-match
                if old_price is not None and abs(float(rate) - float(old_price)) > 0.01:
                    continue
                if _norm(part_name or "") != norm_item:
                    continue
                new_amount = round(float(qty or 0) * req.price, 2) if qty is not None else None
                conn.execute(
                    f"UPDATE {table} SET rate=?, amount=? WHERE rowid=?",
                    (req.price, new_amount, rowid)
                )
                updated += 1
            if updated:
                cascade_counts[table] = updated

        conn.commit()

        # ── Excel formula cascade (for tables not covered above) ────────────
        xl_path = _find_latest_pricebook()
        xl_results = {}
        if xl_path:
            xl_results = _cascade_recalculate(xl_path, conn)
            conn.commit()

        conn.close()
        return {
            "success": True,
            "cascaded": bool(xl_path),
            "direct_cascade": cascade_counts,
            "results": {k: v for k, v in xl_results.items() if "rows" in v},
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
        from calculations.pipes import PipeInputs, calculate_pipe_sizes
        from bom.selectors.selection_engine import select_equipment
        from bom.vlph_builder import build_vlph_120t_df

        burner_inputs = BurnerInputs(
            Ti=req.Ti,
            Tf=req.Tf,
            refractory_weight=req.refractory_weight,
            fuel_cv=req.fuel_cv,
            time_taken_hr=req.time_taken_hr,
            refractory_heat_factor=req.refractory_heat_factor,
            efficiency=req.efficiency,
        )
        br = calculate_burner(burner_inputs)

        pipe_results = calculate_pipe_sizes(PipeInputs(
            ng_flow_nm3hr=br.extra_firing_rate_nm3hr,
            air_flow_nm3hr=br.air_qty_nm3hr,
        ))

        equipment = select_equipment(
            ng_flow_nm3hr=br.extra_firing_rate_nm3hr,
            air_flow_nm3hr=br.air_qty_nm3hr,
        )

        bom_df = build_vlph_120t_df(equipment, ladle_tons=req.ladle_tons)

        # Split summary rows from detail rows
        detail = bom_df[bom_df["MEDIA"] != ""].copy()
        system_total     = float(bom_df.loc[bom_df["ITEM NAME"] == "SYSTEM ITEMS TOTAL", "TOTAL"].values[0])
        bought_out_total = float(bom_df.loc[bom_df["ITEM NAME"] == "BOUGHT OUT ITEMS",   "TOTAL"].values[0])
        encon_total      = float(bom_df.loc[bom_df["ITEM NAME"] == "ENCON ITEMS",        "TOTAL"].values[0])
        grand_total      = float(bom_df.loc[bom_df["ITEM NAME"] == "GRAND TOTAL",        "TOTAL"].values[0])

        return {
            "calculations": {
                "Ti": req.Ti,
                "Tf": req.Tf,
                "refractory_weight": req.refractory_weight,
                "fuel_cv": req.fuel_cv,
                "time_taken_hr": req.time_taken_hr,
                "avg_temp_rise":                  round(br.avg_temp_rise, 2),
                "firing_rate_kcal":               round(br.firing_rate_kcal, 2),
                "heat_load_kcal":                 round(br.heat_load_kcal, 2),
                "fuel_consumption_nm3":           round(br.fuel_consumption_nm3, 2),
                "calculated_firing_rate_nm3hr":   round(br.calculated_firing_rate_nm3hr, 2),
                "extra_firing_rate_nm3hr":        round(br.extra_firing_rate_nm3hr, 2),
                "final_firing_rate_mw":           round(br.final_firing_rate_mw, 2),
                "air_qty_nm3hr":                  round(br.air_qty_nm3hr, 2),
                "cfm":                            round(br.cfm, 2),
                "blower_hp_calc":                 round(br.blower_hp, 2),
            },
            "pipes": {
                "ng_flow":      round(br.extra_firing_rate_nm3hr, 2),
                "ng_velocity":  25.0,
                "ng_dia_mm":    round(pipe_results.ng_pipe_inner_dia_mm, 2),
                "ng_nb":        pipe_results.ng_pipe_nb,
                "air_flow":     round(br.air_qty_nm3hr, 2),
                "air_velocity": 15.0,
                "air_dia_mm":   round(pipe_results.air_pipe_inner_dia_mm, 2),
                "air_nb":       pipe_results.air_pipe_nb,
            },
            "equipment": {
                "burner_model":   equipment["burner"]["model"],
                "blower_model":   equipment["blower"]["model"],
                "blower_hp":      equipment["blower"]["hp"],
                "blower_airflow": equipment["blower"]["airflow_nm3hr"],
                "ng_gas_train":   f'{equipment["ng_gas_train"]["inlet_nb"]} x {equipment["ng_gas_train"]["outlet_nb"]} NB',
                "agr_nb":         equipment["agr"]["nb"],
            },
            "bom": detail[["MEDIA","ITEM NAME","REFERENCE","QTY","UNIT PRICE","TOTAL"]].to_dict(orient="records"),
            "cost_summary": {
                "system_total":     round(system_total, 2),
                "bought_out_total": round(bought_out_total, 2),
                "encon_total":      round(encon_total, 2),
                "grand_total":      round(grand_total, 2),
            },
        }
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
        from calculations.pipes import PipeInputs, calculate_pipe_sizes
        from bom.selectors.selection_engine import select_equipment
        from bom.hlph_builder import build_hlph_df

        burner_inputs = BurnerInputs(
            Ti=req.Ti,
            Tf=req.Tf,
            refractory_weight=req.refractory_weight,
            fuel_cv=req.fuel_cv,
            time_taken_hr=req.time_taken_hr,
            refractory_heat_factor=req.refractory_heat_factor,
            efficiency=req.efficiency,
        )
        br = calculate_burner(burner_inputs)

        pipe_results = calculate_pipe_sizes(PipeInputs(
            ng_flow_nm3hr=br.extra_firing_rate_nm3hr,
            air_flow_nm3hr=br.air_qty_nm3hr,
        ))

        equipment = select_equipment(
            ng_flow_nm3hr=br.extra_firing_rate_nm3hr,
            air_flow_nm3hr=br.air_qty_nm3hr,
        )

        bom_df = build_hlph_df(equipment, ladle_tons=req.ladle_tons)

        detail = bom_df[bom_df["MEDIA"] != ""].copy()
        system_total     = float(bom_df.loc[bom_df["ITEM NAME"] == "SYSTEM ITEMS TOTAL", "TOTAL"].values[0])
        bought_out_total = float(bom_df.loc[bom_df["ITEM NAME"] == "BOUGHT OUT ITEMS",   "TOTAL"].values[0])
        encon_total      = float(bom_df.loc[bom_df["ITEM NAME"] == "ENCON ITEMS",        "TOTAL"].values[0])
        grand_total      = float(bom_df.loc[bom_df["ITEM NAME"] == "GRAND TOTAL",        "TOTAL"].values[0])

        return {
            "calculations": {
                "Ti": req.Ti,
                "Tf": req.Tf,
                "refractory_weight": req.refractory_weight,
                "fuel_cv": req.fuel_cv,
                "time_taken_hr": req.time_taken_hr,
                "avg_temp_rise":                  round(br.avg_temp_rise, 2),
                "firing_rate_kcal":               round(br.firing_rate_kcal, 2),
                "heat_load_kcal":                 round(br.heat_load_kcal, 2),
                "fuel_consumption_nm3":           round(br.fuel_consumption_nm3, 2),
                "calculated_firing_rate_nm3hr":   round(br.calculated_firing_rate_nm3hr, 2),
                "extra_firing_rate_nm3hr":        round(br.extra_firing_rate_nm3hr, 2),
                "final_firing_rate_mw":           round(br.final_firing_rate_mw, 2),
                "air_qty_nm3hr":                  round(br.air_qty_nm3hr, 2),
                "cfm":                            round(br.cfm, 2),
                "blower_hp_calc":                 round(br.blower_hp, 2),
            },
            "pipes": {
                "ng_flow":      round(br.extra_firing_rate_nm3hr, 2),
                "ng_velocity":  25.0,
                "ng_dia_mm":    round(pipe_results.ng_pipe_inner_dia_mm, 2),
                "ng_nb":        pipe_results.ng_pipe_nb,
                "air_flow":     round(br.air_qty_nm3hr, 2),
                "air_velocity": 15.0,
                "air_dia_mm":   round(pipe_results.air_pipe_inner_dia_mm, 2),
                "air_nb":       pipe_results.air_pipe_nb,
            },
            "equipment": {
                "burner_model":   equipment["burner"]["model"],
                "blower_model":   equipment["blower"]["model"],
                "blower_hp":      equipment["blower"]["hp"],
                "blower_airflow": equipment["blower"]["airflow_nm3hr"],
                "ng_gas_train":   f'{equipment["ng_gas_train"]["inlet_nb"]} x {equipment["ng_gas_train"]["outlet_nb"]} NB',
                "agr_nb":         equipment["agr"]["nb"],
            },
            "bom": detail[["MEDIA","ITEM NAME","REFERENCE","QTY","UNIT PRICE","TOTAL"]].to_dict(orient="records"),
            "cost_summary": {
                "system_total":     round(system_total, 2),
                "bought_out_total": round(bought_out_total, 2),
                "encon_total":      round(encon_total, 2),
                "grand_total":      round(grand_total, 2),
            },
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

    def hdr(ws, row, col, val, bg=NAVY, fg=WHITE, bold=True, size=11):
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
    if req.equipment_type == "Regen":
        params = [
            ("Material Weight",   f"{calc.get('material_weight_kg','')} kg"),
            ("Initial Temp (Ti)", f"{calc.get('Ti','')} °C"),
            ("Final Temp (Tf)",   f"{calc.get('Tf','')} °C"),
            ("Temp Rise (ΔT)",    f"{calc.get('delta_T','')} °C"),
            ("Specific Heat Cp",  f"{calc.get('Cp','')} kJ/kg·°C"),
            ("Cycle Time",        f"{calc.get('cycle_time_hr','')} hr"),
            ("Efficiency",        f"{round(calc.get('efficiency',0)*100)}%"),
            ("Heat Required",     f"{calc.get('heat_required_kj','')} kJ"),
            ("Required Power",    f"{calc.get('required_kw','')} kW"),
            ("No. of Pairs",      f"{calc.get('num_pairs','')} × 1000 KW"),
        ]
    else:
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

    if req.equipment_type == "Regen":
        bom_cols = ["SECTION", "ITEM NAME", "SPECIFICATION", "QTY", "COST/UNIT", "TOTAL COST", "SELL/UNIT", "TOTAL SELLING"]
    else:
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

    if req.equipment_type == "Regen":
        summary_rows = [
            ("Total Cost",    cs.get("total_cost", 0)),
            ("Total Selling", cs.get("total_selling", 0)),
            ("Markup",        cs.get("markup", 0)),
        ]
    else:
        summary_rows = [
            ("Bought Out Total", cs.get("bought_out_total", 0)),
            ("ENCON Total",      cs.get("encon_total", 0)),
            ("Grand Total",      cs.get("grand_total", 0)),
        ]

    for label, val in summary_rows:
        is_total = "Total" in label and label not in ("Bought Out Total", "ENCON Total")
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

    # ── Regen supplementary sheets ───────────────────────────────────────────
    if req.equipment_type == "Regen" and req.equipment:
        supp = req.equipment.get("supplementary", {})
        if supp:
            # ── Sheet: Burner Sizing & Costing ────────────────────────────
            ws2 = wb.create_sheet("Burner Sizing")
            ws2.column_dimensions["A"].width = 22
            ws2.column_dimensions["B"].width = 20
            ws2.column_dimensions["C"].width = 18
            ws2.column_dimensions["D"].width = 18
            ws2.column_dimensions["E"].width = 18
            r2 = 1
            ws2.merge_cells(f"A{r2}:E{r2}")
            t2 = ws2.cell(row=r2, column=1, value=f"Burner + Regenerator Sizing & Costing — {supp.get('burner_sizing',{}).get('kw','')} KW")
            t2.font = Font(bold=True, color=WHITE, size=13, name="Calibri")
            t2.fill = PatternFill("solid", fgColor=NAVY)
            t2.alignment = Alignment(horizontal="center", vertical="center")
            ws2.row_dimensions[r2].height = 28
            r2 += 2

            bs = supp.get("burner_sizing", {})
            rates = bs.get("material_rates", {})
            ws2.merge_cells(f"A{r2}:E{r2}")
            hdr(ws2, r2, 1, "MATERIAL RATES (with 10% wastage)", size=9)
            r2 += 1
            for col, label in [(1,"Material"),(2,"Mat. ₹/kg"),(3,"Labour ₹/kg"),(4,"Total ₹/kg")]:
                hdr(ws2, r2, col, label, size=9)
            r2 += 1
            for mat, mat_r, lab_r, tot in [("MS",50,25,rates.get("ms",82.5)),("SS",50,25,rates.get("ss",82.5)),("Refractory",56,25,rates.get("refractory",89.1)),("Ceramic Balls",125,0,rates.get("ceramic_balls",137.5))]:
                for ci, v in enumerate([mat, mat_r, lab_r, tot], 1):
                    cell(ws2, r2, ci, v, bg=GREY if r2%2==0 else WHITE, align="right" if ci>1 else "left")
                r2 += 1
            r2 += 1

            ws2.merge_cells(f"A{r2}:E{r2}")
            hdr(ws2, r2, 1, "COST BREAKDOWN (per unit)", size=9)
            r2 += 1
            for col, label in [(1,"Component"),(2,"Material"),(3,"Weight (kg)"),(4,"Rate ₹/kg"),(5,"Cost ₹")]:
                hdr(ws2, r2, col, label, size=9)
            r2 += 1
            for d in bs.get("cost_detail", []):
                for ci, v in enumerate([d["component"], d["material"], d["weight_kg"], d["rate"], d["cost"]], 1):
                    cell(ws2, r2, ci, v, bg=GREY if r2%2==0 else WHITE, align="right" if ci>2 else "left",
                         num_fmt='#,##0.00' if ci>2 else None)
                r2 += 1
            cell(ws2, r2, 4, "Total per unit", bold=True, bg=GREEN_BG, fg=GREEN, align="right")
            cell(ws2, r2, 5, bs.get("total_unit_cost",0), bold=True, bg=GREEN_BG, fg=GREEN, align="right", num_fmt='#,##0.00')
            r2 += 1
            cell(ws2, r2, 4, "Total per pair (×2)", bold=True, bg=GREEN_BG, fg=GREEN, align="right")
            cell(ws2, r2, 5, bs.get("total_pair_cost",0), bold=True, bg=GREEN_BG, fg=GREEN, align="right", num_fmt='#,##0.00')

            # ── Sheet: Pipe Sizes ─────────────────────────────────────────
            ws3 = wb.create_sheet("Pipe Sizes")
            ws3.column_dimensions["A"].width = 28
            ws3.column_dimensions["B"].width = 22
            r3 = 1
            ws3.merge_cells(f"A{r3}:B{r3}")
            t3 = ws3.cell(row=r3, column=1, value="Line Sizes — Natural Gas (NG) @ 0.05 barg")
            t3.font = Font(bold=True, color=WHITE, size=13, name="Calibri")
            t3.fill = PatternFill("solid", fgColor=NAVY)
            t3.alignment = Alignment(horizontal="center", vertical="center")
            ws3.row_dimensions[r3].height = 28
            r3 += 2
            ps = supp.get("pipe_sizes", {})
            pipe_rows = [
                ("Burner KW", f"{ps.get('kw','')} KW"),
                ("NG Flow Rate", f"{ps.get('ng_flow_nm3hr','')} Nm³/hr"),
                ("Air Flow Rate", f"{ps.get('air_flow_nm3hr','')} Nm³/hr"),
                ("Total Flue Gas", f"{ps.get('flue_flow_nm3hr','')} Nm³/hr"),
                ("Air Line DN", f"DN {ps.get('air_line_dn','')}"),
                ("Gas Line DN", f"DN {ps.get('gas_line_dn','')}"),
                ("Flue Gas Line DN", f"DN {ps.get('flue_line_dn','')}"),
            ]
            hdr(ws3, r3, 1, "Parameter", size=9)
            hdr(ws3, r3, 2, "Value", size=9)
            r3 += 1
            for i, (lbl, val) in enumerate(pipe_rows):
                bg = LIGHT if lbl.endswith("DN") else (GREY if i%2==0 else WHITE)
                cell(ws3, r3, 1, lbl, bold=lbl.endswith("DN"), bg=bg)
                cell(ws3, r3, 2, val, bold=lbl.endswith("DN"), bg=bg)
                r3 += 1

            # ── Sheet: Blower Selection ───────────────────────────────────
            ws4 = wb.create_sheet("Blower Selection")
            ws4.column_dimensions["A"].width = 20
            for col in ["B","C","D","E","F"]:
                ws4.column_dimensions[col].width = 18
            r4 = 1
            ws4.merge_cells(f"A{r4}:F{r4}")
            t4 = ws4.cell(row=r4, column=1, value="ENCON 40\" WG Blower Selection")
            t4.font = Font(bold=True, color=WHITE, size=13, name="Calibri")
            t4.fill = PatternFill("solid", fgColor=NAVY)
            t4.alignment = Alignment(horizontal="center", vertical="center")
            ws4.row_dimensions[r4].height = 28
            r4 += 2
            bl = supp.get("blower_selection", {})
            cat = supp.get("blower_catalogue", [])
            # Selected summary
            sel_rows = [
                ("Burner KW", f"{bl.get('kw','')} KW"),
                ("Selected Model", bl.get("selected_model","")),
                ("HP", bl.get("hp","")),
                ("Flow Rate", f"{bl.get('nm3hr','')} Nm³/hr"),
                ("Qty per pair", str(bl.get("qty_per_pair",""))),
                ("Costing price used", bl.get("costing_price",0)),
            ]
            ws4.merge_cells(f"A{r4}:F{r4}")
            hdr(ws4, r4, 1, "SELECTED BLOWER", size=9)
            r4 += 1
            for i, (lbl, val) in enumerate(sel_rows):
                bg = LIGHT if lbl == "Costing price used" else (GREY if i%2==0 else WHITE)
                cell(ws4, r4, 1, lbl, bold=lbl=="Costing price used", bg=bg)
                ws4.merge_cells(f"B{r4}:F{r4}")
                cell(ws4, r4, 2, val, bold=lbl=="Costing price used", bg=bg,
                     num_fmt='#,##0.00' if isinstance(val, (int,float)) else None)
                r4 += 1
            r4 += 1
            # Full catalogue
            ws4.merge_cells(f"A{r4}:F{r4}")
            hdr(ws4, r4, 1, "FULL CATALOGUE", size=9)
            r4 += 1
            for ci, lbl in enumerate(["Model","HP","CFM","Nm³/hr","Price w/o Motor ₹","Price with Motor ₹"], 1):
                hdr(ws4, r4, ci, lbl, size=9)
            r4 += 1
            for row_b in cat:
                is_sel = row_b["model"] == bl.get("selected_model","")
                bg = LIGHT if is_sel else (GREY if r4%2==0 else WHITE)
                for ci, v in enumerate([row_b["model"], row_b["hp"], row_b["cfm"], row_b["nm3hr"], row_b["price_without_motor"], row_b["price_with_motor"]], 1):
                    cell(ws4, r4, ci, v, bold=is_sel, bg=bg, align="right" if ci>2 else "left",
                         num_fmt='#,##0.00' if ci>4 else None)
                r4 += 1

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