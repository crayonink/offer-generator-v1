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

# Swagger/ReDoc moved off /docs so the Antora handbook can own /docs.
app = FastAPI(docs_url="/api-docs", redoc_url="/api-redoc")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "vlph.db")

# Documentation site (Antora, pre-built into ./docs_site) served at /docs.
# Guarded so a missing build never breaks app start-up.
from fastapi.staticfiles import StaticFiles
_DOCS_DIR = os.path.join(BASE_DIR, "docs_site")
if os.path.isdir(_DOCS_DIR):
    app.mount("/docs", StaticFiles(directory=_DOCS_DIR, html=True), name="docs")
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


def peek_quote_seq() -> str:
    """Return the sequence the next generate-quote call WILL receive,
    without consuming it. Used by the form to preview the auto ref."""
    if os.path.exists(COUNTER_FILE):
        with open(COUNTER_FILE) as f:
            return str(int(f.read().strip()) + 1).zfill(3)
    return "001"


def _with_salutation(salutation: str, name: str) -> str:
    """Return 'Mr. Ashish Gupta' if both supplied; else whichever is present."""
    s = (salutation or "").strip()
    n = (name or "").strip()
    if s and n:
        return f"{s} {n}"
    return n or s


def _person_initials(name: str) -> str:
    """First letter of each whitespace-separated word, uppercased.
    'Jyotirmoy Rabha' -> 'JR'. Returns '' for empty input."""
    parts = [p for p in (name or "").strip().split() if p]
    return "".join(p[0].upper() for p in parts)


_LOCATION_CODES = {
    "goa":       "GOA",
    "vadodara":  "VDD",
    "faridabad": "FBD",
}


def _location_code(location: str) -> str:
    return _LOCATION_CODES.get((location or "").strip().lower(), "")


def build_enquiry_ref(seq: str, technical_person: str,
                      location: str = "",
                      year: Optional[int] = None) -> str:
    """ENCON.04026.{seq}/{LOC}/{initials} DT.{DD/MM/YYYY}.

    The '04026' segment is ENCON's fixed code (applies to ladle and
    tundish offers alike). Location code (GOA/VDD/FBD) and the
    technical-person initials get omitted only when missing.
    """
    from datetime import datetime as _dt
    today = _dt.now()
    ini = _person_initials(technical_person)
    loc = _location_code(location)
    # Assemble the slash segment: include only the parts we have.
    parts = [seq]
    if loc: parts.append(loc)
    if ini: parts.append(ini)
    body = "/".join(parts)
    date_str = today.strftime("%d/%m/%Y")
    return f"ENCON.04026.{body} DT.{date_str}"

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
    'rotary_joint_master',
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


# ── Google Drive OAuth ────────────────────────────────────────────────────
# One-time browser sign-in by an encon.in user (process@encon.in is the
# intended account). The captured refresh token is stored in vlph.db and
# used by drive_uploader for every subsequent offer upload.

@app.get("/auth/drive/login")
def drive_oauth_login():
    """Redirect the user to Google's OAuth consent screen."""
    from fastapi.responses import RedirectResponse
    client_id = os.environ.get("GOOGLE_OAUTH_CLIENT_ID", "").strip()
    redirect_uri = os.environ.get("GOOGLE_OAUTH_REDIRECT_URI", "").strip()
    if not client_id or not redirect_uri:
        return HTMLResponse(
            "<h3>Drive OAuth not configured</h3>"
            "<p>Set GOOGLE_OAUTH_CLIENT_ID and GOOGLE_OAUTH_REDIRECT_URI in Railway env vars.</p>",
            status_code=500,
        )
    from urllib.parse import urlencode
    params = {
        "client_id":     client_id,
        "redirect_uri":  redirect_uri,
        "response_type": "code",
        "scope":         "https://www.googleapis.com/auth/drive.file",
        "access_type":   "offline",   # required to get a refresh token
        "prompt":        "consent",   # forces refresh-token reissue every time
        "include_granted_scopes": "true",
    }
    return RedirectResponse(
        f"https://accounts.google.com/o/oauth2/v2/auth?{urlencode(params)}"
    )


@app.get("/auth/drive/callback")
def drive_oauth_callback(code: str = "", error: str = ""):
    """Handle Google's redirect: exchange the authorisation code for a
    refresh token and persist it. Then show a small confirmation page."""
    if error:
        return HTMLResponse(
            f"<h3>Drive auth failed</h3><p>{error}</p>", status_code=400
        )
    if not code:
        return HTMLResponse(
            "<h3>Drive auth failed</h3><p>No code returned by Google.</p>",
            status_code=400,
        )
    client_id     = os.environ.get("GOOGLE_OAUTH_CLIENT_ID", "").strip()
    client_secret = os.environ.get("GOOGLE_OAUTH_CLIENT_SECRET", "").strip()
    redirect_uri  = os.environ.get("GOOGLE_OAUTH_REDIRECT_URI", "").strip()
    if not (client_id and client_secret and redirect_uri):
        return HTMLResponse(
            "<h3>Drive OAuth not configured</h3>"
            "<p>Missing CLIENT_ID / CLIENT_SECRET / REDIRECT_URI env vars.</p>",
            status_code=500,
        )
    import urllib.request, urllib.parse, json as _json
    body = urllib.parse.urlencode({
        "code":          code,
        "client_id":     client_id,
        "client_secret": client_secret,
        "redirect_uri":  redirect_uri,
        "grant_type":    "authorization_code",
    }).encode()
    try:
        with urllib.request.urlopen(
            "https://oauth2.googleapis.com/token", data=body, timeout=15
        ) as resp:
            token_data = _json.loads(resp.read().decode())
    except Exception as e:
        return HTMLResponse(
            f"<h3>Token exchange failed</h3><pre>{e}</pre>", status_code=500
        )
    refresh_token = token_data.get("refresh_token")
    if not refresh_token:
        return HTMLResponse(
            "<h3>No refresh token returned</h3>"
            "<p>Try /auth/drive/login again. Make sure 'prompt=consent' is set "
            "(it should be — but if Google has already issued a refresh token "
            "to this client and account, it skips it). You may need to revoke "
            "access at https://myaccount.google.com/permissions and retry.</p>"
            f"<pre>{_json.dumps(token_data, indent=2)}</pre>",
            status_code=500,
        )
    from engine.drive_uploader import save_refresh_token
    save_refresh_token(refresh_token)
    # Show the refresh token so the user can persist it as a Railway env
    # var (vlph.db on Railway gets wiped on every redeploy — env vars
    # survive). Once GOOGLE_DRIVE_REFRESH_TOKEN is set the uploader reads
    # from the env var first and never has to re-auth.
    return HTMLResponse(
        "<div style='font-family:Calibri,Arial,sans-serif;max-width:780px;margin:24px auto;padding:24px;'>"
        "<h3 style='color:green'>Google Drive connected ✓</h3>"
        "<p>The offer generator will now upload every new offer (docx + pdf): "
        "VLPH/HLPH go to the Ladle folder, Tundish to the Tundish folder.</p>"
        "<div style='background:#fff7ed;border:1px solid #fdba74;border-radius:8px;padding:14px;margin:18px 0;'>"
        "<b>Important — make it permanent:</b><br/>"
        "On Railway redeploys the local token gets wiped. To keep auth alive "
        "across deploys, copy the refresh token below and set it as a Railway "
        "env var named <code>GOOGLE_DRIVE_REFRESH_TOKEN</code>.<br/>"
        f"<textarea readonly style='width:100%;height:60px;margin-top:10px;font-family:monospace;font-size:12px;'>{refresh_token}</textarea>"
        "<p style='font-size:0.85rem;color:#7c2d12;margin-top:8px;'>Treat this like a password. Don't paste it in chat or commit it to git.</p>"
        "</div>"
        "<p>You can close this tab once the env var is saved.</p>"
        "</div>"
    )


@app.get("/auth/drive/status")
def drive_oauth_status():
    """JSON endpoint the dashboard can poll to show 'Connected' vs not."""
    from engine.drive_uploader import is_authorized
    return {"authorized": is_authorized()}


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
    price: Optional[float] = None

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
    direct_burner_capacity: float = 0.0         # kW (direct mode — per burner for tundish)
    direct_heat_input_kcal_hr: float = 0.0      # kcal/hr (direct mode — by heat; takes precedence over capacity)
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
    num_burners: int = 1                         # Number of burners (tundish: splits firing rate, multiplies burner-line items)
    ms_structure_kg_override: float = 0.0        # Tundish: override MS structure weight used for fabrication cost
    ceramic_rolls_override: int = 0              # Tundish: override ceramic-fibre roll count
    hood_type: str = "up_down"                   # "up_down" | "swivel_manual" | "swivel_geared"
    special_auto_ignition: bool = False            # Special Requirements: auto-ignition requested
    special_auto_controls: bool = False            # Special Requirements: auto-controls requested


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
    subject: Optional[str] = ""
    ref_no: Optional[str] = ""
    your_ref: Optional[str] = ""
    enquiry_ref: Optional[str] = ""
    marketing_person: Optional[str] = ""
    marketing_phone: Optional[str] = ""
    marketing_email: Optional[str] = ""
    technical_person: Optional[str] = ""
    technical_phone: Optional[str] = ""
    technical_email: Optional[str] = ""
    # Technical data (passed from lastCalc to populate template tech table)
    ladle_tons: Optional[float] = 0
    ladle_dim: Optional[str] = ""
    ladle_drawing_no: Optional[str] = ""
    refractory_weight_kg: Optional[str] = ""
    heating_schedule: Optional[str] = ""
    fuel_cv: Optional[str] = ""
    fuel_consumption: Optional[str] = ""
    burner_model: Optional[str] = ""
    blower_model: Optional[str] = ""
    blower_size: Optional[str] = ""
    blower_capacity: Optional[str] = ""
    hydraulic_motor_hp: Optional[str] = ""
    max_electrical_load: Optional[str] = ""
    total_in_words: Optional[str] = ""
    heating_time: Optional[str] = ""
    fuel_name: Optional[str] = ""
    burner_capacity_range: Optional[str] = ""
    pumping_unit: Optional[str] = ""
    hood_movement: Optional[str] = ""
    hood_type: Optional[str] = "up_down"   # 'swivel' or 'up_down'
    pilot_gas_type: Optional[str] = "LPG"  # 'LPG' or 'NG'
    ignition_method: Optional[str] = ""
    # Tundish dual-fuel fields
    num_burners: Optional[str] = ""
    fuel2_cv: Optional[str] = ""
    fuel2_consumption: Optional[str] = ""
    max_fuel_consumption1: Optional[str] = ""
    max_fuel_consumption2: Optional[str] = ""
    is_oil: Optional[bool] = False   # True for oil fuels — drives scope-of-supply rendering
    is_dual: Optional[bool] = False  # True for dual-fuel (gas + oil)
    control_mode: Optional[str] = "automatic"  # "manual" or "automatic" — drives scope text
    auto_control_type: Optional[str] = "plc"   # "plc" | "plc_agr" | "pid" — drives PDF scope wording
    control_valve_type: Optional[str] = "pneumatic"  # "pneumatic" | "motorised" — wording in I&C / Temp Control
    special_auto_ignition: Optional[bool] = False  # Special Requirements: drives pilot-burner sections
    special_auto_controls:  Optional[bool] = False  # Special Requirements
    vertical_qty:   Optional[int] = 1   # Annexure I scope-of-supply header — Vertical units count
    horizontal_qty: Optional[int] = 1   # Annexure I scope-of-supply header — Horizontal units count
    purging_line:    Optional[str] = "no"   # "yes" | "no" — drives the Nitrogen Purging block in the offer
    hpu_variant:     Optional[str] = "Duplex 1"   # "Simplex" | "Duplex 1" | "Duplex 2" — Pumping Unit type
    burner_kw_value: Optional[str] = ""   # Pre-formatted total kW for the ENCON burner body line
    location:        Optional[str] = ""   # "Goa" | "Vadodara" | "Faridabad" — appended to enquiry ref
    poc_salutation:        Optional[str] = ""   # "Mr." | "Mrs." | "Miss" | "Dr." — prefixed to poc_name in offer
    marketing_salutation:  Optional[str] = ""
    technical_salutation:  Optional[str] = ""
    # Editable Annexure IV — Terms & Conditions
    tnc_prices:               Optional[str] = "EX Bhagola (Ex-Works)"
    tnc_delivery:             Optional[str] = "10 – 12 weeks from our ENCON works, from the date of receipt of advance or drawing approval, whichever is later."
    tnc_gst:                  Optional[str] = "18% extra."
    tnc_hsn_code:             Optional[str] = "84541000"
    tnc_pan_gst:              Optional[str] = "PAN: AAACE0327M  |  GST: 06AAACE0327M1ZV"
    tnc_payment_terms:        Optional[str] = "30% advance with the purchase order\n70% against proforma invoice prior to dispatch"
    tnc_packing_forwarding:   Optional[str] = "4% and 2% respectively."
    tnc_freight:              Optional[str] = "In client's scope."
    tnc_transit_insurance:    Optional[str] = "To be arranged by the client."
    tnc_validity:             Optional[str] = "45 days from the date of our offer."
    tnc_inspection:           Optional[str] = "If required, materials can be inspected at our works before dispatch, at the client's cost and with prior intimation."
    tnc_guarantee:            Optional[str] = "Materials are guaranteed for 18 months from the date of dispatch, or 12 months from commissioning, whichever is earlier — against any manufacturing defects or poor workmanship."
    bom_items: Optional[List[dict]] = []   # [{item, make, media, ref}, ...] for offer scope + MAKE LIST
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

@app.get("/tundish", response_class=HTMLResponse)
def tundish_costing_form():
    html_path = os.path.join(BASE_DIR, "tundish_costing.html")
    with open(html_path, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())

@app.get("/equipment-offer", response_class=HTMLResponse)
def equipment_offer_hub():
    """Hub page that lets the user pick which stand-alone equipment
    offer to generate (HPU, Blower or Burner)."""
    html_path = os.path.join(BASE_DIR, "equipment_offer.html")
    with open(html_path, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())


@app.get("/hpu", response_class=HTMLResponse)
def hpu_picker():
    """Sub-hub under Equipment Offer that lets the user pick between
    HPU (Heating + Pumping Unit) and PU (Pumping Unit only)."""
    html_path = os.path.join(BASE_DIR, "hpu_picker.html")
    with open(html_path, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())


@app.get("/hpu/heating-pumping", response_class=HTMLResponse)
def hpu_costing_form():
    """HPU offer form — pumping skid with in-built electric heater
    (model codes HPS / HPD / HPDD)."""
    html_path = os.path.join(BASE_DIR, "hpu_costing.html")
    with open(html_path, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())


@app.get("/hpu/pumping-only", response_class=HTMLResponse)
def pu_costing_form():
    """PU offer form — pumping skid without heater, for pre-heated
    oils such as LDO / LSHS (model codes PUS / PUD / PUDD)."""
    html_path = os.path.join(BASE_DIR, "pu_costing.html")
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

        # ── Rotary Joint ──────────────────────────────────────────────────
        rj_cols = ["rowid"] + [d[0] for d in c.execute("SELECT * FROM rotary_joint_master LIMIT 0").description]
        rotary_joint = [
            dict(zip(rj_cols, r))
            for r in q("SELECT rowid, * FROM rotary_joint_master ORDER BY nb")
        ]

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
            "rotary_joint": rotary_joint,
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
        # FUEL DENSITY rows live in this table only because burner sizing
        # uses _get_fuel_density() to look them up by name. They are not
        # editable rates, so exclude them from the Rates view.
        rows = conn.execute(
            "SELECT rowid, item, category, price, previous_price, updated_at, company, specification "
            "FROM component_price_master "
            "WHERE item NOT LIKE 'FUEL DENSITY%' "
            "  AND (category IS NULL OR category != 'Fuel Density') "
            "ORDER BY category, item"
        ).fetchall()
        conn.close()
        return [{"rowid": r[0], "item": r[1], "category": r[2],
                 "price": r[3], "previous_price": r[4], "updated_at": r[5],
                 "company": r[6], "specification": r[7]} for r in rows]
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
        cols = {r[1] for r in conn.execute(f"PRAGMA table_info({req.table})").fetchall()}
        if req.qty is not None and 'qty' in cols:
            conn.execute(f"UPDATE {req.table} SET qty=? WHERE rowid=?", (req.qty, req.rowid))
        if req.rate is not None and 'rate' in cols:
            conn.execute(f"UPDATE {req.table} SET rate=? WHERE rowid=?", (req.rate, req.rowid))
        if req.price is not None and 'price' in cols:
            conn.execute(f"UPDATE {req.table} SET price=? WHERE rowid=?", (req.price, req.rowid))
        # Recalculate amount if qty/rate/amount all exist
        if {'qty', 'rate', 'amount'} <= cols:
            row = conn.execute(f"SELECT qty, rate FROM {req.table} WHERE rowid=?", (req.rowid,)).fetchone()
            if row and row[0] is not None and row[1] is not None:
                amount = round(float(row[0]) * float(row[1]), 2)
                conn.execute(f"UPDATE {req.table} SET amount=? WHERE rowid=?", (amount, req.rowid))
        conn.commit()
        select_cols = [c for c in ('qty', 'rate', 'amount', 'price') if c in cols]
        row2 = conn.execute(
            f"SELECT {','.join(select_cols)} FROM {req.table} WHERE rowid=?", (req.rowid,)
        ).fetchone() if select_cols else None
        conn.close()
        result = {"success": True}
        if row2 and select_cols:
            result.update(dict(zip(select_cols, row2)))
        return result
    except Exception as e:
        return {"success": False, "error": str(e)}


# ── Vendor / component price tables + Recup rate constants ──────────────
# Surfaces vendor pricelists (Cair / Dembla / L&T / regulator / orifice /
# hose / gas-train / etc.) and the standalone recup_rates table so they can
# be inspected and edited from /pricelist without raw SQL.

VENDOR_TABLES: dict[str, dict] = {
    # name -> { label, key, numeric_cols, order_by }
    # `order_by` is a SQL fragment (no semicolons) used to sort the rows
    # in ascending order. Falls back to the natural key when not provided.
    "aira_valve_master":            {"label": "AIRA",             "key": ["id"], "numeric_cols": ["list_price", "discount_pct", "net_price"], "order_by": "valve_type, nb"},
    "cair_motorized_valve_master":  {"label": "Cair MOV",         "key": ["id"], "numeric_cols": ["list_price", "discount_pct", "net_price"], "order_by": "valve_type, nb"},
    "dembla_valve_master":          {"label": "Dembla CV",        "key": ["id"], "numeric_cols": ["list_price", "discount_pct", "net_price"], "order_by": "valve_type, nb"},
    "lt_ball_valve_master":         {"label": "L&T Ball Valve",   "key": ["cat_no"], "numeric_cols": ["nb_15","nb_20","nb_25","nb_32","nb_40","nb_50","nb_65","nb_80","nb_100","nb_125","nb_150","nb_200","nb_250","nb_300","nb_350","nb_400","nb_450","nb_500","nb_600"], "order_by": "cat_no"},
    "lt_butterfly_valve_master":    {"label": "L&T Butterfly",    "key": ["id"], "numeric_cols": ["price"], "order_by": "operation, nb"},
    "gas_regulator_master":         {"label": "Gas Regulator",    "key": ["id"], "numeric_cols": ["list_price"], "order_by": "category, nb"},
    "orifice_plate_master":         {"label": "Orifice Plate",    "key": ["nb"], "numeric_cols": ["flanges_price", "plate_price", "fasteners_price", "total_price"], "order_by": "nb"},
    "compensator_master":           {"label": "Compensator",      "key": ["nb"], "numeric_cols": ["price"], "order_by": "nb"},
    "flexible_hose_master":         {"label": "Flexible Hose",    "key": ["id"], "numeric_cols": ["price"], "order_by": "dn, length_mm"},
    "gas_train_master":             {"label": "Gas Train",        "key": ["sr_no"], "numeric_cols": ["price_inr"], "order_by": "sr_no"},
    "motorized_valve_master":       {"label": "Motorized Valve",  "key": ["nb"], "numeric_cols": ["price"], "order_by": "nb"},
    "solenoidvalve_component_master": {"label": "Solenoid Valve", "key": ["id"], "numeric_cols": ["list_price"], "order_by": "section, size, description"},
}


class VendorCellUpdate(BaseModel):
    table: str
    key_values: dict           # { key_col: value, ... }  identifies the row
    column: str                # the numeric column being edited
    new_value: float


class RecupRateUpdate(BaseModel):
    key: str
    value: float


@app.get("/api/pricelist/vendor-tables")
def get_vendor_tables():
    """Return all vendor / component price tables as
    { name: { label, columns, key, numeric_cols, rows } }."""
    try:
        conn = sqlite3.connect(DB_PATH)
        out = {}
        for table, meta in VENDOR_TABLES.items():
            try:
                cols = [r[1] for r in conn.execute(f'PRAGMA table_info("{table}")').fetchall()]
                if not cols:
                    out[table] = {"label": meta["label"], "error": "table missing"}
                    continue
                order_clause = f' ORDER BY {meta["order_by"]}' if meta.get("order_by") else ''
                rows = conn.execute(f'SELECT * FROM "{table}"{order_clause}').fetchall()
                out[table] = {
                    "label":        meta["label"],
                    "columns":      cols,
                    "key":          meta["key"],
                    "numeric_cols": meta["numeric_cols"],
                    "rows":         [dict(zip(cols, r)) for r in rows],
                }
            except Exception as inner:
                out[table] = {"label": meta["label"], "error": str(inner)}
        conn.close()
        return out
    except Exception as e:
        return {"error": str(e)}


@app.put("/api/pricelist/vendor-cell")
def update_vendor_cell(req: VendorCellUpdate):
    """Update one numeric cell in a whitelisted vendor table."""
    if req.table not in VENDOR_TABLES:
        return {"success": False, "error": f"Table '{req.table}' not editable here"}
    meta = VENDOR_TABLES[req.table]
    if req.column not in meta["numeric_cols"]:
        return {"success": False, "error": f"Column '{req.column}' not editable"}
    key_cols = meta["key"]
    if set(req.key_values.keys()) != set(key_cols):
        return {"success": False, "error": f"Expected key cols {key_cols}, got {list(req.key_values.keys())}"}
    try:
        conn = sqlite3.connect(DB_PATH)
        where = " AND ".join(f'"{k}"=?' for k in key_cols)
        params = [req.new_value] + [req.key_values[k] for k in key_cols]
        cur = conn.execute(
            f'UPDATE "{req.table}" SET "{req.column}"=? WHERE {where}', params
        )
        # For Cair / Dembla: when list_price or discount_pct changes, recompute net_price.
        if req.column in ("list_price", "discount_pct") and req.table in ("cair_motorized_valve_master", "dembla_valve_master"):
            row = conn.execute(
                f'SELECT list_price, discount_pct FROM "{req.table}" WHERE {where}',
                [req.key_values[k] for k in key_cols]
            ).fetchone()
            if row and row[0] is not None and row[1] is not None:
                net = round(float(row[0]) * (1 - float(row[1]) / 100.0), 2)
                conn.execute(
                    f'UPDATE "{req.table}" SET net_price=? WHERE {where}',
                    [net] + [req.key_values[k] for k in key_cols]
                )
        conn.commit()
        # Read back the affected row so the client can refresh derived cells.
        all_cols = [r[1] for r in conn.execute(f'PRAGMA table_info("{req.table}")').fetchall()]
        row = conn.execute(
            f'SELECT * FROM "{req.table}" WHERE {where}',
            [req.key_values[k] for k in key_cols]
        ).fetchone()
        conn.close()
        return {
            "success":      True,
            "rows_changed": cur.rowcount,
            "row":          dict(zip(all_cols, row)) if row else None,
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


@app.get("/api/pricelist/recup-rates")
def get_recup_rates():
    """Return every key/value in recup_rates (with unit + notes)."""
    try:
        conn = sqlite3.connect(DB_PATH)
        cols = [r[1] for r in conn.execute('PRAGMA table_info("recup_rates")').fetchall()]
        rows = conn.execute('SELECT * FROM recup_rates ORDER BY key').fetchall()
        conn.close()
        return [dict(zip(cols, r)) for r in rows]
    except Exception as e:
        return {"error": str(e)}


@app.put("/api/pricelist/recup-rate")
def update_recup_rate(req: RecupRateUpdate):
    """Update a single key in recup_rates."""
    try:
        conn = sqlite3.connect(DB_PATH)
        cur = conn.execute(
            "UPDATE recup_rates SET value=? WHERE key=?", (req.value, req.key)
        )
        if cur.rowcount == 0:
            conn.close()
            return {"success": False, "error": f"Key '{req.key}' not found"}
        conn.commit()
        conn.close()
        return {"success": True, "key": req.key, "value": req.value}
    except Exception as e:
        return {"success": False, "error": str(e)}


@app.get("/api/ladle-mapping")
def ladle_mapping(tons: float, type: str = "vertical", hood_type: str = ""):
    """Return auto-filled fabrication/pipeline/ceramic values for a ladle
    capacity. The UI calls this to populate read-only form fields.
    hood_type is used only for vertical ('swivel' or 'up_down'); defaults
    to 'swivel' in the lookup if empty."""
    from bom.vlph_builder import lookup_ladle_fab_pipeline
    return lookup_ladle_fab_pipeline(tons, type, hood_type or None) or {}


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
            # --- Direct mode: user enters burner capacity (kW) or heat input (kcal/hr) ---
            # For tundish: capacity is PER BURNER; total = capacity × num_burners.
            # Heat input takes precedence (already total).
            is_dual = req.fuel2_type != "none" and req.fuel2_cv > 0
            n_burners = max(1, int(req.num_burners or 1))
            if req.direct_heat_input_kcal_hr and req.direct_heat_input_kcal_hr > 0:
                heat_kcal_hr = req.direct_heat_input_kcal_hr
            else:
                # kW → kcal/hr (× 860), then × num_burners for total
                heat_kcal_hr = req.direct_burner_capacity * 860 * n_burners
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

        # --- Fuel 2 calculation (if dual fuel) ---
        is_dual = req.fuel2_type != "none" and req.fuel2_cv > 0

        # Burner pressure derived from blower pressure: 28" -> 24" w.g., 40" -> 36" w.g.
        burner_pressure_wg = 36 if req.blower_pressure == "40" else 24

        # For tundish with multiple burners: blower/gas-train/main-line are sized for
        # TOTAL flow, burner is sized for PER-BURNER flow.
        n_burners = max(1, int(req.num_burners or 1))
        per_burner_ng  = ng_flow / n_burners
        per_burner_air = air_flow / n_burners

        # For dual fuel: pre-compute fuel 2 air flow so blower is sized for
        # whichever fuel needs MORE air (lower CV = higher flow).
        blower_air_flow = air_flow
        if is_dual:
            if req.mode == "direct":
                air_flow2_pre = air_flow  # same heat → same air
            else:
                br2_pre = calculate_burner(BurnerInputs(
                    Ti=req.Ti, Tf=req.Tf,
                    refractory_weight=req.refractory_weight,
                    fuel_cv=req.fuel2_cv,
                    time_taken_hr=req.time_taken_hr,
                    refractory_heat_factor=req.refractory_heat_factor,
                    efficiency=req.efficiency,
                ))
                air_flow2_pre = br2_pre.air_qty_nm3hr
            blower_air_flow = max(air_flow, air_flow2_pre)

        # Pre-compute fuel2 oil LPH for blower CFM sizing in dual-fuel
        f2_oil_lph_for_blower = 0
        if is_dual and req.fuel2_type in OIL_FUELS:
            # heat_kcal_hr exists in direct mode; in calc mode derive from fuel1 flow × CV
            heat_for_f2 = heat_kcal_hr if req.mode == "direct" else (ng_flow * f1_cv)
            f2_flow_kghr = heat_for_f2 / req.fuel2_cv if req.fuel2_cv > 0 else 0
            from bom.selectors.encon_burner import _get_fuel_density
            f2_density = _get_fuel_density(req.fuel2_type)
            f2_oil_lph_for_blower = f2_flow_kghr / f2_density if f2_density else 0

        equip1 = select_equipment(
            ng_flow_nm3hr=ng_flow,
            air_flow_nm3hr=blower_air_flow,
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
            fuel2_lph=f2_oil_lph_for_blower,
        )

        # Size pipes AFTER blower selection — gas pipe for fuel 1's flow,
        # air pipe for the SELECTED blower's RATED airflow (e.g. a 10 HP
        # blower rated at 1360 Nm3/hr — pipe must handle that, not just
        # the application's required flow).
        _rated_blower_air = float(equip1.get("blower", {}).get("airflow_nm3hr") or blower_air_flow)
        pipes1 = calculate_pipe_sizes(PipeInputs(
            ng_flow_nm3hr=ng_flow,
            air_flow_nm3hr=_rated_blower_air,
        ))

        # Tundish multi-burner: re-select burner for per-burner flow
        if n_burners > 1:
            equip_pb = select_equipment(
                ng_flow_nm3hr=per_burner_ng,
                air_flow_nm3hr=per_burner_air,
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
            equip1["burner"] = equip_pb["burner"]

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

        # Auto-fill fabrication (ms_structure) AND ceramic-fibre rolls from
        # the fabrication_ladle_mapping table (nearest ladle capacity for the
        # selected hood type). Pipeline weight is always user-supplied — the
        # file leaves pipeline kg blank on most rows so an auto-fill would
        # silently zero the user's input.
        from bom.vlph_builder import lookup_ladle_fab_pipeline
        _mapped = lookup_ladle_fab_pipeline(req.ladle_tons, "vertical", req.hood_type)
        _pipeline_kg  = req.pipeline_weight_kg
        # User-supplied override wins; fall back to DB-mapped value only when override is 0.
        _ms_override  = req.ms_structure_kg_override or (_mapped.get("fabrication_kg") if _mapped else 0)
        _ceramic_rolls = req.ceramic_rolls_override or (_mapped.get("ceramic_rolls") if _mapped else 0)

        # Air is CV-independent, so use fuel1 for air sizing
        if req.control_mode == "manual":
            bom_df = build_vlph_manual_df(
                equipment=equip1,
                ladle_tons=req.ladle_tons,
                fuel1_type=req.fuel1_type,
                fuel2_type=req.fuel2_type,
                equipment2=equip2,
                purging_line=req.purging_line,
                pressure_gauge_vendor=req.pressure_gauge_vendor,
                pilot_burner=req.pilot_burner,
                pipeline_weight_kg=_pipeline_kg,
                include_pilot=req.manual_pilot_burner == "yes",
                pilot_line_fuel=req.pilot_line_fuel,
                hood_type=req.hood_type,
                ceramic_rolls_override=_ceramic_rolls,
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
                pipeline_weight_kg=_pipeline_kg,
                purging_line=req.purging_line,
                num_burners=n_burners,
                ms_structure_kg_override=_ms_override,
                ceramic_rolls_override=_ceramic_rolls,
                hood_type=req.hood_type,
                special_auto_ignition=req.special_auto_ignition,
            )

        # Split summary rows from detail rows
        detail = bom_df[bom_df["MEDIA"] != ""].copy()
        bought_out_total = float(bom_df.loc[bom_df["ITEM NAME"] == "BOUGHT OUT ITEMS",   "TOTAL"].values[0]) if "BOUGHT OUT ITEMS" in bom_df["ITEM NAME"].values else 0
        encon_total      = float(bom_df.loc[bom_df["ITEM NAME"] == "ENCON ITEMS",        "TOTAL"].values[0]) if "ENCON ITEMS" in bom_df["ITEM NAME"].values else 0
        grand_total      = float(bom_df.loc[bom_df["ITEM NAME"] == "GRAND TOTAL",        "TOTAL"].values[0]) if "GRAND TOTAL" in bom_df["ITEM NAME"].values else 0

        # Build response — blower HP at user-selected pressure
        # Match the CFM logic in selection_engine: max(gas_cfm, oil_cfm)
        gas_cfm = blower_air_flow / 1.7
        oil_cfm = 0
        if f1_is_oil and equip1["burner"].get("equivalent_lph"):
            oil_cfm = equip1["burner"]["equivalent_lph"] * 10
        if f2_oil_lph_for_blower > 0:
            oil_cfm = max(oil_cfm, f2_oil_lph_for_blower * 10)
        cfm = max(gas_cfm, oil_cfm) if oil_cfm else gas_cfm
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
                "num_burners":                    n_burners,
                "avg_temp_rise":                  round(br1.avg_temp_rise, 2) if br1 else 0,
                "firing_rate_kcal":               round(br1.firing_rate_kcal, 2) if br1 else 0,
                "heat_load_kcal":                 round(br1.heat_load_kcal, 2) if br1 else 0,
                "fuel_consumption_nm3":           round(br1.fuel_consumption_nm3, 2) if br1 else 0,
                "calculated_firing_rate_nm3hr":   round(br1.calculated_firing_rate_nm3hr, 2) if br1 else round(ng_flow / 1.1, 2),
                "extra_firing_rate_nm3hr":        round(ng_flow, 2),
                "equivalent_lph":                 round(equip1["burner"].get("equivalent_lph", 0), 2),
                "fuel_density":                   equip1["burner"].get("fuel_density", 0),
                "fuel_density_unit":              equip1["burner"].get("fuel_density_unit", "kg/ltr"),
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
                # Actual combustion-air line NB the BOM components are sized to
                # (hydraulic pipe size, floored to 125 & air duct) — see selection_engine.
                "air_line_nb":  equip1.get("air_line_nb") or pipes1.air_pipe_nb,
            },
            "equipment": {
                "burner_model":   equip1["burner"]["model"],
                "burner_max_kcal_hr": equip1["burner"].get("max_firing_kcal_hr", 0),
                "burner_max_lph":     equip1["burner"].get("max_firing_lph", 0),
                "blower_model":   equip1["blower"]["model"],
                "blower_hp":      equip1["blower"]["hp"],
                "blower_airflow": equip1["blower"]["airflow_nm3hr"],
                "blower_cfm":     equip1["blower"].get("cfm", 0),
                "ng_gas_train":   f'{equip1["ng_gas_train"]["inlet_nb"]} x {equip1["ng_gas_train"]["outlet_nb"]} NB',
                # In dual-fuel offers the HPU may live on equip2 instead of
                # equip1 (e.g. fuel 1 = NG gas, fuel 2 = LDO oil). Pick whichever
                # side actually carries it so the Step-4 'Pumping Unit' row is
                # populated regardless of which fuel slot the oil is in.
                "hpu": (
                    (lambda h: f'{h["model"]} — {h["unit_kw"]} KW {h["variant"]}')(
                        (equip1 or {}).get("hpu") or (equip2 or {}).get("hpu")
                    )
                    if ((equip1 or {}).get("hpu") or (equip2 or {}).get("hpu"))
                    else None
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
            resp["calculations"]["fuel2_fuel_density_unit"] = equip2["burner"].get("fuel_density_unit", "kg/ltr") if equip2 else "kg/ltr"
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


# ──────────────────────────────────────────────────────────────────────────
# Cost Variations — runs the VLPH calc once per single-axis swap, applies
# the user's markup factor, and also surfaces markup-only variations
# (1.75 / 1.70 / 1.65 / 1.60) so the user can see total cost across every
# realistic lever for negotiating a deal closer to a target price.
# ──────────────────────────────────────────────────────────────────────────
class CostVariationsRequest(VLPHCalcRequest):
    markup: float = 1.80    # Frontend's bought-out markup multiplier.


@app.post("/api/cost-variations")
def cost_variations(req: CostVariationsRequest):
    """Iterate over the variation axes (vendor swaps, hpu_variant, hood_type,
    auto_control_type, control_mode, special_auto_ignition) plus markup
    levels and return each alternative's GRAND TOTAL (= bought_out * markup
    + encon_total).

    Response shape:
      {
        "current_total":   <float>,
        "current_markup":  <float>,
        "variations":      [{axis, value, total, saves}, ...],
        "best_combined":   {total, saves, swaps:[{axis,value}, ...]} | null
      }
    Variations that don't save money vs the baseline are dropped.
    """
    OIL_FUELS = {"hsd", "ldo", "hdo", "fo", "sko", "cfo", "lshs"}
    is_oil_offer = (
        req.fuel1_type in OIL_FUELS
        or (req.fuel2_type and req.fuel2_type in OIL_FUELS)
    )

    LABEL_MAP = {
        "control_valve_vendor": {
            "_axis": "Control Valve Vendor",
            "dembla": "DEMBLA", "aira": "AIRA", "cair": "CAIR",
        },
        "shutoff_valve_vendor": {
            "_axis": "Shutoff Valve Vendor",
            "dembla": "DEMBLA", "aira": "AIRA", "cair": "CAIR",
        },
        "butterfly_valve_vendor": {
            "_axis": "Butterfly Valve Vendor",
            "lt_lever": "L&T Lever", "lt_gear": "L&T Gear",
        },
        "pressure_gauge_vendor": {
            "_axis": "Pressure Gauge Vendor",
            "baumer": "BAUMER", "hguru": "HGURU",
        },
        "hpu_variant": {
            "_axis": "HPU Variant",
            "Simplex": "Simplex", "Duplex 1": "Duplex-I", "Duplex 2": "Duplex-II",
        },
        "hood_type": {
            "_axis": "Hood Movement",
            "up_down": "Up/Down (hydraulic)",
            "swivel_manual": "Manual Swivelling",
            "swivel_geared": "Geared Swivelling",
        },
        "auto_control_type": {
            "_axis": "Auto Control Type",
            "plc": "PLC", "plc_agr": "PLC + AGR", "pid": "PID",
        },
        "control_mode": {
            "_axis": "Control Mode",
            "automatic": "Automatic", "manual": "Manual",
        },
        "special_auto_ignition": {
            "_axis": "Auto Ignition",
            True: "Auto Ignition ON", False: "Auto Ignition OFF",
        },
    }

    axes = [
        ("control_valve_vendor",   ["dembla", "aira", "cair"]),
        ("shutoff_valve_vendor",   ["dembla", "aira", "cair"]),
        ("butterfly_valve_vendor", ["lt_lever", "lt_gear"]),
        ("pressure_gauge_vendor",  ["baumer", "hguru"]),
        ("hood_type",              ["up_down", "swivel_manual", "swivel_geared"]),
        ("special_auto_ignition",  [True, False]),
    ]
    if is_oil_offer:
        axes.append(("hpu_variant", ["Simplex", "Duplex 1", "Duplex 2"]))
    if req.control_mode == "automatic":
        axes.append(("auto_control_type", ["plc", "plc_agr", "pid"]))
    other_mode = "manual" if req.control_mode == "automatic" else "automatic"
    axes.append(("control_mode", [other_mode]))

    # Strip the markup field before passing to vlph_calculate so it round-trips
    # cleanly through the existing VLPHCalcRequest schema.
    def _to_calc_req(modified):
        d = modified.dict()
        d.pop("markup", None)
        return VLPHCalcRequest(**d)

    def _calc_costs(modified):
        """Returns (bought_out, encon) for the given request, or None."""
        try:
            resp = vlph_calculate(_to_calc_req(modified))
            cs = resp.get("cost_summary", {})
            return (
                float(cs.get("bought_out_total") or 0),
                float(cs.get("encon_total") or 0),
            )
        except Exception:
            return None

    def _grand(bought, encon, markup):
        return bought * markup + encon

    base = _calc_costs(req)
    if base is None:
        return {"error": "Could not compute baseline cost"}
    base_bought, base_encon = base
    base_total = _grand(base_bought, base_encon, req.markup)
    if base_total <= 0:
        return {"error": "Baseline cost is zero — check inputs"}

    variations = []
    best_per_axis = {}  # field -> (alt, total) of cheapest swap on this axis

    for field, alts in axes:
        cur_value = getattr(req, field)
        for alt in alts:
            if alt == cur_value:
                continue
            modified = req.copy(update={field: alt})
            costs = _calc_costs(modified)
            if costs is None:
                continue
            total = _grand(costs[0], costs[1], req.markup)
            # Drop variations where the calc returned 0 (silent failure) or
            # didn't actually save money against the baseline.
            if total <= 0:
                continue
            saves = base_total - total
            if saves <= 0:
                continue
            label = LABEL_MAP.get(field, {})
            variations.append({
                "axis":  label.get("_axis", field),
                "value": label.get(alt, str(alt)),
                "field": field,
                "alt":   alt,
                "total": round(total, 2),
                "saves": round(saves, 2),
            })
            cur_best = best_per_axis.get(field)
            if cur_best is None or total < cur_best[1]:
                best_per_axis[field] = (alt, total)

    # Markup-only variations — same BOM, different markup multiplier.
    for mk in [1.75, 1.70, 1.65, 1.60]:
        if abs(mk - req.markup) < 1e-6:
            continue
        total = _grand(base_bought, base_encon, mk)
        saves = base_total - total
        if saves <= 0:
            continue
        variations.append({
            "axis":  "Markup",
            "value": f"{mk:.2f}×",
            "field": "markup",
            "alt":   mk,
            "total": round(total, 2),
            "saves": round(saves, 2),
        })

    variations.sort(key=lambda r: r["saves"], reverse=True)

    # Combined-best: cheapest alt for each axis applied at once + lowest markup.
    best_combined = None
    update = {field: best for field, (best, _) in best_per_axis.items()}
    combined_req = req.copy(update=update) if update else req
    costs = _calc_costs(combined_req)
    if costs is not None:
        lowest_markup = 1.60 if req.markup > 1.60 else req.markup
        combined_total = _grand(costs[0], costs[1], lowest_markup)
        if combined_total > 0 and combined_total < base_total:
            swaps = []
            for field, (alt, _) in best_per_axis.items():
                label = LABEL_MAP.get(field, {})
                swaps.append({
                    "axis":  label.get("_axis", field),
                    "value": label.get(alt, str(alt)),
                })
            if abs(lowest_markup - req.markup) > 1e-6:
                swaps.append({"axis": "Markup", "value": f"{lowest_markup:.2f}×"})
            best_combined = {
                "total": round(combined_total, 2),
                "saves": round(base_total - combined_total, 2),
                "swaps": swaps,
            }

    return {
        "current_total":  round(base_total, 2),
        "current_markup": round(req.markup, 2),
        "variations":     variations,
        "best_combined":  best_combined,
    }


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
        is_dual = req.fuel2_type != "none" and req.fuel2_cv > 0

        if req.mode == "direct":
            # Direct mode: burner capacity in kW (× 860 → kcal/hr)
            if req.direct_heat_input_kcal_hr and req.direct_heat_input_kcal_hr > 0:
                heat_kcal_hr = req.direct_heat_input_kcal_hr
            else:
                heat_kcal_hr = req.direct_burner_capacity * 860
            ng_flow  = heat_kcal_hr / f1_cv
            air_flow = heat_kcal_hr * 118 / 100000
            br = None
        else:
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

        burner_pressure_wg = 36 if req.blower_pressure == "40" else 24

        # Dual fuel: blower sized for max(air1, air2)
        blower_air_flow = air_flow
        if is_dual:
            if req.mode == "direct":
                air_flow2_pre = air_flow
            else:
                br2_pre = calculate_burner(BurnerInputs(
                    Ti=req.Ti, Tf=req.Tf,
                    refractory_weight=req.refractory_weight,
                    fuel_cv=req.fuel2_cv,
                    time_taken_hr=req.time_taken_hr,
                    refractory_heat_factor=req.refractory_heat_factor,
                    efficiency=req.efficiency,
                ))
                air_flow2_pre = br2_pre.air_qty_nm3hr
            blower_air_flow = max(air_flow, air_flow2_pre)

        # Pre-compute fuel2 oil LPH for dual-fuel blower CFM
        f2_oil_lph_for_blower = 0
        if is_dual and req.fuel2_type in OIL_FUELS:
            heat_for_f2 = ng_flow * f1_cv if req.mode != "direct" else heat_kcal_hr
            f2_flow_kghr = heat_for_f2 / req.fuel2_cv if req.fuel2_cv > 0 else 0
            from bom.selectors.encon_burner import _get_fuel_density
            f2_density = _get_fuel_density(req.fuel2_type)
            f2_oil_lph_for_blower = f2_flow_kghr / f2_density if f2_density else 0

        equip1 = select_equipment(
            ng_flow_nm3hr=ng_flow, air_flow_nm3hr=blower_air_flow,
            is_dual_fuel=is_dual, fuel_cv=f1_cv,
            blower_pressure=req.blower_pressure, fuel_type=req.fuel1_type,
            hpu_variant=req.hpu_variant, burner_pressure_wg=burner_pressure_wg,
            butterfly_valve_vendor=req.butterfly_valve_vendor,
            shutoff_valve_vendor=req.shutoff_valve_vendor,
            control_mode=req.control_mode, auto_control_type=req.auto_control_type,
            fuel2_lph=f2_oil_lph_for_blower,
        )

        # Size pipes AFTER blower selection — gas pipe for fuel 1's flow,
        # air pipe for the SELECTED blower's RATED airflow (e.g. a 10 HP
        # blower rated at 1360 Nm3/hr — pipe must handle that, not just
        # the application's required flow).
        _rated_blower_air = float(equip1.get("blower", {}).get("airflow_nm3hr") or blower_air_flow)
        pipes1 = calculate_pipe_sizes(PipeInputs(ng_flow_nm3hr=ng_flow, air_flow_nm3hr=_rated_blower_air))

        f1_is_oil = req.fuel1_type in OIL_FUELS
        f1_oil_lph = equip1["burner"].get("equivalent_lph", 0) if f1_is_oil else 0

        # Fuel 2
        br2, equip2, ng_flow2, air_flow2 = None, None, 0, 0
        if is_dual:
            if req.mode == "direct":
                # Same heat output, fuel2 flow at fuel2 CV; air is CV-independent
                ng_flow2  = heat_kcal_hr / req.fuel2_cv
                air_flow2 = air_flow
            else:
                br2 = calculate_burner(BurnerInputs(
                    Ti=req.Ti, Tf=req.Tf,
                    refractory_weight=req.refractory_weight,
                    fuel_cv=req.fuel2_cv,
                    time_taken_hr=req.time_taken_hr,
                    refractory_heat_factor=req.refractory_heat_factor,
                    efficiency=req.efficiency,
                ))
                ng_flow2  = br2.extra_firing_rate_nm3hr
                air_flow2 = br2.air_qty_nm3hr
            equip2 = select_equipment(
                ng_flow_nm3hr=ng_flow2, air_flow_nm3hr=air_flow2,
                is_dual_fuel=is_dual, fuel_cv=req.fuel2_cv,
                blower_pressure=req.blower_pressure, fuel_type=req.fuel2_type,
                hpu_variant=req.hpu_variant, burner_pressure_wg=burner_pressure_wg,
                butterfly_valve_vendor=req.butterfly_valve_vendor,
                shutoff_valve_vendor=req.shutoff_valve_vendor,
                control_mode=req.control_mode, auto_control_type=req.auto_control_type,
            )
            # Pipe sizing for fuel-2 — needed by the front-end calc-result
            # panel ('Fuel-2 line size: gas flow / pipe NB'). Without this
            # the JS reads undefined and renders 'NaN Nm3/hr / undefined NB'.
            pipes2 = calculate_pipe_sizes(PipeInputs(
                ng_flow_nm3hr=ng_flow2, air_flow_nm3hr=air_flow2,
            ))
        else:
            pipes2 = None

        # Auto-fill fabrication + ceramic-fibre rolls from
        # fabrication_ladle_mapping (nearest horizontal row). Pipeline weight
        # is always user-supplied (file carries no horizontal pipeline data).
        from bom.vlph_builder import lookup_ladle_fab_pipeline
        _mapped_h = lookup_ladle_fab_pipeline(req.ladle_tons, "horizontal")
        _pipeline_kg_h  = req.pipeline_weight_kg
        # User-supplied override wins; fall back to DB-mapped value only when override is 0.
        _ms_override_h  = req.ms_structure_kg_override or (_mapped_h.get("fabrication_kg") if _mapped_h else 0)
        _ceramic_rolls_h = req.ceramic_rolls_override or (_mapped_h.get("ceramic_rolls") if _mapped_h else 0)

        if req.control_mode == "manual":
            bom_df = build_hlph_manual_df(
                equipment=equip1, ladle_tons=req.ladle_tons,
                fuel1_type=req.fuel1_type,
                purging_line=req.purging_line,
                pressure_gauge_vendor=req.pressure_gauge_vendor,
                pilot_burner=req.pilot_burner,
                pipeline_weight_kg=_pipeline_kg_h,
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
                pipeline_weight_kg=_pipeline_kg_h,
                purging_line=req.purging_line,
                ms_structure_kg_override=_ms_override_h,
                ceramic_rolls_override=_ceramic_rolls_h,
                special_auto_ignition=req.special_auto_ignition,
            )

        detail = bom_df[bom_df["MEDIA"] != ""].copy()
        bought_out_total = float(bom_df.loc[bom_df["ITEM NAME"] == "BOUGHT OUT ITEMS", "TOTAL"].values[0]) if "BOUGHT OUT ITEMS" in bom_df["ITEM NAME"].values else 0
        encon_total = float(bom_df.loc[bom_df["ITEM NAME"] == "ENCON ITEMS", "TOTAL"].values[0]) if "ENCON ITEMS" in bom_df["ITEM NAME"].values else 0
        grand_total = float(bom_df.loc[bom_df["ITEM NAME"] == "GRAND TOTAL", "TOTAL"].values[0]) if "GRAND TOTAL" in bom_df["ITEM NAME"].values else 0

        # Match CFM logic in selection_engine: max(gas_cfm, oil_cfm)
        gas_cfm = blower_air_flow / 1.7
        oil_cfm = 0
        if f1_is_oil and equip1["burner"].get("equivalent_lph"):
            oil_cfm = equip1["burner"]["equivalent_lph"] * 10
        if f2_oil_lph_for_blower > 0:
            oil_cfm = max(oil_cfm, f2_oil_lph_for_blower * 10)
        cfm = max(gas_cfm, oil_cfm) if oil_cfm else gas_cfm
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
                "mode": req.mode,
                "avg_temp_rise":              round(br.avg_temp_rise, 2)              if br else 0,
                "firing_rate_kcal":           round(br.firing_rate_kcal, 2)           if br else 0,
                "heat_load_kcal":             round(br.heat_load_kcal, 2)             if br else 0,
                "fuel_consumption_nm3":       round(br.fuel_consumption_nm3, 2)       if br else 0,
                "calculated_firing_rate_nm3hr": round(br.calculated_firing_rate_nm3hr, 2) if br else round(ng_flow / 1.1, 2),
                "extra_firing_rate_nm3hr":    round(ng_flow, 2),
                "equivalent_lph":             round(equip1["burner"].get("equivalent_lph", 0), 2),
                "fuel_density":               equip1["burner"].get("fuel_density", 0),
                "fuel_density_unit":          equip1["burner"].get("fuel_density_unit", "kg/ltr"),
                "final_firing_rate_mw":       round(br.final_firing_rate_mw, 2)       if br else round(ng_flow * f1_cv / (860 * 1000), 2),
                "air_qty_nm3hr":              round(air_flow, 2),
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
                "air_line_nb": equip1.get("air_line_nb") or pipes1.air_pipe_nb,
                "gas_train_flow": round(equip1["ng_gas_train"]["max_flow"], 0) if equip1.get("ng_gas_train") else 0,
                "gas_train_model": f'{equip1["ng_gas_train"]["inlet_nb"]} x {equip1["ng_gas_train"]["outlet_nb"]}' if equip1.get("ng_gas_train") else "",
            },
            "equipment": {
                "burner_model": equip1["burner"]["model"],
                "burner_max_kcal_hr": equip1["burner"].get("max_firing_kcal_hr", 0),
                "burner_max_lph":     equip1["burner"].get("max_firing_lph", 0),
                "blower_model": equip1["blower"]["model"],
                "blower_hp": equip1["blower"]["hp"],
                "blower_airflow": equip1["blower"]["airflow_nm3hr"],
                "blower_cfm":   equip1["blower"].get("cfm", 0),
                "ng_gas_train": f'{equip1["ng_gas_train"]["inlet_nb"]} x {equip1["ng_gas_train"]["outlet_nb"]} NB' if equip1.get("ng_gas_train") else "",
                # In dual-fuel offers the HPU may live on equip2 (e.g. fuel 1 NG
                # + fuel 2 LDO). Pick whichever side actually carries it.
                "hpu": (
                    (lambda h: f'{h["model"]} — {h["unit_kw"]} KW {h.get("variant", "")}'.strip())(
                        (equip1 or {}).get("hpu") or (equip2 or {}).get("hpu")
                    )
                    if ((equip1 or {}).get("hpu") or (equip2 or {}).get("hpu"))
                    else None
                ),
            },
            "bom": detail.to_dict(orient="records"),
            "cost_summary": {
                "bought_out_total": round(bought_out_total, 2),
                "encon_total": round(encon_total, 2),
                "grand_total": round(grand_total, 2),
            },
        }

        # Surface fuel-2 fields the front-end + offer pipeline rely on (mirrors
        # the VLPH endpoint).
        if is_dual and equip2 is not None:
            resp["calculations"]["fuel2_type"] = req.fuel2_type
            resp["calculations"]["fuel2_name"] = FUEL_NAMES.get(req.fuel2_type, req.fuel2_type)
            resp["calculations"]["fuel2_cv"]   = req.fuel2_cv
            resp["calculations"]["fuel2_extra_firing_rate_nm3hr"] = round(ng_flow2, 2)
            # Calc-mode fields the front-end needs to render Fuel Consumption
            # and Calc. Firing Rate in the result table. Direct mode falls
            # back to deriving from ng_flow2 (the +10% extra rate).
            resp["calculations"]["fuel2_consumption_nm3"] = (
                round(br2.fuel_consumption_nm3, 2) if br2 else 0
            )
            resp["calculations"]["fuel2_firing_rate_nm3hr"] = (
                round(br2.calculated_firing_rate_nm3hr, 2) if br2 else round(ng_flow2 / 1.1, 2)
            )
            resp["calculations"]["fuel2_equivalent_lph"] = round(
                equip2["burner"].get("equivalent_lph", 0), 2
            )
            resp["calculations"]["fuel2_fuel_density"] = equip2["burner"].get("fuel_density", 0)
            # Pipe-line size info for the Client Data Summary card. For oil
            # fuel-2 we use the oil pipe NB (sized from LPH); for gas we use
            # the standard pipe NB from calculate_pipe_sizes.
            f2_is_oil = req.fuel2_type in OIL_FUELS
            f2_oil_lph = (equip2["burner"].get("equivalent_lph", 0) if f2_is_oil else 0)
            f2_oil_nb = select_oil_pipe_nb(f2_oil_lph) if f2_is_oil else 0
            resp["pipes"]["fuel2_label"]  = FUEL_NAMES.get(req.fuel2_type, "Fuel 2")
            resp["pipes"]["fuel2_flow"]   = round(ng_flow2, 2)
            resp["pipes"]["fuel2_is_oil"] = f2_is_oil
            resp["pipes"]["fuel2_oil_lph"] = round(f2_oil_lph, 2) if f2_is_oil else None
            if pipes2 is not None:
                resp["pipes"]["fuel2_dia_mm"] = (
                    round(pipes2.ng_pipe_inner_dia_mm, 2) if not f2_is_oil else f2_oil_nb
                )
                resp["pipes"]["fuel2_nb"] = (
                    pipes2.ng_pipe_nb if not f2_is_oil else f2_oil_nb
                )
            if not f2_is_oil and equip2.get("ng_gas_train"):
                resp["pipes"]["fuel2_gas_train_flow"] = round(
                    equip2["ng_gas_train"]["max_flow"], 0
                )
                resp["pipes"]["fuel2_gas_train_model"] = (
                    f'{equip2["ng_gas_train"]["inlet_nb"]} x '
                    f'{equip2["ng_gas_train"]["outlet_nb"]}'
                )

        return resp
    except Exception as e:
        import traceback
        return {"error": str(e), "trace": traceback.format_exc()}


# ──────────────────────────────────────────────────────────────────────────
# HLPH Cost Variations — same idea as /api/cost-variations but runs the
# horizontal-ladle-preheater calc instead.
# ──────────────────────────────────────────────────────────────────────────
@app.post("/api/hlph-cost-variations")
def hlph_cost_variations(req: CostVariationsRequest):
    """Run hlph_calculate for every single-axis swap + markup level and
    return each alternative's grand total. Same response shape as
    /api/cost-variations."""
    OIL_FUELS = {"hsd", "ldo", "hdo", "fo", "sko", "cfo", "lshs"}
    is_oil_offer = (
        req.fuel1_type in OIL_FUELS
        or (req.fuel2_type and req.fuel2_type in OIL_FUELS)
    )

    LABEL_MAP = {
        "control_valve_vendor": {
            "_axis": "Control Valve Vendor",
            "dembla": "DEMBLA", "aira": "AIRA", "cair": "CAIR",
        },
        "shutoff_valve_vendor": {
            "_axis": "Shutoff Valve Vendor",
            "dembla": "DEMBLA", "aira": "AIRA", "cair": "CAIR",
        },
        "butterfly_valve_vendor": {
            "_axis": "Butterfly Valve Vendor",
            "lt_lever": "L&T Lever", "lt_gear": "L&T Gear",
        },
        "pressure_gauge_vendor": {
            "_axis": "Pressure Gauge Vendor",
            "baumer": "BAUMER", "hguru": "HGURU",
        },
        "hpu_variant": {
            "_axis": "HPU Variant",
            "Simplex": "Simplex", "Duplex 1": "Duplex-I", "Duplex 2": "Duplex-II",
        },
        "auto_control_type": {
            "_axis": "Auto Control Type",
            "plc": "PLC", "plc_agr": "PLC + AGR", "pid": "PID",
        },
        "control_mode": {
            "_axis": "Control Mode",
            "automatic": "Automatic", "manual": "Manual",
        },
        "special_auto_ignition": {
            "_axis": "Auto Ignition",
            True: "Auto Ignition ON", False: "Auto Ignition OFF",
        },
    }

    # HLPH does NOT vary by hood_type (horizontal uses a trolley drive,
    # not a swivel/up-down hood mechanism).
    axes = [
        ("control_valve_vendor",   ["dembla", "aira", "cair"]),
        ("shutoff_valve_vendor",   ["dembla", "aira", "cair"]),
        ("butterfly_valve_vendor", ["lt_lever", "lt_gear"]),
        ("pressure_gauge_vendor",  ["baumer", "hguru"]),
        ("special_auto_ignition",  [True, False]),
    ]
    if is_oil_offer:
        axes.append(("hpu_variant", ["Simplex", "Duplex 1", "Duplex 2"]))
    if req.control_mode == "automatic":
        axes.append(("auto_control_type", ["plc", "plc_agr", "pid"]))
    other_mode = "manual" if req.control_mode == "automatic" else "automatic"
    axes.append(("control_mode", [other_mode]))

    def _to_calc_req(modified):
        d = modified.dict()
        d.pop("markup", None)
        return VLPHCalcRequest(**d)

    def _calc_costs(modified):
        try:
            resp = hlph_calculate(_to_calc_req(modified))
            cs = resp.get("cost_summary", {})
            return (
                float(cs.get("bought_out_total") or 0),
                float(cs.get("encon_total") or 0),
            )
        except Exception:
            return None

    def _grand(bought, encon, markup):
        return bought * markup + encon

    base = _calc_costs(req)
    if base is None:
        return {"error": "Could not compute baseline cost"}
    base_bought, base_encon = base
    base_total = _grand(base_bought, base_encon, req.markup)
    if base_total <= 0:
        return {"error": "Baseline cost is zero — check inputs"}

    variations = []
    best_per_axis = {}

    for field, alts in axes:
        cur_value = getattr(req, field)
        for alt in alts:
            if alt == cur_value:
                continue
            modified = req.copy(update={field: alt})
            costs = _calc_costs(modified)
            if costs is None:
                continue
            total = _grand(costs[0], costs[1], req.markup)
            if total <= 0:
                continue
            saves = base_total - total
            if saves <= 0:
                continue
            label = LABEL_MAP.get(field, {})
            variations.append({
                "axis":  label.get("_axis", field),
                "value": label.get(alt, str(alt)),
                "field": field,
                "alt":   alt,
                "total": round(total, 2),
                "saves": round(saves, 2),
            })
            cur_best = best_per_axis.get(field)
            if cur_best is None or total < cur_best[1]:
                best_per_axis[field] = (alt, total)

    for mk in [1.75, 1.70, 1.65, 1.60]:
        if abs(mk - req.markup) < 1e-6:
            continue
        total = _grand(base_bought, base_encon, mk)
        saves = base_total - total
        if saves <= 0:
            continue
        variations.append({
            "axis":  "Markup",
            "value": f"{mk:.2f}×",
            "field": "markup",
            "alt":   mk,
            "total": round(total, 2),
            "saves": round(saves, 2),
        })

    variations.sort(key=lambda r: r["saves"], reverse=True)

    best_combined = None
    update = {field: best for field, (best, _) in best_per_axis.items()}
    combined_req = req.copy(update=update) if update else req
    costs = _calc_costs(combined_req)
    if costs is not None:
        lowest_markup = 1.60 if req.markup > 1.60 else req.markup
        combined_total = _grand(costs[0], costs[1], lowest_markup)
        if combined_total > 0 and combined_total < base_total:
            swaps = []
            for field, (alt, _) in best_per_axis.items():
                label = LABEL_MAP.get(field, {})
                swaps.append({
                    "axis":  label.get("_axis", field),
                    "value": label.get(alt, str(alt)),
                })
            if abs(lowest_markup - req.markup) > 1e-6:
                swaps.append({"axis": "Markup", "value": f"{lowest_markup:.2f}×"})
            best_combined = {
                "total": round(combined_total, 2),
                "saves": round(base_total - combined_total, 2),
                "swaps": swaps,
            }

    return {
        "current_total":  round(base_total, 2),
        "current_markup": round(req.markup, 2),
        "variations":     variations,
        "best_combined":  best_combined,
    }


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


# ── Recuperator (USK Exports model) ─────────────────────────────────────────

@app.get("/recup", response_class=HTMLResponse)
def recup_costing_form():
    html_path = os.path.join(BASE_DIR, "recup_costing.html")
    with open(html_path, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())


class RecupCalcRequest(BaseModel):
    # Flue gas (only flow + inlet temp are inputs; mass and final temp are derived)
    flue_flow_nm3hr:    float = 1900.0
    flue_temp_in_C:     float = 800.0
    cp_flue_kcal_kgC:   float = 0.23
    # Combustion air
    air_volume_nm3hr:   float = 1750.0
    air_temp_in_C:      float = 35.0
    air_temp_out_C:     float = 400.0
    cp_air_kcal_kgC:    float = 0.247
    # Geometry / overrides
    heat_transfer_coef: float = 30.0
    pipe_dia_mm:        float = 48.3
    pipe_thick_mm:      float = 2.77
    pipe_kg_per_m:      float = 3.16
    pipe_length_m_per_bank: float = 0.63
    bank_gap_mm:        float = 150.0
    pipes_total_override: int = 0    # 0 = auto-derive from surface area
    hot_bank_material:    str = "SS" # "SS" (Rs 250/kg) or "MS" (Rs 70/kg)
    cold_bank_material:   str = "SS" # can differ from hot bank
    side_hood_kg:         float = 1500.0   # MS side hood weight
    cai_rate_override:    float = 0.0      # CAI Assembly Rs/kg override; 0 = use DB default


@app.post("/api/recup-calculate")
def recup_calculate(req: RecupCalcRequest):
    try:
        from calculations.recup import RecupInputs, calculate_recup
        from bom.recup_builder import build_recup_df, recup_summary, _load_rates

        results = calculate_recup(RecupInputs(
            flue_flow_nm3hr=req.flue_flow_nm3hr,
            flue_temp_in_C=req.flue_temp_in_C,
            cp_flue_kcal_kgC=req.cp_flue_kcal_kgC,
            air_volume_nm3hr=req.air_volume_nm3hr,
            air_temp_in_C=req.air_temp_in_C,
            air_temp_out_C=req.air_temp_out_C,
            cp_air_kcal_kgC=req.cp_air_kcal_kgC,
            heat_transfer_coef=req.heat_transfer_coef,
            pipe_dia_mm=req.pipe_dia_mm,
            pipe_thick_mm=req.pipe_thick_mm,
            pipe_kg_per_m=req.pipe_kg_per_m,
            pipe_length_m_per_bank=req.pipe_length_m_per_bank,
            bank_gap_mm=req.bank_gap_mm,
            pipes_total_override=req.pipes_total_override,
            hot_bank_material=req.hot_bank_material,
            cold_bank_material=req.cold_bank_material,
            side_hood_kg=req.side_hood_kg,
            cai_rate_override=req.cai_rate_override,
        ))

        rates = _load_rates()
        df = build_recup_df(results, rates)
        summary = recup_summary(results, rates)

        detail = df[df["MEDIA"] != ""].copy()

        return {
            "calculations": {
                "flue_mass_kghr":       results.flue_mass_kghr,
                "heat_required_kcal":   results.heat_required_kcal,
                "flue_temp_out_C":      results.flue_temp_out_C,
                "lmtd_C":               results.lmtd_C,
                "surface_area_m2":      results.surface_area_m2,
                "pipes_total_raw":      results.pipes_total_raw,
                "pipes_total":          results.pipes_total,
                "pipes_in_row":         results.pipes_in_row,
                "pipes_in_column":      results.pipes_in_column,
                "pipes_per_bank":       results.pipes_per_bank,
                "bank_length_mm":       results.bank_length_mm,
                "bank_width_mm":        results.bank_width_mm,
                "weight_per_pipe_kg":   results.weight_per_pipe_kg,
                "weight_hot_bank_kg":   results.weight_hot_bank_kg,
                "weight_cold_bank_kg":  results.weight_cold_bank_kg,
                "weight_total_pipes_kg": results.weight_total_pipes_kg,
                "ms_outer_shell_kg":    results.ms_outer_shell_kg,
                "ms_air_inlet_duct_kg": results.ms_air_inlet_duct_kg,
                "ms_hot_outlet_duct_kg": results.ms_hot_outlet_duct_kg,
                "ms_pipe_holding_kg":   results.ms_pipe_holding_kg,
                "ms_bottom_box_kg":     results.ms_bottom_box_kg,
                "hot_bank_material":    results.hot_bank_material,
                "cold_bank_material":   results.cold_bank_material,
                "cai_rate_override":    results.cai_rate_override,
                # Echo of inputs — the offer template needs these for the
                # Designing Parameters table (flue/air conditions are
                # process inputs, not derived values).
                "flue_flow_nm3hr":      req.flue_flow_nm3hr,
                "flue_temp_in_C":       req.flue_temp_in_C,
                "air_volume_nm3hr":     req.air_volume_nm3hr,
                "air_temp_in_C":        req.air_temp_in_C,
                "air_temp_out_C":       req.air_temp_out_C,
                "pipe_dia_mm":          req.pipe_dia_mm,
                "pipe_thick_mm":        req.pipe_thick_mm,
                "pipe_length_m_per_bank": req.pipe_length_m_per_bank,
            },
            "bom": detail[["MEDIA","ITEM NAME","REFERENCE","QTY","MAKE","UNIT PRICE","TOTAL"]].to_dict(orient="records"),
            "cost_summary": summary,
        }
    except Exception as e:
        import traceback
        return {"error": str(e), "trace": traceback.format_exc()}


class RecupQuoteRequest(BaseModel):
    # Customer / project
    salutation:      str = ""
    poc_name:        str = ""
    poc_designation: str = ""
    company_name:    str = ""
    company_address: str = ""
    email:           str = ""
    mobile_no:       str = ""
    project_name:    str = ""
    subject:         str = ""
    application:     str = ""
    client_enq_ref:  str = ""
    # Marketing / ref
    marketing_person: str = ""
    marketing_email:  str = ""
    marketing_phone:  str = ""
    technical_person: str = ""
    technical_phone:  str = ""
    technical_email:  str = ""
    location:         str = "FBD"
    # Supervision charges (VLPH-style supervision table in Annexure III)
    supervision_mech: str = ""
    supervision_plc:  str = ""
    # T&C
    tnc_prices:             str = ""
    tnc_delivery:           str = ""
    tnc_gst:                str = ""
    tnc_hsn_code:           str = ""
    tnc_pan_gst:            str = ""
    tnc_payment_terms:      str = ""
    tnc_packing_forwarding: str = ""
    tnc_freight:            str = ""
    tnc_transit_insurance:  str = ""
    tnc_validity:           str = ""
    tnc_inspection:         str = ""
    tnc_guarantee:          str = ""
    # Recup calc payload
    calculations: dict = {}
    bom:          list = []        # the full BOM detail rows for the price schedule
    final_total:  float = 0.0
    grand_total:  float = 0.0
    qty:          int = 1
    # Price schedule style — "single" (VLPH-like, one row + amount in words)
    # or "full" (Viraj-like, every BOM line + supervision + totals)
    price_schedule_style: str = "single"
    # Supervision charges row (optional — only meaningful in 'full' style)
    supervision_include: bool = False
    supervision_rate:    str = "Rs. 10,500 / Man / Day"
    supervision_note:    str = "Plus: To and Fro fare from Delhi to Site, Boarding and Lodging, Local Conveyance, Medical assistance"


def _material_of_construction(hot_mat: str, cold_mat: str, pipe_dia_mm: float = 48.3) -> dict:
    """Translate Step-2's Hot/Cold tube-material picks into the four
    placeholders the Material of Construction table in Annexure I needs.
    SS -> 'SS-316' (typical hot-bank grade), MS -> 'Mild Steel (CS)'.
    The Tube line carries the bend/straight + size + schedule spec."""
    def block(mat: str | None, bend: bool):
        m = (mat or "SS").upper()
        if m == "MS":
            plate = "Mild Steel (CS)"
            tube  = f"CS, ERW {'Bended' if bend else 'Straight'} pipes, OD: {pipe_dia_mm}, Schedule-40"
        else:
            plate = "SS-316"
            tube  = f"SS-316 {'Bended' if bend else 'Straight'}, ERW pipes, OD: {pipe_dia_mm}, Schedule-40"
        return plate, tube
    hot_plate, hot_tube   = block(hot_mat,  bend=True)   # hot bank is bent (for expansion)
    cold_plate, cold_tube = block(cold_mat, bend=False)  # cold bank is straight
    return {
        "hot_tube_plate_material":  hot_plate,
        "hot_tube_material":        hot_tube,
        "cold_tube_plate_material": cold_plate,
        "cold_tube_material":       cold_tube,
    }


def _bom_rows_for_offer(bom: list, *, supervision_include: bool,
                        supervision_rate: str, supervision_note: str,
                        single_line_total: float = 0.0,
                        single_line_mode: bool = False) -> dict:
    """Translate the calculator's BOM detail (12 line items + 3 summary
    rows) into the placeholders the offer's Price Schedule needs:
      - bom_rows               iterable list of {sno,item,qty,unit_price,total}
      - bought_out_total       formatted ₹
      - encon_total            formatted ₹
      - grand_total            formatted ₹
      - grand_total_in_words   Indian-English words
      - supervision_*          row content (only rendered when flag is on)
    Summary rows from the BOM dataframe (MEDIA == '') are skipped here;
    we recompute totals from the line items so user-edited unit prices
    on Step 3 propagate to the offer.
    """
    from engine.quote_writer import amount_in_words_indian as _words, _format_inr

    rows = []
    bought_out = 0.0
    encon      = 0.0
    for r in bom or []:
        media = (r.get("MEDIA") or "").strip()
        if not media:
            continue  # skip the BOUGHT OUT / ENCON / GRAND TOTAL summary rows
        total = float(r.get("TOTAL") or 0)
        unit  = float(r.get("UNIT PRICE") or 0)
        qty   = r.get("QTY") or 1
        rows.append({
            "sno":        len(rows) + 1,
            "item":       r.get("ITEM NAME", ""),
            "qty":        qty,
            "unit_price": _format_inr(unit),
            "total":      _format_inr(total),
        })
        if media == "ENCON ITEMS":
            encon += total
        else:
            bought_out += total
    grand = bought_out + encon
    # In single-line mode the footer should show the FINAL price (the one
    # rendered in the lone Price Schedule row), not the raw BOM sum.
    footer_total = single_line_total if single_line_mode else grand
    return {
        "bom_rows":            rows,
        "bought_out_total":    _format_inr(bought_out),
        "encon_total":         _format_inr(encon),
        "grand_total":         _format_inr(footer_total),
        "grand_total_in_words": f"INR. {_words(footer_total)} ONLY.",
        "supervision_include": bool(supervision_include),
        "supervision_rate":    supervision_rate or "",
        "supervision_note":    supervision_note or "",
    }


@app.post("/api/generate-recup-quote")
def generate_recup_quote(req: RecupQuoteRequest):
    """Render Recup_Offer_Template.docx with payload data, save to
    quotes/, and optionally convert to PDF.

    Template note: the recup template is cloned from the VLPH offer
    template (Offer_Template.docx) via build_recup_template_from_vlph.py
    — see that script for the full layout (cover page, Annexures, MoC
    sub-table + 3D image block + Designing Parameters table + scope)."""
    try:
        from datetime import datetime as _dt
        from docxtpl import DocxTemplate
        from engine.quote_writer import amount_in_words_indian, _format_inr

        seq = next_quote_seq()
        full_ref = build_enquiry_ref(seq, req.technical_person or "", req.location or "")
        # Split full ref (e.g. "ENCON.04026.050/FBD/JS DT.16/05/2026")
        # into the short part and the date.
        short_ref = full_ref.split(" DT.")[0]
        date_str = full_ref.split(" DT.")[-1] if " DT." in full_ref else _dt.now().strftime("%d/%m/%Y")

        c = req.calculations or {}
        unit_price = float(req.final_total or 0)
        qty = max(1, int(req.qty or 1))
        total_price = unit_price * qty

        ctx = {
            "project_name":     req.project_name or (f"Recuperator for {req.application}" if req.application else "Recuperator"),
            "subject":          req.subject or (f"Offer for Recuperator — {req.application}" if req.application else "Offer for Recuperator"),
            "application":      req.application or "Furnace",
            # VLPH-style top-of-document equipment name (renders on cover
            # page + 'About the Equipment' section + Annexure I banner).
            "equipment_name":   f"Recuperator for {req.application}" if req.application else "Recuperator",
            "company_name":     req.company_name,
            "company_address":  req.company_address,
            "email":            req.email,
            "mobile_no":        req.mobile_no,
            "poc_name":         _with_salutation(req.salutation, req.poc_name),
            "poc_designation":  req.poc_designation or "",
            "client_enq_ref":   req.client_enq_ref,
            "enquiry_ref":      full_ref,
            "enquiry_ref_short": short_ref,
            "enquiry_date_str":  date_str,
            "marketing_person": req.marketing_person,
            "marketing_email":  req.marketing_email,
            "marketing_phone":  req.marketing_phone,
            # Technical person (cover-letter signature in VLPH layout)
            "technical_person": req.technical_person or req.marketing_person,
            "technical_phone":  req.technical_phone  or req.marketing_phone,
            "technical_email":  req.technical_email  or req.marketing_email,
            # Supervision charges - VLPH renders these in a small sub-table
            # right after the Price Schedule. For recup we expose mech +
            # plc lines (PLC line may stay blank if not in scope).
            "supervision_mech": req.supervision_mech or "",
            "supervision_plc":  req.supervision_plc  or "",
            # Designing parameters table
            "flue_flow_nm3hr":  f"{c.get('flue_flow_nm3hr', 0):,.0f}" if c.get('flue_flow_nm3hr') else "",
            "flue_temp_in_C":   f"{int(round(c.get('flue_temp_in_C', 0)))}" if c.get('flue_temp_in_C') else "",
            "flue_temp_out_C":  f"{int(round(c.get('flue_temp_out_C', 0)))}" if c.get('flue_temp_out_C') else "",
            "air_volume_nm3hr": f"{c.get('air_volume_nm3hr', 0):,.0f}" if c.get('air_volume_nm3hr') else "",
            "air_temp_in_C":    f"{int(round(c.get('air_temp_in_C', 0)))}" if c.get('air_temp_in_C') else "",
            "air_temp_out_C":   f"{int(round(c.get('air_temp_out_C', 0)))}" if c.get('air_temp_out_C') else "",
            # Marketing-friendly rounded-up version of the air outlet
            # temperature: ceil to the next multiple of 50 °C. Used in
            # the Annexure I description paragraph (e.g. an exact
            # value of 412 °C reads as 'about 450 °C').
            "air_temp_out_C_rounded": (
                f"{int(((c.get('air_temp_out_C', 0) + 49.999) // 50) * 50)}"
                if c.get('air_temp_out_C') else ""
            ),
            "surface_area_m2":  f"{c.get('surface_area_m2', 0):.2f}",
            "pipe_dia_mm":      f"{c.get('pipe_dia_mm', 48.3):.1f}" if c.get('pipe_dia_mm') else "48.3",
            "pipe_length_m":    f"{c.get('pipe_length_m_per_bank', 0.63):.2f}" if c.get('pipe_length_m_per_bank') else "0.63",
            "pipe_thick_mm":    f"{c.get('pipe_thick_mm', 2.77):.2f}" if c.get('pipe_thick_mm') else "2.77",
            # Material of Construction (Annexure I — auto-derived from
            # the engineer's Hot/Cold material picks on Step 2).
            **_material_of_construction(c.get('hot_bank_material'),
                                        c.get('cold_bank_material'),
                                        c.get('pipe_dia_mm', 48.3)),
            # Price schedule (legacy single-line placeholders — kept for
            # backwards compatibility in case any text still references them)
            "recup_qty":           f"{qty:02d} No.",
            "recup_unit_price":    _format_inr(unit_price),
            "recup_total_price":   _format_inr(total_price),
            "recup_total_in_words": f"INR. {amount_in_words_indian(total_price)} ONLY.",
            # Price-schedule mode flag (template branches on this)
            "price_schedule_style": (req.price_schedule_style or "single").lower(),
            # Full-BOM iterable price schedule (Annexure III, 'full' mode)
            **_bom_rows_for_offer(req.bom or [],
                                  supervision_include=req.supervision_include,
                                  supervision_rate=req.supervision_rate,
                                  supervision_note=req.supervision_note,
                                  single_line_total=total_price,
                                  single_line_mode=(req.price_schedule_style or 'single').lower() == 'single'),
            # T&C
            "tnc_prices":             req.tnc_prices,
            "tnc_delivery":           req.tnc_delivery,
            "tnc_gst":                req.tnc_gst,
            "tnc_hsn_code":           req.tnc_hsn_code,
            "tnc_pan_gst":            req.tnc_pan_gst,
            "tnc_payment_terms":      req.tnc_payment_terms,
            "tnc_packing_forwarding": req.tnc_packing_forwarding,
            "tnc_freight":            req.tnc_freight,
            "tnc_transit_insurance":  req.tnc_transit_insurance,
            "tnc_validity":           req.tnc_validity,
            "tnc_inspection":         req.tnc_inspection,
            "tnc_guarantee":          req.tnc_guarantee,
        }

        tpl_path = os.path.join(BASE_DIR, "Recup_Offer_Template.docx")
        tpl = DocxTemplate(tpl_path)

        # 3D recuperator images (Annexure I) — embed only if the JPEGs are
        # present in static/recup/. Missing files render as blank paragraphs
        # rather than crashing the whole offer.
        from docxtpl import InlineImage
        from docx.shared import Mm
        _img_dir = os.path.join(BASE_DIR, "static", "recup")
        for key, fname, width_mm in (
            ("image_recup_side",  "recup_3d_side.jpeg",  120),
            ("image_recup_front", "recup_3d_front.jpeg", 120),
            ("image_recup_top",   "recup_3d_top.jpeg",   120),
        ):
            fpath = os.path.join(_img_dir, fname)
            ctx[key] = InlineImage(tpl, fpath, width=Mm(width_mm)) if os.path.exists(fpath) else ""

        tpl.render(ctx)

        # Save with a safe, sequenced filename
        safe_company = "".join(ch for ch in (req.company_name or "Client") if ch.isalnum() or ch in " _-").strip().replace(" ", "_") or "Client"
        docx_name = f"Recup_Offer_{safe_company}_{seq}.docx"
        docx_path = os.path.join(QUOTES_FOLDER, docx_name)
        tpl.save(docx_path)

        # PDF (best effort — LibreOffice may not be present locally)
        pdf_name = docx_name.replace(".docx", ".pdf")
        pdf_path = os.path.join(QUOTES_FOLDER, pdf_name)
        pdf_ok = _docx_to_pdf(docx_path, pdf_path)

        # Mirror docx + pdf to Google Drive (recup-specific folder) in a
        # background thread. product_type='recuperator' triggers the
        # recup routing in drive_uploader._folder_id_for_product.
        # Failures are logged but never break offer generation.
        try:
            from engine.drive_uploader import upload_offer_async
            upload_offer_async(docx_path, docx_name, "recuperator")
            if pdf_ok:
                upload_offer_async(pdf_path, pdf_name, "recuperator")
        except Exception as _drv_err:
            print(f"WARN: drive upload kickoff failed: {_drv_err}")

        return {
            "filename":     docx_name,
            "pdf_filename": pdf_name if pdf_ok else None,
            "download_url": f"/api/download-quote/{docx_name}",
            "pdf_url":      f"/api/pdf-quote/{pdf_name}" if pdf_ok else None,
            # Preview endpoint converts DOCX -> HTML via mammoth, so it
            # always uses the docx filename (works whether PDF rendered
            # or not).
            "preview_url":  f"/api/preview-quote/{docx_name}",
            "quote_no":     full_ref,
            "enquiry_ref":  full_ref,
            "final_total":  unit_price,
            "total_price":  total_price,
        }
    except Exception as e:
        import traceback
        return {"error": str(e), "trace": traceback.format_exc()}


# ── HPU stand-alone offer ────────────────────────────────────────────────────
class HpuCustomer(BaseModel):
    salutation: Optional[str] = ""
    name: Optional[str] = ""
    email: Optional[str] = ""
    phone: Optional[str] = ""
    company: Optional[str] = ""
    designation: Optional[str] = ""
    address: Optional[str] = ""
    city: Optional[str] = ""
    state: Optional[str] = ""
    pin: Optional[str] = ""
    gstin: Optional[str] = ""
    ref_no: Optional[str] = ""
    location: Optional[str] = ""
    subject: Optional[str] = ""
    marketing_salutation: Optional[str] = ""
    marketing: Optional[str] = ""
    marketing_email: Optional[str] = ""
    marketing_phone: Optional[str] = ""
    technical_salutation: Optional[str] = ""
    technical: Optional[str] = ""
    technical_email: Optional[str] = ""
    technical_phone: Optional[str] = ""
    # Annexure IV — Terms & Conditions (HPU Step 3, mirrored from VLPH).
    tnc_prices: Optional[str] = ""
    tnc_delivery: Optional[str] = ""
    tnc_gst: Optional[str] = ""
    tnc_hsn_code: Optional[str] = ""
    tnc_pan_gst: Optional[str] = ""
    tnc_payment_terms: Optional[str] = ""
    tnc_packing_forwarding: Optional[str] = ""
    tnc_freight: Optional[str] = ""
    tnc_transit_insurance: Optional[str] = ""
    tnc_validity: Optional[str] = ""
    tnc_inspection: Optional[str] = ""
    tnc_guarantee: Optional[str] = ""


class HpuQuoteRequest(BaseModel):
    customer: HpuCustomer
    hpu_variant: str
    hpu_kw: float
    qty: int = 1
    # Oil flow rate (LPH) is catalog-driven — the backend looks it up
    # from pumping_unit_price.flow_lph using (kw, variant). The client
    # may post it as a hint but it is overridden by the DB value.
    fuel_lph: Optional[float] = None
    # Legacy: kept optional so older clients posting fuel_type don't 400.
    # No longer surfaced in the offer doc.
    fuel_type: Optional[str] = ""


@app.get("/api/hpu/flow-lph")
def hpu_flow_lph(kw: float, variant: str):
    """Look up the catalog oil flow rate (LPH) for a given (kW, variant)
    from pumping_unit_price.flow_lph. Used by the HPU form to auto-fill
    the Oil Flow Rate field so the user can't type a wrong number."""
    try:
        conn = sqlite3.connect(DB_PATH)
        row = conn.execute(
            "SELECT flow_lph, model_code, sell_price FROM pumping_unit_price "
            "WHERE unit_kw = ? AND variant = ? LIMIT 1",
            (int(round(kw)), variant),
        ).fetchone()
        conn.close()
        if not row:
            return {"error": f"no row for {variant} @ {kw} kW", "flow_lph": None}
        return {
            "flow_lph":   float(row[0]) if row[0] is not None else None,
            "model_code": row[1] or "",
            "sell_price": float(row[2]) if row[2] is not None else None,
        }
    except Exception as e:
        return {"error": str(e), "flow_lph": None}


def _generate_pumping_unit_offer(req: "HpuQuoteRequest", *, mode: str) -> dict:
    """Shared body for the HPU and PU stand-alone offer endpoints.

    The two products share the form, the template shell and the
    render path; only price source, model-code prefix, equipment
    label, filename and template file differ. `mode` is "hpu" or
    "pu" and picks all five.
    """
    from engine.quote_engine import calculate_quote
    from engine.quote_writer import generate_quote_docx
    from bom.hpu_calculator import HPU_MARKUP
    from bom.selectors.hpu_selector import VARIANT_PREFIX, PU_VARIANT_PREFIX

    cust = req.customer
    kw_int = int(round(req.hpu_kw))

    # ── 1. Unit price + catalog flow rate ──────────────────────────────
    # Flow rate (LPH) is identical for HPU and PU since the pump skid
    # is the same — read from pumping_unit_price.flow_lph either way.
    conn = sqlite3.connect(DB_PATH)
    flow_row = conn.execute(
        "SELECT flow_lph FROM pumping_unit_price "
        "WHERE unit_kw = ? AND variant = ? LIMIT 1",
        (kw_int, req.hpu_variant),
    ).fetchone()

    if mode == "hpu":
        # HPU sell price = SUM(hpu_master.amount) × HPU_MARKUP (1.8).
        # See bom/hpu_calculator.py for the formula source.
        cost_row = conn.execute(
            "SELECT SUM(amount) FROM hpu_master "
            "WHERE unit_kw = ? AND variant = ?",
            (kw_int, req.hpu_variant),
        ).fetchone()
        conn.close()
        if not cost_row or cost_row[0] is None:
            return {"error": f"No HPU rows in hpu_master for "
                             f"{req.hpu_variant} @ {req.hpu_kw} kW."}
        unit_price = round(float(cost_row[0]) * HPU_MARKUP, 2)
        model_code = f"{VARIANT_PREFIX[req.hpu_variant]}-{kw_int}"
        equipment_label   = "Heating and Pumping Unit"
        item_product_type = "Hydraulic Pumping Unit"   # gates preheater scope
        filename_infix    = "HPU"
        template_name     = "HPU_Offer_Template.docx"
    else:  # mode == "pu"
        # PU sell price is already markup-included in pumping_unit_price
        # (rebuilt from hpu_master minus heater/thermostat × 1.8).
        price_row = conn.execute(
            "SELECT sell_price FROM pumping_unit_price "
            "WHERE unit_kw = ? AND variant = ? LIMIT 1",
            (kw_int, req.hpu_variant),
        ).fetchone()
        conn.close()
        if not price_row or price_row[0] is None:
            return {"error": f"No PU price in pumping_unit_price for "
                             f"{req.hpu_variant} @ {req.hpu_kw} kW."}
        unit_price = float(price_row[0])
        model_code = f"{PU_VARIANT_PREFIX[req.hpu_variant]}-{kw_int}"
        equipment_label   = "Pumping Unit"
        item_product_type = "Pumping Unit"
        filename_infix    = "PU"
        template_name     = "PU_Offer_Template.docx"

    catalog_lph = float(flow_row[0]) if flow_row and flow_row[0] is not None \
                  else (req.fuel_lph or 0.0)
    qty = max(1, int(req.qty or 1))

    # ── 2. Enquiry ref (canonical ENCON pattern) ──────────────────────
    seq = next_quote_seq()
    auto_ref = build_enquiry_ref(seq, cust.technical or "", cust.location or "")

    # ── 3. Build form_data mirroring VLPH/Tundish shape ──────────────
    equipment_name = (
        f"{equipment_label} – {req.hpu_variant}, "
        f"{req.hpu_kw:g} kW @ {catalog_lph:g} LPH"
    )
    form_data = {
        "quote_seq": seq,
        "customer": {
            "company_name":     cust.company or "",
            "company_city":     cust.city or "",
            "company_state":    cust.state or "",
            "address":          ", ".join(filter(None, [
                                    cust.address or "",
                                    cust.city or "",
                                    cust.state or "",
                                    cust.pin or "",
                                ])),
            "poc_name":         _with_salutation(cust.salutation, cust.name),
            "poc_designation":  cust.designation or "",
            "mobile_no":        cust.phone or "",
            "email":            cust.email or "",
            "project_name":     cust.subject or equipment_name,
            "subject":          cust.subject or equipment_name,
            "ref_no":           auto_ref,
            "your_ref":         cust.ref_no or auto_ref,
            "enquiry_ref":      auto_ref,
            "marketing_person": _with_salutation(cust.marketing_salutation, cust.marketing),
            "marketing_phone":  cust.marketing_phone or "",
            "marketing_email":  cust.marketing_email or "",
            "technical_person": _with_salutation(cust.technical_salutation, cust.technical),
            "technical_phone":  cust.technical_phone or "",
            "technical_email":  cust.technical_email or "",
            "gstin":            cust.gstin or "",
            # Picked up by _build_equipment_name + by the HPU/PU template
            # via {{ hpu_variant }} / {{ hpu_kw }} / {{ hpu_lph }} /
            # {{ hpu_qty }} (the PU template reuses the hpu_* keys so a
            # single context dict serves both).
            "equipment_name_override": equipment_name,
            "hpu_variant":      req.hpu_variant,
            "hpu_kw":           f"{req.hpu_kw:g}",
            "hpu_lph":          f"{catalog_lph:g}",
            "hpu_model":        model_code,
            "hpu_qty":          str(qty),
            # Preheater-specific tech-data: blank so _strip_empty_tech_rows
            # drops the rows entirely.
            "ladle_tons":          "", "ladle_dim":           "",
            "ladle_drawing_no":    "", "refractory_weight_kg": "",
            "heating_schedule":    "", "heating_time":        "",
            "fuel_cv":             "", "fuel_consumption":    "",
            "fuel2_cv":            "", "fuel2_consumption":   "",
            "burner_model":        "", "blower_model":        "",
            "blower_size":         "", "blower_capacity":     "",
            "hydraulic_motor_hp":  "", "max_electrical_load": "",
            "fuel_name":           req.fuel_type or "",
            "burner_capacity_range": "",
            "pumping_unit":        f"{req.hpu_variant}, {req.hpu_kw:g} kW",
            "hood_movement":       "", "hood_type":           "",
            "pilot_gas_type":      "", "ignition_method":     "",
            "num_burners":         "",
            "max_fuel_consumption1": "", "max_fuel_consumption2": "",
            "is_oil":              True,
            "is_dual":             False,
            "control_mode":        "manual",
            "auto_control_type":   "",
            "control_valve_type":  "",
            "special_auto_ignition": False,
            "special_auto_controls": False,
            "vertical_qty":        qty,
            "horizontal_qty":      0,
            "nitrogen_purging":    False,
            "burner_kw_value":     "",
            "bom_items":           [],
            # Annexure IV — T&Cs sourced from Step 3 (mirrors VLPH).
            "tnc_prices":             cust.tnc_prices or "",
            "tnc_delivery":           cust.tnc_delivery or "",
            "tnc_gst":                cust.tnc_gst or "",
            "tnc_hsn_code":           cust.tnc_hsn_code or "",
            "tnc_pan_gst":            cust.tnc_pan_gst or "",
            "tnc_payment_terms":      cust.tnc_payment_terms or "",
            "tnc_packing_forwarding": cust.tnc_packing_forwarding or "",
            "tnc_freight":            cust.tnc_freight or "",
            "tnc_transit_insurance":  cust.tnc_transit_insurance or "",
            "tnc_validity":           cust.tnc_validity or "",
            "tnc_inspection":         cust.tnc_inspection or "",
            "tnc_guarantee":          cust.tnc_guarantee or "",
        },
        # Single-line price schedule. The product_type isn't 'Vertical
        # Ladle Preheater' / 'Horizontal Ladle Preheater' / 'Tundish'
        # so quote_writer's is_vertical / is_horizontal / is_tundish
        # flags stay False and the preheater scope sections are
        # suppressed; the total still pours into the vertical price
        # slot via the 'else' branch at quote_writer:692.
        "items": [{
            "product_type": item_product_type,
            "model":        f"{req.hpu_variant} {req.hpu_kw:g} kW",
            "description":  equipment_name,
            "qty":          qty,
            "unit_price":   unit_price,
        }],
        "valid_days": 30,
    }

    quote_data = calculate_quote(form_data)
    total_price = float(quote_data.get("grand_total") or unit_price * qty)

    # ── 4. Filename: {YYYY-MM-DD}_{Customer}_{HPU|PU}-{kW}kW-{variant}.docx ─
    _safe_company = "".join(ch for ch in (cust.company or "Client")
                            if ch.isalnum() or ch in " _-").strip().replace(" ", "_") or "Client"
    _safe_variant = req.hpu_variant.replace(" ", "")
    _date = datetime.now().strftime("%Y-%m-%d")
    filename = f"{_date}_{_safe_company}_{filename_infix}-{kw_int}kW-{_safe_variant}.docx"
    output_path = os.path.join(QUOTES_FOLDER, filename)

    template_path = os.path.join(BASE_DIR, template_name)
    generate_quote_docx(quote_data, output_path, template_path=template_path)

    return {
        "success":      True,
        "filename":     filename,
        "download_url": f"/api/download-quote/{filename}",
        "unit_price":   unit_price,
        "total_price":  total_price,
        "variant":      req.hpu_variant,
        "kw":           req.hpu_kw,
        "lph":          catalog_lph,
        "model_code":   model_code,
        "qty":          qty,
        "mode":         mode,
    }


@app.post("/api/generate-hpu-quote")
def generate_hpu_quote(req: HpuQuoteRequest):
    """Stand-alone HPU offer — pumping skid WITH in-built electric heater
    (model codes HPS / HPD / HPDD). Price = SUM(hpu_master) × 1.8."""
    try:
        return _generate_pumping_unit_offer(req, mode="hpu")
    except Exception as e:
        import traceback
        return {"error": str(e), "trace": traceback.format_exc()}


@app.post("/api/generate-pu-quote")
def generate_pu_quote(req: HpuQuoteRequest):
    """Stand-alone PU offer — pumping skid WITHOUT heater, for pre-heated
    oils (LDO / LSHS). Model codes PUS / PUD / PUDD. Price comes from
    pumping_unit_price.sell_price (already includes the markup)."""
    try:
        return _generate_pumping_unit_offer(req, mode="pu")
    except Exception as e:
        import traceback
        return {"error": str(e), "trace": traceback.format_exc()}


# ══════════════════════════════════════════════════════════════════════════
#  STAND-ALONE BLOWER & BURNER OFFERS
#  Catalog-pick offers modelled on the HPU/PU + Recuperator pattern:
#  template (Blower_/Burner_Offer_Template.docx, built by
#  build_blower_burner_templates.py) rendered self-contained via docxtpl.
# ══════════════════════════════════════════════════════════════════════════

@app.get("/blower", response_class=HTMLResponse)
def blower_costing_form():
    """Stand-alone Blower offer form (catalog pick from blower_master)."""
    html_path = os.path.join(BASE_DIR, "blower_costing.html")
    with open(html_path, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())


@app.get("/burner", response_class=HTMLResponse)
def burner_costing_form():
    """Stand-alone Burner offer form (catalog pick — ENCON oil/gas/dual + GAIL)."""
    html_path = os.path.join(BASE_DIR, "burner_costing.html")
    with open(html_path, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())


def _finite_price(v):
    """Coerce to a finite float > 0, or None. Guards against NaN/inf values
    (e.g. incomplete blower_master rows) that aren't JSON-serialisable and
    would 500 the response."""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    if f != f or f in (float("inf"), float("-inf")) or f <= 0:
        return None
    return f


def _fmt_num(v) -> str:
    """'10.0' -> '10', '7.5' -> '7.5', '2040.0' -> '2,040'."""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v or "")
    if f != f or f in (float("inf"), float("-inf")):  # NaN / inf -> blank
        return ""
    return f"{int(f):,}" if f == int(f) else f"{f:g}"


@app.get("/api/blower/catalog")
def blower_catalog():
    """ENCON blower models for the offer dropdown (price = with-motor)."""
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute(
        "SELECT model, hp, airflow, pressure, price_premium "
        "FROM blower_master ORDER BY pressure, CAST(hp AS REAL)"
    ).fetchall()
    conn.close()
    items = []
    for r in rows:
        price = _finite_price(r[4])
        if price is None:
            continue  # skip incomplete / NaN-priced rows (unsellable)
        items.append({
            "model":    r[0],
            "hp":       _fmt_num(r[1]),
            "airflow":  _fmt_num(r[2]),
            "pressure": (r[3] or "").strip(),
            "price":    price,
        })
    return {"items": items}


# Burner catalog config: group key -> (label, fuel text, pricelist section).
_BURNER_SECTIONS = {
    "oil":  ("ENCON Oil (Film) Burner", "Oil (LDO / HSD / SKO)",
             "PRICE FOR VARIOUS SIZES OF ENCON 'FILM' BURNER & ACCESSORIES"),
    "gas":  ("ENCON Gas Burner",        "Gas (NG / LPG / COG)",
             "PRICE FOR VARIOUS SIZES OF ENCON 'GAS' BURNER & ACCESSORIES"),
    "dual": ("ENCON Dual Fuel Burner",  "Dual Fuel (Gas + Oil)",
             "PRICE FOR VARIOUS SIZES OF ENCON DUAL FUEL BURNER & ACCESSORIES"),
}


def _burner_firing_capacity(conn, model: str) -> str:
    """Max firing rate (LPH at 24\" W.G.) for an ENCON burner size, if known."""
    row = conn.execute(
        "SELECT max_firing_lph FROM burner_selection_master "
        "WHERE model = ? AND pressure_wg = 24 LIMIT 1", (model,)
    ).fetchone()
    if row and row[0]:
        return f"up to {_fmt_num(row[0])} LPH"
    return ""


def _burner_lookup(group: str, model: str):
    """Return (price, fuel_text, capacity, label) for a burner pick, or None."""
    conn = sqlite3.connect(DB_PATH)
    try:
        if group == "gail":
            row = conn.execute(
                "SELECT burner_set FROM gail_gas_burner_master "
                "WHERE burner_size = ? AND burner_set IS NOT NULL LIMIT 1", (model,)
            ).fetchone()
            price = _finite_price(row[0]) if row else None
            if price is None:
                return None
            # Capacity is inherent in the GAIL model name (e.g. 'ENCON-500 KW').
            import re as _re
            m = _re.search(r"(\d[\d,]*)\s*KW", model, _re.I)
            cap = f"{m.group(1)} kW" if m else ""
            return price, "Natural Gas", cap, "GAIL Gas Burner"
        cfg = _BURNER_SECTIONS.get(group)
        if not cfg:
            return None
        label, fuel, section = cfg
        row = conn.execute(
            "SELECT price FROM burner_pricelist_master "
            "WHERE section = ? AND burner_size = ? AND component = 'BURNER SET' LIMIT 1",
            (section, model),
        ).fetchone()
        price = _finite_price(row[0]) if row else None
        if price is None:
            return None
        return price, fuel, _burner_firing_capacity(conn, model), label
    finally:
        conn.close()


@app.get("/api/burner/catalog")
def burner_catalog():
    """Burner models grouped by type for the offer dropdown."""
    conn = sqlite3.connect(DB_PATH)
    groups = []
    for key, (label, fuel, section) in _BURNER_SECTIONS.items():
        rows = conn.execute(
            "SELECT burner_size, price FROM burner_pricelist_master "
            "WHERE section = ? AND component = 'BURNER SET' ORDER BY burner_size",
            (section,),
        ).fetchall()
        items = []
        for r in rows:
            price = _finite_price(r[1])
            if price is None:
                continue
            items.append({"model": r[0], "price": price,
                          "capacity": _burner_firing_capacity(conn, r[0])})
        groups.append({"key": key, "label": label, "fuel": fuel, "items": items})
    # GAIL pre-assembled sets
    gail_rows = conn.execute(
        "SELECT burner_size, burner_set FROM gail_gas_burner_master "
        "WHERE burner_set IS NOT NULL AND section LIKE '%GAIL GAS BURNER%' ORDER BY burner_set"
    ).fetchall()
    import re as _re
    gail_items = []
    for sz, raw_price in gail_rows:
        price = _finite_price(raw_price)
        if price is None:
            continue
        m = _re.search(r"(\d[\d,]*)\s*KW", sz or "", _re.I)
        gail_items.append({"model": sz, "price": price,
                           "capacity": f"{m.group(1)} kW" if m else ""})
    groups.append({"key": "gail", "label": "GAIL Gas Burner",
                   "fuel": "Natural Gas", "items": gail_items})
    conn.close()
    return {"groups": groups}


class BlowerQuoteRequest(BaseModel):
    customer: HpuCustomer
    blower_model: str
    qty: int = 1


class BurnerQuoteRequest(BaseModel):
    customer: HpuCustomer
    burner_group: str = "gas"   # oil | gas | dual | gail
    burner_model: str
    qty: int = 1


def _generate_equipment_offer(cust: HpuCustomer, *, equipment_name: str,
                              specs: dict, unit_price: float, qty: int,
                              template_name: str, filename_infix: str,
                              drive_product: str) -> dict:
    """Shared minimal-offer generator for stand-alone equipment (blower /
    burner). Builds the docxtpl context from the customer block + the
    equipment specs, renders template_name, saves to quotes/, best-effort
    PDF + Drive upload. Returns the API response dict."""
    from datetime import datetime as _dt
    from docxtpl import DocxTemplate
    from engine.quote_writer import amount_in_words_indian, _format_inr

    qty = max(1, int(qty or 1))
    total_price = float(unit_price) * qty

    seq = next_quote_seq()
    full_ref = build_enquiry_ref(seq, cust.technical or "", cust.location or "")
    short_ref = full_ref.split(" DT.")[0]
    date_str = full_ref.split(" DT.")[-1] if " DT." in full_ref else _dt.now().strftime("%d/%m/%Y")

    company_address = ", ".join(filter(None, [
        (cust.address or "").strip(), (cust.city or "").strip(),
        (cust.state or "").strip(), (cust.pin or "").strip()]))

    ctx = {
        "project_name":      cust.subject or equipment_name,
        "subject":           cust.subject or f"Offer for {equipment_name}",
        "application":       equipment_name,
        "equipment_name":    equipment_name,
        "company_name":      cust.company or "",
        "company_address":   company_address,
        "email":             cust.email or "",
        "mobile_no":         cust.phone or "",
        "poc_name":          _with_salutation(cust.salutation, cust.name),
        "poc_designation":   cust.designation or "",
        "client_enq_ref":    "",
        "enquiry_ref":       full_ref,
        "enquiry_ref_short": short_ref,
        "enquiry_date_str":  date_str,
        "marketing_person":  _with_salutation(cust.marketing_salutation, cust.marketing),
        "marketing_email":   cust.marketing_email or "",
        "marketing_phone":   cust.marketing_phone or "",
        "technical_person":  _with_salutation(cust.technical_salutation, cust.technical) or _with_salutation(cust.marketing_salutation, cust.marketing),
        "technical_phone":   cust.technical_phone or cust.marketing_phone or "",
        "technical_email":   cust.technical_email or cust.marketing_email or "",
        # Price schedule (single line)
        "item_qty":          f"{qty:02d} No.",
        "unit_price":        _format_inr(unit_price),
        "total_price":       _format_inr(total_price),
        "grand_total":       _format_inr(total_price),
        "grand_total_in_words": f"INR. {amount_in_words_indian(total_price)} ONLY.",
        # T&C (Annexure renumbered to III)
        "tnc_prices":             cust.tnc_prices or "",
        "tnc_delivery":           cust.tnc_delivery or "",
        "tnc_gst":                cust.tnc_gst or "",
        "tnc_hsn_code":           cust.tnc_hsn_code or "",
        "tnc_pan_gst":            cust.tnc_pan_gst or "",
        "tnc_payment_terms":      cust.tnc_payment_terms or "",
        "tnc_packing_forwarding": cust.tnc_packing_forwarding or "",
        "tnc_freight":            cust.tnc_freight or "",
        "tnc_transit_insurance":  cust.tnc_transit_insurance or "",
        "tnc_validity":           cust.tnc_validity or "",
        "tnc_inspection":         cust.tnc_inspection or "",
        "tnc_guarantee":          cust.tnc_guarantee or "",
    }
    ctx.update(specs)

    tpl_path = os.path.join(BASE_DIR, template_name)
    tpl = DocxTemplate(tpl_path)
    tpl.render(ctx)

    safe_company = "".join(ch for ch in (cust.company or "Client")
                           if ch.isalnum() or ch in " _-").strip().replace(" ", "_") or "Client"
    docx_name = f"{filename_infix}_Offer_{safe_company}_{seq}.docx"
    docx_path = os.path.join(QUOTES_FOLDER, docx_name)
    tpl.save(docx_path)

    pdf_name = docx_name.replace(".docx", ".pdf")
    pdf_path = os.path.join(QUOTES_FOLDER, pdf_name)
    pdf_ok = _docx_to_pdf(docx_path, pdf_path)

    try:
        from engine.drive_uploader import upload_offer_async
        upload_offer_async(docx_path, docx_name, drive_product)
        if pdf_ok:
            upload_offer_async(pdf_path, pdf_name, drive_product)
    except Exception as _drv_err:
        print(f"WARN: drive upload kickoff failed: {_drv_err}")

    return {
        "success":      True,
        "filename":     docx_name,
        "pdf_filename": pdf_name if pdf_ok else None,
        "download_url": f"/api/download-quote/{docx_name}",
        "pdf_url":      f"/api/pdf-quote/{pdf_name}" if pdf_ok else None,
        "preview_url":  f"/api/preview-quote/{docx_name}",
        "quote_no":     full_ref,
        "enquiry_ref":  full_ref,
        "unit_price":   float(unit_price),
        "total_price":  total_price,
        "qty":          qty,
    }


@app.post("/api/generate-blower-quote")
def generate_blower_quote(req: BlowerQuoteRequest):
    """Stand-alone Blower offer — catalog pick from blower_master
    (price = price_premium, i.e. with motor)."""
    try:
        conn = sqlite3.connect(DB_PATH)
        row = conn.execute(
            "SELECT model, hp, airflow, pressure, price_premium "
            "FROM blower_master WHERE model = ? LIMIT 1", (req.blower_model,)
        ).fetchone()
        conn.close()
        if not row:
            return {"error": f"unknown blower model: {req.blower_model}"}
        model, hp, airflow, pressure, price = row
        unit_price = _finite_price(price)
        if unit_price is None:
            return {"error": f"blower '{model}' has no valid price in the catalog"}
        equipment_name = f"Centrifugal Blower – {model}"
        specs = {
            "blower_model":    model,
            "blower_hp":       _fmt_num(hp),
            "blower_airflow":  _fmt_num(airflow),
            "blower_pressure": (pressure or "").strip(),
        }
        result = _generate_equipment_offer(
            req.customer, equipment_name=equipment_name, specs=specs,
            unit_price=unit_price, qty=req.qty,
            template_name="Blower_Offer_Template.docx",
            filename_infix="Blower", drive_product="blower")
        result.update({"model": model, "config": f"{_fmt_num(hp)} HP • {(pressure or '').strip()}"})
        return result
    except Exception as e:
        import traceback
        return {"error": str(e), "trace": traceback.format_exc()}


@app.post("/api/generate-burner-quote")
def generate_burner_quote(req: BurnerQuoteRequest):
    """Stand-alone Burner offer — ENCON oil/gas/dual (burner_pricelist_master,
    BURNER SET) or GAIL pre-assembled set (gail_gas_burner_master)."""
    try:
        looked = _burner_lookup(req.burner_group, req.burner_model)
        if not looked:
            return {"error": f"unknown burner: {req.burner_group}/{req.burner_model}"}
        unit_price, fuel, capacity, label = looked
        equipment_name = f"{label} – {req.burner_model}"
        specs = {
            "burner_model":    req.burner_model,
            "burner_fuel":     fuel,
            "burner_capacity": capacity,
        }
        result = _generate_equipment_offer(
            req.customer, equipment_name=equipment_name, specs=specs,
            unit_price=unit_price, qty=req.qty,
            template_name="Burner_Offer_Template.docx",
            filename_infix="Burner", drive_product="burner")
        result.update({"model": req.burner_model,
                       "config": f"{fuel}{(' • ' + capacity) if capacity else ''}"})
        return result
    except Exception as e:
        import traceback
        return {"error": str(e), "trace": traceback.format_exc()}


# ══════════════════════════════════════════════════════════════════════════
#  COMBINED OFFER — merge several individually-generated offers into one .docx
# ══════════════════════════════════════════════════════════════════════════

@app.get("/combined", response_class=HTMLResponse)
@app.get("/combined-builder", response_class=HTMLResponse)
def combined_builder_page():
    """The Combined Offer builder: customer + T&C once, an input section per
    equipment, then one unified offer + a costing Excel. ('/combined-builder'
    is kept as an alias for older links/bookmarks.)"""
    html_path = os.path.join(BASE_DIR, "combined_builder.html")
    with open(html_path, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())


# ── Combined costing workbook: one sheet per equipment + a summary ──────────
class CombinedEquipment(BaseModel):
    name: str                          # e.g. "Vertical Ladle Preheater (10 T)"
    bom: List[dict] = []               # rows: {media?, item, ref?, qty, unit_price, total}
    total: Optional[float] = None      # equipment grand total (falls back to sum of rows)


class CombinedCostingRequest(BaseModel):
    project_name: Optional[str] = ""
    company_name: Optional[str] = ""
    equipments: List[CombinedEquipment] = []


@app.get("/api/download-xlsx/{filename}")
def download_xlsx(filename: str):
    """Serve a generated .xlsx from the quotes folder with the right type."""
    safe = os.path.basename(filename)
    file_path = os.path.join(QUOTES_FOLDER, safe)
    if not safe.lower().endswith(".xlsx") or not os.path.exists(file_path):
        return {"error": "File not found"}
    return FileResponse(
        path=file_path, filename=safe,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.post("/api/combined-costing-excel")
def combined_costing_excel(req: CombinedCostingRequest):
    """Build one .xlsx with a sheet per equipment (full itemised BOM) plus a
    Summary sheet totalling all equipment. Returns a download link."""
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

        if not req.equipments:
            return {"error": "No equipment supplied."}

        navy = "1A3A5C"
        hdr_fill = PatternFill("solid", fgColor=navy)
        hdr_font = Font(bold=True, color="FFFFFF")
        bold = Font(bold=True)
        thin = Side(style="thin", color="D0D7DE")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        money = '#,##0.00'

        wb = openpyxl.Workbook()

        # ── Summary sheet (first) ────────────────────────────────────────
        ws = wb.active
        ws.title = "Summary"
        ws["A1"] = "COMBINED COSTING SUMMARY"; ws["A1"].font = Font(bold=True, size=14, color=navy)
        ws["A2"] = f"Project: {req.project_name or '—'}"
        ws["A3"] = f"Client: {req.company_name or '—'}"
        r = 5
        for c, h in enumerate(["S. No.", "Equipment", "Total (Rs.)"], start=1):
            cell = ws.cell(r, c, h); cell.fill = hdr_fill; cell.font = hdr_font; cell.border = border
        grand = 0.0
        eq_totals = []
        for i, eq in enumerate(req.equipments, start=1):
            t = eq.total if eq.total is not None else sum(float(x.get("total") or 0) for x in eq.bom)
            t = float(t or 0); grand += t; eq_totals.append(t)
            ws.cell(r + i, 1, i).border = border
            ws.cell(r + i, 2, eq.name).border = border
            tc = ws.cell(r + i, 3, t); tc.number_format = money; tc.border = border
        gr = r + len(req.equipments) + 1
        ws.cell(gr, 2, "GRAND TOTAL").font = bold
        gc = ws.cell(gr, 3, grand); gc.font = bold; gc.number_format = money
        ws.column_dimensions["A"].width = 8; ws.column_dimensions["B"].width = 46; ws.column_dimensions["C"].width = 18

        # ── One sheet per equipment (full BOM) ───────────────────────────
        used = {"Summary"}
        for idx, eq in enumerate(req.equipments):
            base = "".join(ch for ch in eq.name if ch not in '[]:*?/\\').strip()[:28] or f"Equipment {idx+1}"
            name = base; n = 2
            while name in used:
                name = f"{base[:26]} {n}"; n += 1
            used.add(name)
            sh = wb.create_sheet(name)
            sh["A1"] = eq.name; sh["A1"].font = Font(bold=True, size=12, color=navy)
            hr = 3
            cols = ["S. No.", "Media", "Item", "Ref / Size", "Qty", "Unit Price", "Total"]
            for c, h in enumerate(cols, start=1):
                cell = sh.cell(hr, c, h); cell.fill = hdr_fill; cell.font = hdr_font; cell.border = border
                cell.alignment = Alignment(horizontal="center")
            sub = 0.0
            for j, row in enumerate(eq.bom, start=1):
                rr = hr + j
                tot = float(row.get("total") or 0); sub += tot
                vals = [j, row.get("media", ""), row.get("item", ""),
                        row.get("ref", row.get("size", "")),
                        row.get("qty", ""), float(row.get("unit_price") or 0), tot]
                for c, v in enumerate(vals, start=1):
                    cell = sh.cell(rr, c, v); cell.border = border
                    if c in (6, 7): cell.number_format = money
            tr = hr + len(eq.bom) + 1
            sh.cell(tr, 3, "TOTAL").font = bold
            tc = sh.cell(tr, 7, eq.total if eq.total is not None else sub)
            tc.font = bold; tc.number_format = money
            widths = [8, 16, 40, 18, 8, 14, 16]
            for c, w in enumerate(widths, start=1):
                sh.column_dimensions[openpyxl.utils.get_column_letter(c)].width = w

        safe_company = "".join(ch for ch in (req.company_name or "Client")
                               if ch.isalnum() or ch in " _-").strip().replace(" ", "_") or "Client"
        stamp = datetime.now().strftime("%d%b%Y_%H%M%S")
        out_name = f"Combined_Costing_{safe_company}_{stamp}.xlsx"
        out_path = os.path.join(QUOTES_FOLDER, out_name)
        wb.save(out_path)
        return {
            "success":      True,
            "filename":     out_name,
            "download_url": f"/api/download-xlsx/{out_name}",
            "sheets":       len(req.equipments) + 1,
            "grand_total":  grand,
        }
    except Exception as e:
        import traceback
        return {"error": str(e), "trace": traceback.format_exc()}


# ── Unified combined offer: one cover/customer/T&C, equipment sections ──────
class CombinedOfferEquipment(BaseModel):
    name: str
    specs: Optional[str] = ""          # one-line specs summary
    spec_rows: List[dict] = []         # [{label, value}] -> Technical Specifications
    qty: int = 1
    unit_price: float = 0
    bom: List[dict] = []               # itemised BOM -> Scope of Supply
    quote_payload: dict = {}           # full standalone QuoteRequest -> narrative Scope of Supply


class CombinedOfferRequest(BaseModel):
    customer: HpuCustomer
    equipments: List[CombinedOfferEquipment] = []
    project_name: Optional[str] = ""
    # Commercial adjustments applied once to the combined grand total.
    pf_pct: float = 0          # Packaging & Forwarding %
    design_pct: float = 0      # Designing %
    neg_pct: float = 0         # Negotiation %
    transport_amt: float = 0   # Transport (flat Rs.)


def _build_spec_comparison_table(docx_path, columns, rows):
    """Rewrite the rendered 'Parameter | Specification' placeholder table into a
    side-by-side comparison: parameters down the left column, one column per
    equipment with the equipment name as the heading. Modifies the placeholder
    in place so it keeps the offer's table style (borders, header shading)."""
    if not columns:
        return
    import copy as _copy
    from docx import Document as _Docx
    from docx.oxml import OxmlElement as _OE
    from docx.oxml.ns import qn as _qn

    doc = _Docx(docx_path)
    table = None
    for t in doc.tables:
        if t.rows and t.rows[0].cells and t.rows[0].cells[0].text.strip().lower() == "parameter":
            table = t
            break
    if table is None:
        return

    ncol = 1 + len(columns)
    grid = table._tbl.find(_qn("w:tblGrid"))
    # Extend to one column per equipment by cloning the last cell of each row
    # (so new cells inherit borders / the header shading).
    guard = 0
    while len(table.columns) < ncol and guard < 40:
        guard += 1
        if grid is not None:
            grid.append(_OE("w:gridCol"))
        for row in table.rows:
            tcs = row._tr.findall(_qn("w:tc"))
            new_tc = _copy.deepcopy(tcs[-1])
            for t_el in new_tc.iter(_qn("w:t")):
                t_el.text = ""
            row._tr.append(new_tc)

    hdr = table.rows[0].cells
    hdr[0].text = "Parameter"
    for j, name in enumerate(columns):
        if 1 + j < len(hdr):
            hdr[1 + j].text = str(name)
    for c in hdr:
        for p in c.paragraphs:
            for r in p.runs:
                r.bold = True

    for r in rows:
        cells = table.add_row().cells
        cells[0].text = str(r.get("param", ""))
        for j, v in enumerate(r.get("values", [])):
            if 1 + j < len(cells):
                cells[1 + j].text = str(v)

    table.autofit = True
    doc.save(docx_path)


def _build_narrative_scope_combined(combined_path, equipments, cust_base):
    """Replace the grouped Scope-of-Supply table in the combined offer with the
    full narrative scope of each equipment, reusing the standalone renderer
    (generate_quote_docx) per equipment and splicing in its scope section.
    Safe: any failure leaves the existing grouped scope untouched."""
    import copy as _copy
    from docx import Document as _Docx
    from docx.oxml import OxmlElement as _OE
    from docx.oxml.ns import qn as _qn
    from engine.quote_writer import generate_quote_docx as _gqd
    MARK = "►"

    def _retext(p_el, text):
        runs = p_el.findall(_qn('w:r'))
        if not runs:
            return
        first = runs[0]
        for t in first.findall(_qn('w:t')):
            first.remove(t)
        for r in runs[1:]:
            p_el.remove(r)
        t = _OE('w:t'); t.set(_qn('xml:space'), 'preserve'); t.text = text
        first.append(t)

    blocks = []
    for eq in equipments:
        qp = getattr(eq, "quote_payload", None) or {}
        if not qp:
            continue
        customer = dict(cust_base)
        for k, v in qp.items():
            if k not in ("bom_items", "items"):
                customer[k] = v
        customer["bom_items"] = qp.get("bom_items", [])
        qd = {"quote_seq": 0, "quote_no": "", "date": "", "customer": customer,
              "items": qp.get("items", [])}
        tmp = os.path.join(QUOTES_FOLDER, f"_scope_{abs(hash(eq.name)) % 999999}.docx")
        head_el, paras = None, []
        try:
            _gqd(qd, tmp)
            ed = _Docx(tmp)
            collecting = False
            for p in ed.paragraphs:
                u = (p.text or "").strip().upper()
                sty = (p.style.name if p.style else "") or ""
                if not collecting and u == "SCOPE OF SUPPLY":
                    head_el = _copy.deepcopy(p._p)
                    collecting = True
                    continue
                if collecting:
                    # Stop at the next annexure — including the formal
                    # "ANNEXURE I — SCOPE OF SUPPLY" heading that follows the
                    # body scope (it contains "SCOPE OF SUPPLY", so must NOT be
                    # excluded or we'd swallow it + its intro into every block).
                    if u.startswith("ANNEXURE"):
                        break
                    if not paras and not sty.startswith("Heading 3"):
                        continue   # skip the per-equipment intro line
                    paras.append(_copy.deepcopy(p._p))
        except Exception:
            paras = []
        finally:
            if os.path.exists(tmp):
                try: os.remove(tmp)
                except Exception: pass
        if paras and head_el is not None:
            blocks.append((eq.name, head_el, paras))

    if not blocks:
        return
    doc = _Docx(combined_path)
    target = None
    for t in doc.tables:
        full = "\n".join(c.text for r in t.rows for c in r.cells)
        if MARK in full or "Combustion Air Train" in full or "ENCON Supplied" in full:
            target = t
            break
    if target is None:
        return
    anchor = target._tbl
    for eqname, head_el, paras in blocks:
        h = _copy.deepcopy(head_el)
        _retext(h, (eqname or "Equipment").upper())
        anchor.addprevious(h)
        for pel in paras:
            anchor.addprevious(_copy.deepcopy(pel))
    anchor.getparent().remove(anchor)
    doc.save(combined_path)


@app.post("/api/generate-combined-offer")
def generate_combined_offer(req: CombinedOfferRequest):
    """Render Combined_Offer_Template.docx — one shared cover/customer/T&C,
    a technical row per equipment, and one combined price schedule."""
    try:
        from datetime import datetime as _dt
        from docxtpl import DocxTemplate
        from engine.quote_writer import amount_in_words_indian, _format_inr

        if not req.equipments:
            return {"error": "No equipment supplied."}
        cust = req.customer

        seq = next_quote_seq()
        full_ref = build_enquiry_ref(seq, cust.technical or "", cust.location or "")
        short_ref = full_ref.split(" DT.")[0]
        date_str = full_ref.split(" DT.")[-1] if " DT." in full_ref else _dt.now().strftime("%d/%m/%Y")
        company_address = ", ".join(filter(None, [
            (cust.address or "").strip(), (cust.city or "").strip(),
            (cust.state or "").strip(), (cust.pin or "").strip()]))

        price_lines, grand = [], 0.0
        for i, eq in enumerate(req.equipments, start=1):
            qty = max(1, int(eq.qty or 1))
            line_total = float(eq.unit_price or 0) * qty
            grand += line_total
            price_lines.append({
                "sno":        f"{i}.",
                "name":       eq.name,
                "qty":        f"{qty:02d} No.",
                "unit_price": _format_inr(eq.unit_price or 0),
                "total":      _format_inr(line_total),
            })

        # Scope of Supply (Annexure I) — modelled on the reference offer:
        # per equipment, the BOM is grouped by system (Combustion Air Train,
        # Gas Train, Nitrogen Train, …) with lettered sub-items + makes, then
        # the standard ENCON inclusions (Control Panel internals, I&C, safety)
        # are appended. docxtpl doesn't XML-escape free text reliably here, so
        # escape & < > ourselves (e.g. makes like L&T).
        import xml.sax.saxutils as _saxutils, string as _string
        def _xesc(s):
            return _saxutils.escape(str(s))

        def _system_name(media: str) -> str:
            m = (media or "").strip()
            mu = m.upper()
            mapping = {
                "COMB AIR": "Combustion Air Train",
                "ENCON ITEMS": "ENCON Supplied Equipment",
                "PURGING LINE": "Nitrogen Purging Train",
                "MISC ITEMS": "Miscellaneous Items",
            }
            if mu in mapping:
                return mapping[mu]
            if "PILOT LINE" in mu:
                return "Pilot Gas Train"
            if mu.endswith("LINE"):
                base = m[:-4].strip()          # keep fuel acronym case, e.g. "NG", "COG"
                return base + (" Train" if "GAS" in base.upper() else " Gas Train")
            return m.title() if m else "Items"

        # Standard inclusions present on every ENCON offer (not all are priced
        # BOM lines). Makes follow the reference offer.
        STANDARD_SCOPE = [
            ("Instrumentation & Control", [
                ("Thermocouples", "2 Nos", "Toshniwal / Tempsen / SBI"),
                ("Temperature Transmitters", "2 Nos", "Honeywell / ABB"),
            ]),
            ("Control Panel", [
                ("Isolation Switch", "1 No", "Siemens / ABB"),
                ("Emergency Stop", "1 No", "Standard"),
                ("Ammeter & Voltmeter", "1 Set", "Standard"),
                ("MCB / MPCB", "1 Lot", "Siemens / ABB"),
                ("Contactors & Relays", "1 Lot", "Siemens / ABB"),
                ("Temperature Indicator", "1 No", "Masibus"),
                ("PLC with HMI", "1 No", "Siemens"),
            ]),
            ("Safety Systems", [
                ("Gas Leakage Detection System", "1 Set", "Honeywell"),
                ("Limit Switches", "2 Nos", "BCH"),
            ]),
        ]
        _ltr = _string.ascii_lowercase

        def _subdesc(item, ref, qty, make):
            d = item
            if ref:
                d += f" — {ref}"
            extra = []
            if qty:
                extra.append(f"{qty} No.")
            if make:
                extra.append(make)
            if extra:
                d += "   ·   " + "  ·  ".join(str(x) for x in extra)
            return d

        scope_rows = []
        for eq in req.equipments:
            scope_rows.append({"sno": "", "desc": _xesc((eq.name or "Equipment").upper())})
            # group this equipment's BOM by media, preserving first-seen order
            groups, order = {}, []
            for b in (eq.bom or []):
                item = (b.get("item") or "").strip()
                if not item or item.upper() == "CONTROL PANEL":
                    continue  # control panel is expanded via standard inclusions
                media = (b.get("media") or "Items").strip() or "Items"
                if media not in groups:
                    groups[media] = []
                    order.append(media)
                groups[media].append(b)
            sysno = 0
            for media in order:
                sysno += 1
                scope_rows.append({"sno": f"{sysno}.", "desc": _xesc(_system_name(media))})
                for i, b in enumerate(groups[media]):
                    lbl = _ltr[i] if i < 26 else str(i + 1)
                    desc = _subdesc((b.get("item") or "").strip(),
                                    (b.get("ref") or b.get("size") or "").strip(),
                                    b.get("qty"), (b.get("make") or "").strip())
                    scope_rows.append({"sno": lbl, "desc": _xesc(desc)})
            for sysname, items in STANDARD_SCOPE:
                sysno += 1
                scope_rows.append({"sno": f"{sysno}.", "desc": _xesc(sysname)})
                for i, (it, q, mk) in enumerate(items):
                    scope_rows.append({"sno": _ltr[i], "desc": _xesc(f"{it}   ·   {q}  ·  {mk}")})

        # Technical Specifications — side by side: parameters down the left, one
        # column per equipment (equipment name = column heading). Built directly
        # with python-docx after rendering (dynamic column count), so plain text
        # here — no XML escaping needed.
        spec_columns = []          # equipment names (column headings)
        _eq_maps     = []          # per-equipment {param: value}
        _param_order = []          # union of parameter names, first-seen order
        for eq in req.equipments:
            spec_columns.append(eq.name or "Equipment")
            rows = eq.spec_rows or []
            if not rows and (eq.specs or "").strip():
                rows = [{"label": "Specifications", "value": eq.specs}]
            m = {}
            for s in rows:
                lbl = str((s or {}).get("label", "")).strip()
                if not lbl:
                    continue
                if lbl not in m:
                    m[lbl] = str((s or {}).get("value", "")).strip()
                if lbl not in _param_order:
                    _param_order.append(lbl)
            _eq_maps.append(m)
        spec_rows = [{"param": p, "values": [m.get(p, "") for m in _eq_maps]}
                     for p in _param_order]

        # Commercial adjustments on the combined grand total -> Final Total
        # (rounded to the nearest Rs.1000, matching the standalone forms).
        _pf   = grand * (req.pf_pct or 0) / 100
        _des  = grand * (req.design_pct or 0) / 100
        _neg  = grand * (req.neg_pct or 0) / 100
        _trn  = float(req.transport_amt or 0)
        _combined_final = round((grand + _pf + _des + _neg + _trn) / 1000) * 1000

        from engine.quote_writer import _supervision_rates
        _sup_mech, _sup_plc = _supervision_rates()

        ctx = {
            "project_name":      req.project_name or "Combined Equipment Offer",
            "subject":           cust.subject or "Offer for Combined Equipment",
            "application":       "Combined Equipment",
            "equipment_name":    "Combined Equipment Offer",
            "company_name":      cust.company or "",
            "company_address":   company_address,
            "email":             cust.email or "",
            "mobile_no":         cust.phone or "",
            "poc_name":          _with_salutation(cust.salutation, cust.name),
            "poc_designation":   cust.designation or "",
            "client_enq_ref":    "",
            "enquiry_ref":       full_ref,
            "enquiry_ref_short": short_ref,
            "enquiry_date_str":  date_str,
            "marketing_person":  _with_salutation(cust.marketing_salutation, cust.marketing),
            "marketing_email":   cust.marketing_email or "",
            "marketing_phone":   cust.marketing_phone or "",
            "technical_person":  _with_salutation(cust.technical_salutation, cust.technical) or _with_salutation(cust.marketing_salutation, cust.marketing),
            "technical_phone":   cust.technical_phone or cust.marketing_phone or "",
            "technical_email":   cust.technical_email or cust.marketing_email or "",
            # technical section + scope of supply + price schedule loops
            "equipments":   [{"name": e.name, "specs": e.specs or ""} for e in req.equipments],
            "scope_rows":   scope_rows,
            "price_lines":  price_lines,
            "grand_total":  _format_inr(grand),
            # Commercial adjustments applied once to the combined grand total.
            "pf_amount":         _format_inr(round(grand * (req.pf_pct or 0) / 100, 2)),
            "design_amount":     _format_inr(round(grand * (req.design_pct or 0) / 100, 2)),
            "neg_amount":        _format_inr(round(grand * (req.neg_pct or 0) / 100, 2)),
            "transport_amount":  _format_inr(float(req.transport_amt or 0)),
            "show_pf":           (grand * (req.pf_pct or 0) / 100) > 0,
            "show_design":       (grand * (req.design_pct or 0) / 100) > 0,
            "show_neg":          (grand * (req.neg_pct or 0) / 100) != 0,
            "show_transport":    float(req.transport_amt or 0) > 0,
            "final_total":       _format_inr(_combined_final),
            "grand_total_in_words": f"INR. {amount_in_words_indian(_combined_final)} ONLY.",
            # Supervision charges — pulled from the price master (component_price_master).
            "supervision_mech":  _sup_mech,
            "supervision_plc":   _sup_plc,
            # shared T&C
            "tnc_prices":             cust.tnc_prices or "",
            "tnc_delivery":           cust.tnc_delivery or "",
            "tnc_gst":                cust.tnc_gst or "",
            "tnc_hsn_code":           cust.tnc_hsn_code or "",
            "tnc_pan_gst":            cust.tnc_pan_gst or "",
            "tnc_payment_terms":      cust.tnc_payment_terms or "",
            "tnc_packing_forwarding": cust.tnc_packing_forwarding or "",
            "tnc_freight":            cust.tnc_freight or "",
            "tnc_transit_insurance":  cust.tnc_transit_insurance or "",
            "tnc_validity":           cust.tnc_validity or "",
            "tnc_inspection":         cust.tnc_inspection or "",
            "tnc_guarantee":          cust.tnc_guarantee or "",
        }

        tpl = DocxTemplate(os.path.join(BASE_DIR, "Combined_Offer_Template.docx"))
        tpl.render(ctx)
        safe_company = "".join(ch for ch in (cust.company or "Client")
                               if ch.isalnum() or ch in " _-").strip().replace(" ", "_") or "Client"
        docx_name = f"Combined_Offer_{safe_company}_{seq}.docx"
        docx_path = os.path.join(QUOTES_FOLDER, docx_name)
        tpl.save(docx_path)
        # Build the side-by-side technical-spec table (one column per equipment).
        _build_spec_comparison_table(docx_path, spec_columns, spec_rows)
        # Replace the grouped Scope of Supply with the full narrative scope per
        # equipment (falls back to the grouped scope if anything goes wrong).
        try:
            _build_narrative_scope_combined(
                docx_path, req.equipments,
                {"company_name": cust.company or "Client",
                 "project_name": req.project_name or "",
                 "subject": cust.subject or "", "address": "", "poc_name": ""})
        except Exception as _scope_err:
            print(f"WARN: narrative scope build failed, keeping grouped scope: {_scope_err}")

        pdf_name = docx_name.replace(".docx", ".pdf")
        pdf_ok = _docx_to_pdf(docx_path, os.path.join(QUOTES_FOLDER, pdf_name))
        try:
            from engine.drive_uploader import upload_offer_async
            upload_offer_async(docx_path, docx_name, "combined")
            if pdf_ok:
                upload_offer_async(os.path.join(QUOTES_FOLDER, pdf_name), pdf_name, "combined")
        except Exception as _drv_err:
            print(f"WARN: drive upload kickoff failed: {_drv_err}")

        return {
            "success":      True,
            "filename":     docx_name,
            "download_url": f"/api/download-quote/{docx_name}",
            "pdf_url":      f"/api/pdf-quote/{pdf_name}" if pdf_ok else None,
            "preview_url":  f"/api/preview-quote/{docx_name}",
            "quote_no":     full_ref,
            "grand_total":  grand,
            "count":        len(req.equipments),
        }
    except Exception as e:
        import traceback
        return {"error": str(e), "trace": traceback.format_exc()}


@app.get("/api/next-quote-ref")
def api_next_quote_ref(technical_person: str = "", location: str = ""):
    """Preview the auto-generated enquiry ref for the form.
    Reads (does not consume) the next sequence number."""
    seq = peek_quote_seq()
    return {"seq": seq, "ref": build_enquiry_ref(seq, technical_person, location)}


@app.post("/api/generate-quote")
async def generate_quote(req: QuoteRequest):
    try:
        from engine.quote_engine import calculate_quote
        from engine.quote_writer import generate_quote_docx
        from engine.pdf_writer import generate_quote_pdf

        seq = next_quote_seq()
        # Auto-generate enquiry reference: ET{YY}-{seq}-{initials}.
        # Overrides whatever the form sent so the reference always follows
        # the canonical ENCON pattern (ET26-001-JR for FY2026/Jyotirmoy Rabha).
        auto_ref = build_enquiry_ref(seq, req.technical_person or "", req.location or "")
        form_data = {
            "quote_seq": seq,
            "customer": {
                "company_name":    req.company_name,
                "company_city":    req.company_city,
                "company_state":   req.company_state,
                "address":         ", ".join(filter(None, [req.company_address, req.company_city, req.company_state, req.company_pin])),
                "poc_name":        _with_salutation(req.poc_salutation, req.poc_name),
                "poc_designation": req.poc_designation,
                "mobile_no":       req.mobile_no,
                "email":           req.email,
                "project_name":    req.project_name,
                "subject":         req.subject or req.project_name,
                "ref_no":          auto_ref,
                "your_ref":        auto_ref,
                "enquiry_ref":     auto_ref,
                "marketing_person": _with_salutation(req.marketing_salutation, req.marketing_person),
                "marketing_phone": req.marketing_phone,
                "marketing_email": req.marketing_email,
                "technical_person": _with_salutation(req.technical_salutation, req.technical_person),
                "technical_phone": req.technical_phone,
                "technical_email": req.technical_email,
                "gstin":           req.company_gstin,
                # Technical data (for template tech table)
                "ladle_tons":          req.ladle_tons,
                "ladle_dim":           req.ladle_dim,
                "ladle_drawing_no":    req.ladle_drawing_no,
                "refractory_weight_kg": req.refractory_weight_kg,
                "heating_schedule":    req.heating_schedule,
                "fuel_cv":             req.fuel_cv,
                "fuel_consumption":    req.fuel_consumption,
                "burner_model":        req.burner_model,
                "blower_model":        req.blower_model,
                "blower_size":         req.blower_size,
                "blower_capacity":     req.blower_capacity,
                "hydraulic_motor_hp":  req.hydraulic_motor_hp,
                "max_electrical_load": req.max_electrical_load,
                "total_in_words":      req.total_in_words,
                "heating_time":        req.heating_time,
                "fuel_name":           req.fuel_name,
                "burner_capacity_range": req.burner_capacity_range,
                "pumping_unit":        req.pumping_unit,
                "hood_movement":       req.hood_movement,
                "hood_type":           req.hood_type or "up_down",
                "pilot_gas_type":      req.pilot_gas_type or "LPG",
                "ignition_method":     req.ignition_method,
                "num_burners":         req.num_burners,
                "fuel2_cv":            req.fuel2_cv,
                "fuel2_consumption":   req.fuel2_consumption,
                "max_fuel_consumption1": req.max_fuel_consumption1,
                "max_fuel_consumption2": req.max_fuel_consumption2,
                "is_oil":              bool(req.is_oil),
                "is_dual":             bool(req.is_dual),
                "control_mode":        req.control_mode or "automatic",
                "auto_control_type":   req.auto_control_type or "plc",
                "control_valve_type":  req.control_valve_type or "pneumatic",
                "special_auto_ignition": bool(req.special_auto_ignition),
                "special_auto_controls": bool(req.special_auto_controls),
                "vertical_qty":        int(req.vertical_qty or 1),
                "horizontal_qty":      int(req.horizontal_qty or 1),
                "nitrogen_purging":    (req.purging_line or "no").lower() == "yes",
                "hpu_variant":         req.hpu_variant or "Duplex 1",
                "burner_kw_value":     req.burner_kw_value or "",
                "bom_items":           req.bom_items or [],
                # Annexure IV — Terms & Conditions (editable per offer)
                "tnc_prices":             req.tnc_prices or "",
                "tnc_delivery":           req.tnc_delivery or "",
                "tnc_gst":                req.tnc_gst or "",
                "tnc_hsn_code":           req.tnc_hsn_code or "",
                "tnc_pan_gst":            req.tnc_pan_gst or "",
                "tnc_payment_terms":      req.tnc_payment_terms or "",
                "tnc_packing_forwarding": req.tnc_packing_forwarding or "",
                "tnc_freight":            req.tnc_freight or "",
                "tnc_transit_insurance":  req.tnc_transit_insurance or "",
                "tnc_validity":           req.tnc_validity or "",
                "tnc_inspection":         req.tnc_inspection or "",
                "tnc_guarantee":          req.tnc_guarantee or "",
            },
            "items": [item.dict() for item in req.items],
            "gst_percent": req.gst_percent,
            "freight": req.freight,
            "valid_days": req.valid_days,
        }

        quote_data = calculate_quote(form_data)
        # Filename format: {YYYY-MM-DD}_{Customer}_{Product}-{Capacity}T.docx
        # (e.g. '2026-05-30_MagnoSteel_VLPH-10T.docx'). Falls back gracefully
        # if fields are missing.
        _PRODUCT_SHORT = {
            "Vertical Ladle Preheater":   "VLPH",
            "Horizontal Ladle Preheater": "HLPH",
        }
        _first_pt = (req.items[0].product_type if req.items else "") or ""
        _product = _PRODUCT_SHORT.get(_first_pt) or (
            "Tundish" if "Tundish" in _first_pt
            else "Recup" if "Recuperator" in _first_pt
            else (_first_pt.replace(" ", "") or "Offer")
        )
        _safe_company = "".join(ch for ch in (req.company_name or "Client")
                                if ch.isalnum() or ch in " _-").strip().replace(" ", "_") or "Client"
        _date = datetime.now().strftime("%Y-%m-%d")
        _capacity = f"-{int(req.ladle_tons)}T" if req.ladle_tons else ""
        filename = f"{_date}_{_safe_company}_{_product}{_capacity}.docx"
        output_path = os.path.join(QUOTES_FOLDER, filename)
        generate_quote_docx(quote_data, output_path)

        # Build the PDF as a faithful render of the .docx via LibreOffice
        # so the PDF and Word offer are visually identical. Fall back to the
        # reportlab generator only if LibreOffice is missing or fails (e.g.
        # local Windows dev box without LibreOffice installed).
        pdf_path = os.path.splitext(output_path)[0] + ".pdf"
        try:
            if not _docx_to_pdf(output_path, pdf_path):
                generate_quote_pdf(quote_data, pdf_path)
        except Exception as _pdf_err:
            print(f"WARN: PDF generation failed for {filename}: {_pdf_err}")

        # Persist the raw quote_data so the PDF can be regenerated later
        # if the .pdf file gets wiped (e.g., container redeploy).
        try:
            import json as _json
            json_path = os.path.splitext(output_path)[0] + ".json"
            with open(json_path, "w", encoding="utf-8") as _jf:
                _json.dump(quote_data, _jf, default=str)
        except Exception as _json_err:
            print(f"WARN: quote-data persist failed for {filename}: {_json_err}")

        # Mirror the docx + pdf to Google Drive (process@encon.in) in a
        # background thread. Drive credentials live in env vars; failures
        # are logged but never break offer generation.
        try:
            from engine.drive_uploader import upload_offer_async
            _product = (req.items[0].product_type if req.items else "")
            upload_offer_async(output_path, filename, _product)
            pdf_filename = f"{os.path.splitext(filename)[0]}.pdf"
            upload_offer_async(pdf_path, pdf_filename, _product)
        except Exception as _drv_err:
            print(f"WARN: drive upload kickoff failed: {_drv_err}")

        # Persist to quotes_log so we can list/re-download past quotes later
        try:
            conn = sqlite3.connect(DB_PATH)
            conn.execute("""
                INSERT INTO quotes_log (
                    quote_no, ref_no, company_name, poc_name, email, mobile_no,
                    project_name, equipment_type, ladle_tons, grand_total,
                    marketing_person, technical_person, file_path
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (
                quote_data["quote_no"],
                req.ref_no or "",
                req.company_name or "",
                req.poc_name or "",
                req.email or "",
                req.mobile_no or "",
                req.project_name or "",
                "VLPH" if not req.is_oil else "VLPH (Oil)",  # quick flag
                float(req.ladle_tons or 0),
                float(quote_data["grand_total"]),
                req.marketing_person or "",
                req.technical_person or "",
                output_path,
            ))
            conn.commit()
            conn.close()
        except Exception as log_err:
            print(f"WARN: quotes_log insert failed: {log_err}")

        return {
            "success": True,
            "quote_no": quote_data["quote_no"],
            "download_url": f"/api/download-quote/{filename}",
            "preview_url":  f"/api/preview-quote/{filename}",
            "pdf_url":      f"/api/pdf-quote/{filename}",
            "summary": {
                "subtotal": quote_data["subtotal"],
                "total":    quote_data["grand_total"],
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
    fuel_name = calc.get("fuel1_name", "")
    fuel_type = calc.get("fuel1_type", "ng")
    is_oil = fuel_type in {"hsd", "ldo", "hdo", "fo", "sko", "cfo", "lshs"}
    cv_unit = "kcal/kg" if is_oil else "kcal/Nm³"
    flow_unit = "kg" if is_oil else "Nm³"
    rate_unit = "kg/hr" if is_oil else "Nm³/hr"
    params = [
        ("Initial Temp (Ti)",       f"{calc.get('Ti','')} °C"),
        ("Final Temp (Tf)",         f"{calc.get('Tf','')} °C"),
        ("Refractory Weight",       f"{calc.get('refractory_weight','')} kg"),
        ("Fuel",                    fuel_name),
        ("Fuel CV",                 f"{calc.get('fuel_cv','')} {cv_unit}"),
        ("Time Taken",              f"{calc.get('time_taken_hr','')} hr"),
        ("Heat Load",               f"{calc.get('heat_load_kcal','')} kcal"),
        ("Firing Rate",             f"{calc.get('firing_rate_kcal','')} kcal/hr"),
        ("Fuel Consumption",        f"{calc.get('fuel_consumption_nm3','')} {flow_unit}"),
        ("Calc. Firing Rate",       f"{calc.get('calculated_firing_rate_nm3hr','')} {rate_unit}"),
        ("Design Firing Rate",      f"{calc.get('extra_firing_rate_nm3hr','')} {rate_unit}"),
        ("Equiv. Firing Rate (LPH)", f"{calc.get('equivalent_lph','')} ltr/hr"),
        ("Fuel Density",            f"{calc.get('fuel_density','')} {calc.get('fuel_density_unit','kg/ltr')}"),
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


def _soffice_binary():
    """Locate LibreOffice binary across Linux / macOS / Windows."""
    import shutil as _sh
    for cand in ("soffice", "libreoffice",
                 "/usr/bin/soffice", "/usr/bin/libreoffice",
                 r"C:\Program Files\LibreOffice\program\soffice.exe",
                 r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if os.path.exists(cand) or _sh.which(cand):
            return _sh.which(cand) or cand
    return None


def _docx_to_pdf(docx_path: str, pdf_path: str) -> bool:
    """Convert docx -> pdf via LibreOffice headless. Returns True on success.

    Each invocation gets its own UserInstallation profile so concurrent
    requests don't collide on the shared lock file.
    """
    soffice = _soffice_binary()
    if not soffice:
        return False

    import subprocess
    import tempfile
    import shutil as _sh

    out_dir = os.path.dirname(pdf_path) or "."
    profile_dir = tempfile.mkdtemp(prefix="lo_profile_")
    try:
        proc = subprocess.run(
            [
                soffice,
                "--headless",
                "--norestore",
                "--nolockcheck",
                f"-env:UserInstallation=file://{profile_dir}",
                "--convert-to", "pdf",
                "--outdir", out_dir,
                docx_path,
            ],
            capture_output=True, text=True, timeout=120,
        )
        if proc.returncode != 0:
            print(f"WARN: soffice convert failed (rc={proc.returncode}): "
                  f"stdout={proc.stdout!r} stderr={proc.stderr!r}")
            return False

        produced = os.path.join(
            out_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf"
        )
        if produced != pdf_path and os.path.exists(produced):
            _sh.move(produced, pdf_path)
        return os.path.exists(pdf_path)
    except Exception as e:
        print(f"WARN: soffice convert exception: {e}")
        return False
    finally:
        _sh.rmtree(profile_dir, ignore_errors=True)


@app.get("/api/pdf-quote/{filename}")
def pdf_quote(filename: str):
    """Serve the PDF version of a generated offer.

    Each call to /api/generate-quote now writes both a .docx (Word
    template) and a .pdf (built directly with reportlab) into
    QUOTES_FOLDER. This endpoint just serves the pre-generated PDF
    so the client doesn't need any docx->pdf conversion at request
    time. If the PDF doesn't exist (older quote or generator error),
    it builds it on demand from the docx-equivalent quote_data.
    """
    base = os.path.splitext(filename)[0]
    pdf_dst = os.path.join(QUOTES_FOLDER, f"{base}.pdf")

    if os.path.exists(pdf_dst):
        return FileResponse(path=pdf_dst, filename=f"{base}.pdf",
            media_type="application/pdf")

    # Prefer reconverting the existing .docx via LibreOffice so the PDF is
    # a faithful copy of the Word offer.
    docx_path = os.path.join(QUOTES_FOLDER, f"{base}.docx")
    if os.path.exists(docx_path):
        if _docx_to_pdf(docx_path, pdf_dst):
            return FileResponse(path=pdf_dst, filename=f"{base}.pdf",
                media_type="application/pdf")

    # Last-resort fallback: rebuild from persisted quote JSON via reportlab.
    json_path = os.path.join(QUOTES_FOLDER, f"{base}.json")
    if os.path.exists(json_path):
        try:
            import json as _json
            from engine.pdf_writer import generate_quote_pdf
            with open(json_path, encoding="utf-8") as f:
                quote_data = _json.load(f)
            generate_quote_pdf(quote_data, pdf_dst)
            return FileResponse(path=pdf_dst, filename=f"{base}.pdf",
                media_type="application/pdf")
        except Exception as e:
            return {"error": f"PDF build failed: {e}"}

    return {"error": "PDF not available — regenerate the quote to produce it."}


def _stamp_page_numbers(pdf_path: str):
    """Overlay 'Page X / Y' at the bottom-right of every page."""
    from pypdf import PdfReader, PdfWriter
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
    import io

    reader = PdfReader(pdf_path)
    total = len(reader.pages)
    writer = PdfWriter()

    for i, page in enumerate(reader.pages, start=1):
        # Build a single-page overlay with the label
        box = page.mediabox
        width  = float(box.width)
        height = float(box.height)
        packet = io.BytesIO()
        c = canvas.Canvas(packet, pagesize=(width, height))
        c.setFont("Helvetica-Bold", 10)
        c.setFillColorRGB(0.10, 0.22, 0.36)   # dark navy
        c.drawRightString(width - 36, 20, f"Page {i} / {total}")
        c.save()
        packet.seek(0)
        stamp = PdfReader(packet).pages[0]
        page.merge_page(stamp)
        writer.add_page(page)

    with open(pdf_path, "wb") as f:
        writer.write(f)


@app.get("/api/preview-quote/{filename}")
def preview_quote(filename: str):
    """Render a generated offer .docx as HTML so it can be shown in the browser."""
    import mammoth
    file_path = os.path.join(QUOTES_FOLDER, filename)
    if not os.path.exists(file_path):
        return HTMLResponse("<p style='color:red'>File not found</p>", status_code=404)
    with open(file_path, "rb") as f:
        result = mammoth.convert_to_html(f)
    html = f"""<!doctype html><html><head><meta charset="utf-8">
<style>
  body{{font-family:'Calibri','Segoe UI',Arial,sans-serif;font-size:11pt;color:#222;max-width:900px;margin:24px auto;padding:0 28px;line-height:1.4;}}
  h1,h2,h3,h4{{color:#1a3a5c;margin:14px 0 6px;}}
  table{{border-collapse:collapse;margin:10px 0;width:100%;}}
  table td,table th{{border:1px solid #bbb;padding:4px 8px;font-size:10.5pt;vertical-align:top;}}
  table th{{background:#eef2f7;}}
  p{{margin:4px 0;}}
  img{{max-width:100%;}}
  a{{color:#1d4ed8;}}
</style></head><body>{result.value}</body></html>"""
    return HTMLResponse(html)


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