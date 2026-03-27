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
            "Blower": q("SELECT DISTINCT model as id, model as label, section, price_without_motor, price__with_motor, hp FROM blower_pricelist_master WHERE price_without_motor IS NOT NULL ORDER BY section, hp"),
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
            col = "price__with_motor" if with_motor else "price_without_motor"
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
        from bom.regen_builder import build_regen_df

        result = calculate_regen(RegenInputs(
            material_weight_kg=req.material_weight_kg,
            Ti=req.Ti,
            Tf=req.Tf,
            Cp=req.Cp,
            cycle_time_hr=req.cycle_time_hr,
            efficiency=req.efficiency,
            num_pairs_override=req.num_pairs_override,
        ))

        bom_df = build_regen_df(result.num_pairs, req.markup)

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
                "total_kw": result.total_kw,
            },
            "bom": bom_df.to_dict(orient="records"),
            "cost_summary": {
                "total_cost": round(total_cost, 2),
                "total_selling": round(total_selling, 2),
                "markup": req.markup,
            },
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
    return {
        "covered":  sorted(needed & have),
        "missing":  sorted(needed - have),
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