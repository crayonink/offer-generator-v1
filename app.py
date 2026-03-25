# -------------------------------------------------
# MAIN APP
# -------------------------------------------------

import streamlit as st
import io
import pandas as pd
import re
import os
import pickle

from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google_auth_oauthlib.flow import InstalledAppFlow

from config import MARKUP

# VLHP calculation
from calculations.burner import BurnerInputs, calculate_burner
from calculations.pipes import PipeInputs, calculate_pipe_sizes

# REGEN calculation
from calculations.regen_burner import (
    RegenBurnerInputs,
    calculate_regen_burner,
    AVAILABLE_KW
)

# BOM builders
from bom.vlph_builder import build_vlph_120t_df

# NEW PIPELINE
from engine.run_pipeline import run_pipeline
from bom.selectors.selection_engine import SystemType

# summaries & export
from summary.cost_summary import build_cost_summary_df
from export.excel_writer import write_excel
from export.word_writer import generate_word_offer


# -------------------------------------------------
# CONFIG
# -------------------------------------------------

FOLDER_ID = "1HKNWIisZzO03CE_WMLuqg17zjMxpd81M"

FUEL_CV_MAP = {
    "LDO": 10200,
    "Furnace Oil": 10000,
    "Diesel": 10200,
    "LPG": 11000,
    "Natural Gas": 8500,
    "Furnace Gas": 1000,
    "Producer Gas": 1200,
    "Coke Oven Gas": 4300,
    "Mixed Gas": 2500,
    "Miscellaneous Gas": 2000,
}


# -------------------------------------------------
# GOOGLE DRIVE UPLOAD
# -------------------------------------------------

def upload_excel_to_google_sheets(buffer, filename):
    SCOPES = ["https://www.googleapis.com/auth/drive.file"]

    creds = None

    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)

    if not creds:
        flow = InstalledAppFlow.from_client_secrets_file(
            "client_secret.json", SCOPES
        )
        creds = flow.run_local_server(port=0)

        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)

    drive_service = build("drive", "v3", credentials=creds)

    buffer.seek(0)

    media = MediaIoBaseUpload(
        buffer,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    file = drive_service.files().create(
        body={"name": f"{filename}.xlsx", "parents": [FOLDER_ID]},
        media_body=media,
        fields="id"
    ).execute()

    return f"https://docs.google.com/spreadsheets/d/{file['id']}"


# -------------------------------------------------
# SESSION STATE
# -------------------------------------------------

def init_session_state():
    if "excel_buffer" not in st.session_state:
        st.session_state.excel_buffer = None
    if "word_buffer" not in st.session_state:
        st.session_state.word_buffer = None
    if "regen_results" not in st.session_state:
        st.session_state.regen_results = None


# -------------------------------------------------
# INPUT SECTIONS (UNCHANGED)
# -------------------------------------------------

def render_commercial_inputs():
    st.subheader("Step 1: Customer & Organization Details")

    company_name = st.text_input("Company Name")
    company_address = st.text_area("Company Address")

    project_name = st.selectbox(
        "Project",
        ["Ladle Preheater", "Preheater Hoods", "Dryers"]
    )

    fuel_category = st.selectbox("Fuel Category", ["Oil", "Gas"])

    if fuel_category == "Oil":
        fuel_type = st.selectbox("Fuel Type", ["LDO", "Furnace Oil", "Diesel"])
    else:
        fuel_type = st.selectbox(
            "Fuel Type",
            list(FUEL_CV_MAP.keys()),
        )

    fuel_cv_default = FUEL_CV_MAP.get(fuel_type, 8500)

    poc_designation = st.text_input("Point of Contact (Designation)")
    poc_name = st.text_input("POC Name")
    mobile_no = st.text_input("Mobile Number")
    email = st.text_input("Email")

    return locals()


def render_equipment_inputs():
    st.subheader("Step 2: Equipment Configuration")

    equipment_type = st.selectbox(
        "Equipment",
        ["Ladle", "Tundish", "AOD Vessel", "Launder"]
    )

    num_burners = st.number_input("Number of Burners", min_value=1, value=2)

    burner_type = st.selectbox(
        "Burner Type",
        ["Conventional", "Oxyfuel", "Oxy-enriched", "Regenerative"]
    )

    control_type = st.selectbox("Control System", ["Automatic", "Manual"])

    return locals()


def render_regen_inputs(fuel_cv_default):
    st.subheader("Step 3: Regen Burner Inputs")

    power_kw = st.selectbox("Burner Power (kW)", AVAILABLE_KW)
    num_burners = st.number_input("Number of Burner Pairs", min_value=1, value=2)

    fuel_type = st.text_input("Fuel Type", value="Natural Gas")
    fuel_cv = st.number_input("Fuel CV", value=int(fuel_cv_default))

    return {
        "power_kw": power_kw,
        "num_burners": num_burners,
        "fuel_type": fuel_type,
        "fuel_cv": fuel_cv,
    }


# -------------------------------------------------
# REGEN OFFER (UPDATED)
# -------------------------------------------------

def generate_regen_offer(commercial, equipment, regen_values):

    # STEP 1: Existing engineering calc (to get flows)
    regen_results = calculate_regen_burner(
        RegenBurnerInputs(**regen_values)
    )

    # STEP 2: Use flows in new pipeline
    result = run_pipeline(
    system_type=SystemType.REGEN,
    capacity_kw=regen_values["power_kw"],
    ng_flow_nm3hr = regen_results.ng_flow_nm3hr,
    air_flow_nm3hr=regen_results.air_flow_nm3hr/ regen_results.num_burners,
)

    bom_df = result["bom"]

    # STEP 3: Export
    buffer = io.BytesIO()
    write_excel(buffer, {"REGEN": bom_df}, None, None, None)
    buffer.seek(0)

    return buffer, regen_results, result


# -------------------------------------------------
# MAIN
# -------------------------------------------------

st.set_page_config(page_title="Offer Generator", layout="centered")
st.title("Offer Generator")

init_session_state()

commercial = render_commercial_inputs()
st.divider()

equipment = render_equipment_inputs()
st.divider()

is_regen = equipment["burner_type"] == "Regenerative"

if is_regen:
    regen_values = render_regen_inputs(commercial["fuel_cv_default"])

st.divider()


if st.button("Generate Offer & Calculation Excel"):

    if is_regen:
        excel_buffer, regen_results, result = generate_regen_offer(
            commercial, equipment, regen_values
        )

        st.session_state.regen_results = regen_results

        # 👇 NEW: Show total cost
        st.metric("Total Cost (₹)", f"{result['total_cost']:,.0f}")

    st.session_state.excel_buffer = excel_buffer

    try:
        sheet_url = upload_excel_to_google_sheets(
            excel_buffer,
            f"{commercial['company_name']}_Costing"
        )
        st.success(f"Google Sheet created: {sheet_url}")
    except Exception as e:
        st.error(f"Upload failed: {e}")

    st.session_state.word_buffer = generate_word_offer(
        template_path="Offer_Template.docx",
        context={**commercial, **equipment},
    )


if st.session_state.excel_buffer:
    st.download_button("Download Excel", st.session_state.excel_buffer)

if st.session_state.word_buffer:
    st.download_button("Download Word", st.session_state.word_buffer)

    st.write(regen_results)