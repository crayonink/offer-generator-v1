# main app
import streamlit as st
import io
import pandas as pd

# -----------------------------
# Internal modules
# -----------------------------
from config import MARKUP
from calculations.burner import BurnerInputs, calculate_burner
from calculations.pipes import PipeInputs, calculate_pipe_sizes
from bom.vlph_builder import build_vlph_120t_df
from summary.cost_summary import build_cost_summary_df
from export.excel_writer import write_excel
from export.word_writer import generate_word_offer


# -------------------------------------------------
# SESSION STATE
# -------------------------------------------------
if "excel_buffer" not in st.session_state:
    st.session_state.excel_buffer = None

if "word_buffer" not in st.session_state:
    st.session_state.word_buffer = None


# -------------------------------------------------
# PAGE CONFIG
# -------------------------------------------------
st.set_page_config(page_title="Offer Generator", layout="centered")
st.title("Offer Generator")


# -------------------------------------------------
# STEP 1: COMMERCIAL DETAILS
# -------------------------------------------------
st.subheader("Step 1: Customer & Organization Details")

company_name = st.text_input("Company Name")
company_address = st.text_area("Company Address")

project_name = st.selectbox(
    "Project",
    ["Ladle Preheater", "Preheater Hoods", "Dryers"]
)

fuel_type = st.selectbox("Fuel Type", ["Oil", "Gas"])

poc_designation = st.text_input("Point of Contact (Designation)")
poc_name = st.text_input("POC Name")
mobile_no = st.text_input("Mobile Number")

st.divider()


# -------------------------------------------------
# STEP 2: BURNER INPUTS
# -------------------------------------------------
st.subheader("Step 2: Burner Size Calculation – Inputs")

input_df = pd.DataFrame({
    "Parameter": [
        "Ti",
        "Tf",
        "Actual Refractory Weight",
        "MG Fuel CV",
        "Time Taken"
    ],
    "Value": [
        650.0,
        1200.0,
        21500.0,
        8500.0,
        1.0
    ],
    "Unit": ["°C", "°C", "Kg", "Kcal/Nm³", "Hours"]
})

edited_df = st.data_editor(
    input_df,
    hide_index=True,
    num_rows="fixed",
    use_container_width=True
)

values = dict(zip(edited_df["Parameter"], edited_df["Value"]))

Ti = values["Ti"]
Tf = values["Tf"]
refractory_weight = values["Actual Refractory Weight"]
fuel_cv = values["MG Fuel CV"]
time_taken_hr = values["Time Taken"]

st.divider()


# -------------------------------------------------
# GENERATE FILES
# -------------------------------------------------
if st.button("Generate Offer & Calculation Excel"):

    # -----------------------------
    # Validation
    # -----------------------------
    if not company_name or not company_address or not poc_name or not mobile_no:
        st.error("Please fill all mandatory commercial fields")
        st.stop()

    # -----------------------------
    # BURNER CALCULATION
    # -----------------------------
    burner_inputs = BurnerInputs(
        Ti=Ti,
        Tf=Tf,
        refractory_weight=refractory_weight,
        fuel_cv=fuel_cv,
        time_taken_hr=time_taken_hr,
    )

    burner_results = calculate_burner(burner_inputs)

    # -----------------------------
    # PIPE CALCULATION
    # -----------------------------
    pipe_inputs = PipeInputs(
        ng_flow_nm3hr=burner_results.extra_firing_rate_nm3hr,
        air_flow_nm3hr=burner_results.air_qty_nm3hr,
    )

    pipe_results = calculate_pipe_sizes(pipe_inputs)

    # -----------------------------
    # BOM
    # -----------------------------
    vlph_df = build_vlph_120t_df(
        burner_results=burner_results,
        pipe_results=pipe_results,
    )

    # -----------------------------
    # ✅ COST SUMMARY (LEGACY SAFE)
    # -----------------------------
    bought_out_cost = vlph_df.loc[
        vlph_df["MEDIA"] != "ENCON ITEMS", "TOTAL"
    ].sum()

    inhouse_sell = vlph_df.loc[
        vlph_df["MEDIA"] == "ENCON ITEMS", "TOTAL"
    ].sum()

    bought_out_sell = bought_out_cost * MARKUP
    inhouse_cost = inhouse_sell / MARKUP

    cost_summary_df = build_cost_summary_df(
        bought_out_cost=bought_out_cost,
        bought_out_sell=bought_out_sell,
        inhouse_cost=inhouse_cost,
        inhouse_sell=inhouse_sell,
    )

    # -----------------------------
    # WRITE EXCEL
    # -----------------------------
    st.session_state.excel_buffer = io.BytesIO()

    write_excel(
        buffer=st.session_state.excel_buffer,
        sheets={
            "VLPH-120T": vlph_df,
            "Cost Summary": cost_summary_df,
        },
        burner_inputs=burner_inputs,
        burner_results=burner_results,
        pipe_results=pipe_results,
    )

    st.session_state.excel_buffer.seek(0)

    # -----------------------------
    # WORD OFFER
    # -----------------------------
    offer_context = {
        "company_name": company_name,
        "company_address": company_address,
        "project_name": project_name,
        "fuel_type": fuel_type,
        "poc_name": poc_name,
        "poc_designation": poc_designation,
        "mobile_no": mobile_no,
    }

    st.session_state.word_buffer = generate_word_offer(
        template_path="Offer_Template.docx",
        context=offer_context,
    )


# -------------------------------------------------
# DOWNLOAD BUTTONS
# -------------------------------------------------
if st.session_state.excel_buffer:
    st.download_button(
        "⬇ Download Costing Excel",
        st.session_state.excel_buffer,
        file_name="Costing.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

if st.session_state.word_buffer:
    st.download_button(
        "⬇ Download Word Offer",
        st.session_state.word_buffer,
        file_name="Final_Offer.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
