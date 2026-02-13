# bom/vlph_rules.py

def vlph_rules(firing_rate_nm3_hr):
    """
    Rules extracted directly from your BOM sheets
    """

    if firing_rate_nm3_hr <= 450:
        return {
            "burner": "ENCON MG BURNER WITH B. BLOCK",
            "burner_ref": f"NATURAL GAS FLOW: {round(firing_rate_nm3_hr)} Nm3/hr",
            "ng_gas_train": "NG GAS TRAIN – FLOW: 400 Nm3/hr",
            "ng_pipe_nb": 125,
            "air_pipe_nb": 300,
            "air_flow_nm3_hr": 4000,
            "blower": '25 HP, 28" WC, 5100 Nm3/hr'
        }

    raise ValueError("VLPH rules not defined for this firing rate")
