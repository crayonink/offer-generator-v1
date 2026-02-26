def select_air_duct(air_flow_nm3hr: float) -> dict:
    """
    Select combustion air duct size based on air flow.
    """

    if air_flow_nm3hr <= 2000:
        nb = 200
    elif air_flow_nm3hr <= 3500:
        nb = 250
    elif air_flow_nm3hr <= 5000:
        nb = 300
    else:
        nb = 350

    return {
        "nb": nb
    }
