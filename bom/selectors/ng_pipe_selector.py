def select_ng_pipe(ng_flow_nm3hr: float) -> dict:
    if ng_flow_nm3hr <= 250:
        return {"size_nb": 50}
    elif ng_flow_nm3hr <= 600:
        return {"size_nb": 80}
    else:
        return {"size_nb": 100}
