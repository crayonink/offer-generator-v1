import sqlite3


DB_PATH = "vlph.db"


def calculate_blower_hp(firing_rate: float) -> float:
    """
    Calculate required blower HP based on firing rate.
    Formula from engineering sheet.
    """

    heat_load = firing_rate * 8600 * 118 / 100000
    air_requirement = heat_load / 1.7
    blower_hp = air_requirement * 40 / 3200

    return blower_hp


def select_blower(required_hp: float, series: str = "28") -> dict:
    """
    Select combustion air blower from blower_master.
    Filters by ENCON series (28" WG or 40" WG) and picks the smallest HP >= required.
    hp and numeric columns are stored as text in the DB — cast at query time.
    """

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.execute("""
        SELECT model,
               CAST(hp         AS REAL) AS hp_val,
               CAST(airflow    AS REAL) AS airflow_val,
               CAST(cfm        AS REAL) AS cfm_val,
               pressure,
               CAST(price_basic    AS REAL) AS price_basic_val,
               CAST(price_premium  AS REAL) AS price_premium_val
        FROM blower_master
        WHERE model LIKE ?
          AND CAST(hp AS REAL) >= ?
        ORDER BY CAST(hp AS REAL) ASC
        LIMIT 1
    """, (f"ENCON {series}/%", required_hp))

    row = cursor.fetchone()
    conn.close()

    if not row:
        raise ValueError(
            f"No ENCON {series}\" WG blower found for HP >= {required_hp:.2f}"
        )

    return {
        "model":         row[0],
        "hp":            row[1],
        "airflow_nm3hr": row[2],
        "cfm":           row[3],
        "pressure":      row[4],
        "price_basic":   row[5],
        "price_premium": row[6],
    }


def select_blower_from_firing_rate(firing_rate: float) -> dict:
    """
    Full pipeline:
    firing rate → required HP → blower selection
    """

    required_hp = calculate_blower_hp(firing_rate)

    blower = select_blower(required_hp)

    blower["required_hp"] = round(required_hp, 2)

    return blower


# -------------------------------------------------
# TEST RUN
# -------------------------------------------------
if __name__ == "__main__":

    firing_rate = 70

    result = select_blower_from_firing_rate(firing_rate)

    print("Blower Selection Result")
    print(result)