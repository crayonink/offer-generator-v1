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


def select_blower(required_hp: float, pressure: int = 40) -> dict:
    """
    Select combustion air blower from database
    based on required HP and pressure.
    """

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.execute("""
        SELECT model, hp, airflow, cfm, pressure, price_basic, price_premium
        FROM blower_master
        WHERE pressure = ?
        AND hp >= ?
        ORDER BY hp ASC
        LIMIT 1
    """, (pressure, required_hp))

    row = cursor.fetchone()
    conn.close()

    if not row:
        raise ValueError(
            f"No blower found for HP >= {required_hp} at pressure {pressure}"
        )

    return {
        "model": row[0],
        "hp": row[1],
        "airflow_nm3hr": row[2],
        "cfm": row[3],
        "pressure": row[4],
        "price_basic": row[5],
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