import sqlite3

# Connect to database (creates file if not exists)
conn = sqlite3.connect("vlph.db")

# Create cursor FIRST
cursor = conn.cursor()

# ----------------------------
# AGR TABLE
# ----------------------------
# ----------------------------
# AGR TABLE
# ----------------------------
cursor.execute("DROP TABLE IF EXISTS agr_master")  # ← add this line

cursor.execute("""
CREATE TABLE IF NOT EXISTS agr_master  (
    enag        TEXT PRIMARY KEY,
    item_code   TEXT,
    nb          INTEGER,
    connection  TEXT,
    ratio       TEXT,
    compact     TEXT,
    currency    TEXT,
    list_price  REAL,
    pmax_mbar   INTEGER
)
""")

cursor.execute("DELETE FROM agr_master")

agr_data = [
    ("ENAG-001", "AG02V",   15,  "Threaded", "1:1",         "No",  "INR", 11463.27,  500),
    ("ENAG-002", "AG02VR",  15,  "Threaded", "1:1 to 1:10", "No",  "INR", 14877.38,  500),
    ("ENAG-003", "AG03V",   20,  "Threaded", "1:1",         "No",  "INR", 11463.27,  500),
    ("ENAG-004", "AG03VR",  20,  "Threaded", "1:1 to 1:10", "No",  "INR", 14877.38,  500),
    ("ENAG-005", "AG04V",   25,  "Threaded", "1:1",         "No",  "INR", 11463.27,  500),
    ("ENAG-006", "AG04VR",  25,  "Threaded", "1:1 to 1:10", "No",  "INR", 14877.38,  500),
    ("ENAG-007", "AG05V",   32,  "Threaded", "1:1",         "No",  "INR", 18336.71,  500),
    ("ENAG-008", "AG05VR",  32,  "Threaded", "1:1 to 1:10", "No",  "INR", 21750.82,  500),
    ("ENAG-009", "AG06V",   40,  "Threaded", "1:1",         "No",  "INR", 18336.71,  500),
    ("ENAG-010", "AG06VR",  40,  "Threaded", "1:1 to 1:10", "No",  "INR", 21750.82,  500),
    ("ENAG-011", "AG07V",   50,  "Threaded", "1:1",         "No",  "INR", 22711.75,  500),
    ("ENAG-012", "AG07VR",  50,  "Threaded", "1:1 to 1:10", "No",  "INR", 26125.86,  500),
    ("ENAG-013", "AG08AV",  65,  "Flanged",  "1:1",         "No",  "INR", 67158.48,  500),
    ("ENAG-014", "AG08AVR", 65,  "Flanged",  "1:1 to 1:10", "No",  "INR", 70572.59,  500),
    ("ENAG-015", "AG09AV",  80,  "Flanged",  "1:1",         "No",  "INR", 73873.65,  500),
    ("ENAG-016", "AG09AVR", 80,  "Flanged",  "1:1 to 1:10", "No",  "INR", 77287.76,  500),
    ("ENAG-017", "AG10AV",  100, "Flanged",  "1:1",         "No",  "INR", 142099.33, 500),
    ("ENAG-018", "AG10AVR", 100, "Flanged",  "1:1 to 1:10", "No",  "INR", 145513.44, 500),
    ("ENAG-019", "AG32AV",  32,  "Flanged",  "1:1",         "No",  "INR", 25725.66,  500),
    ("ENAG-020", "AG32AVR", 32,  "Flanged",  "1:1 to 1:10", "No",  "INR", 29139.77,  500),
    ("ENAG-021", "AG40AV",  40,  "Flanged",  "1:1",         "No",  "INR", 21556.37,  500),
    ("ENAG-022", "AG40AVR", 40,  "Flanged",  "1:1 to 1:10", "No",  "INR", 24970.48,  500),
    ("ENAG-023", "AG50AV",  50,  "Flanged",  "1:1",         "No",  "INR", 27225.83,  500),
    ("ENAG-024", "AG50AVR", 50,  "Flanged",  "1:1 to 1:10", "No",  "INR", 30639.94,  500),
    ("ENAG-025", "AGP02V",  15,  "Threaded", "1:1",         "Yes", "INR", 8377.01,   500),
    ("ENAG-026", "AGP02VR", 15,  "Threaded", "1:1 to 1:10", "Yes", "INR", 11791.12,  500),
    ("ENAG-027", "AGP03V",  20,  "Threaded", "1:1",         "Yes", "INR", 8377.01,   500),
    ("ENAG-028", "AGP03VR", 20,  "Threaded", "1:1 to 1:10", "Yes", "INR", 11791.12,  500),
    ("ENAG-029", "AGP04V",  25,  "Threaded", "1:1",         "Yes", "INR", 8377.01,   500),
    ("ENAG-030", "AGP04VR", 25,  "Threaded", "1:1 to 1:10", "Yes", "INR", 11791.12,  500),
]

cursor.executemany(
    "INSERT INTO agr_master VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
    agr_data
)

# ----------------------------
# BURNER TABLE
# ----------------------------
cursor.execute("""
CREATE TABLE IF NOT EXISTS burner_master (
    model TEXT,
    max_flow REAL,
    price REAL
)
""")

cursor.execute("DELETE FROM burner_master")

burner_data = [
    ("MG-150", 180, 65000),
    ("MG-250", 280, 82000),
    ("MG-350", 380, 100000),
    ("MG-450", 480, 118000),
    ("MG-600", 650, 155000),
]

cursor.executemany(
    "INSERT INTO burner_master VALUES (?, ?, ?)",
    burner_data
)

# ----------------------------
# GAS TRAIN TABLE
# ----------------------------
cursor.execute("""
CREATE TABLE IF NOT EXISTS gas_train_master (
    burner_model TEXT,
    flow_nm3hr REAL,
    price REAL
)
""")

cursor.execute("DELETE FROM gas_train_master")

gas_train_data = [
    ("MG-150", 150, 180000),
    ("MG-250", 250, 220000),
    ("MG-350", 350, 260000),
    ("MG-450", 400, 295200),   # your legacy value
    ("MG-600", 650, 345000),
]

cursor.executemany(
    "INSERT INTO gas_train_master VALUES (?, ?, ?)",
    gas_train_data
)

# ----------------------------
# MOTORIZED CONTROL VALVE TABLE
# ----------------------------
cursor.execute("""
CREATE TABLE IF NOT EXISTS motorized_valve_master (
    min_flow REAL,
    max_flow REAL,
    nb INTEGER,
    rated_flow REAL,
    price REAL
)
""")

cursor.execute("DELETE FROM motorized_valve_master")

motorized_valve_data = [
    (0, 4000, 250, 4000, 70000),
    (4001, 6500, 300, 5000, 80000),
    (6501, 999999, 350, 8000, 110000),
]

cursor.executemany(
    "INSERT INTO motorized_valve_master VALUES (?, ?, ?, ?, ?)",
    motorized_valve_data
)

# ----------------------------
# BUTTERFLY VALVE TABLE
# ----------------------------
cursor.execute("""
CREATE TABLE IF NOT EXISTS butterfly_valve_master (
    nb INTEGER PRIMARY KEY,
    price REAL
)
""")

cursor.execute("DELETE FROM butterfly_valve_master")

butterfly_valve_data = [
    (150, 8000),
    (300, 16950),
    (350, 22000)   # fallback case
]

cursor.executemany(
    "INSERT INTO butterfly_valve_master VALUES (?, ?)",
    butterfly_valve_data
)

# ----------------------------
# BLOWER TABLE
# ----------------------------
cursor.execute("""
CREATE TABLE IF NOT EXISTS blower_master (
    min_flow REAL,
    max_flow REAL,
    hp INTEGER,
    pressure_mm_wc REAL,
    rated_flow REAL,
    price REAL
)
""")

cursor.execute("DELETE FROM blower_master")

blower_data = [
    (0, 3000, 15, 600, 3000, 145000),
    (3001, 4500, 20, 650, 4500, 170000),
    (4501, 6000, 25, 700, 5100, 195000),  # your legacy value
    (6001, 8000, 30, 750, 7000, 230000),
]

cursor.executemany(
    "INSERT INTO blower_master VALUES (?, ?, ?, ?, ?, ?)",
    blower_data
)

# ----------------------------
# COMPENSATOR TABLE
# ----------------------------
cursor.execute("""
CREATE TABLE IF NOT EXISTS compensator_master (
    min_flow REAL,
    max_flow REAL,
    nb INTEGER,
    price REAL
)
""")

cursor.execute("DELETE FROM compensator_master")

compensator_data = [
    (0, 2000, 150, 4000),
    (2001, 3500, 200, 6000),
    (3501, 5000, 300, 8000),
    (5001, 999999, 350, 10000),
]

cursor.executemany(
    "INSERT INTO compensator_master VALUES (?, ?, ?, ?)",
    compensator_data
)

# ----------------------------
# ROTARY JOINT TABLE
# ----------------------------
cursor.execute("""
CREATE TABLE IF NOT EXISTS rotary_joint_master (
    min_flow REAL,
    max_flow REAL,
    nb INTEGER,
    price REAL
)
""")

cursor.execute("DELETE FROM rotary_joint_master")

rotary_joint_data = [
    (0, 2000, 200, 40000),
    (2001, 3500, 250, 45000),
    (3501, 5000, 300, 50000),   # legacy value
    (5001, 999999, 350, 60000),
]

cursor.executemany(
    "INSERT INTO rotary_joint_master VALUES (?, ?, ?, ?)",
    rotary_joint_data
)

# Save changes
conn.commit()

# Close connection
conn.close()

print("Database initialized successfully.")