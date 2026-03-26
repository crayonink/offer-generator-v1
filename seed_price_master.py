"""
Seed component_price_master table from the Pricelist WorkBook 28-08-2025.
Run once: python seed_price_master.py
"""
import sqlite3, os

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "vlph.db")

# (item, category, unit, price_2025, previous_price)
PRICES = [
    # ── RAW MATERIALS ─────────────────────────────────────────────────────────
    ("C.I. Gills",                    "Raw Material", "kg",  170.00,   170.00),
    ("C.I. Hub",                      "Raw Material", "kg",  375.00,   375.00),
    ("Coupling",                      "Raw Material", "nos", 350.00,   250.00),
    ("Hardware Bolt",                 "Raw Material", "kg",  150.00,   150.00),
    ("M.S. Angle 100,100",            "Raw Material", "kg",   71.00,    71.00),
    ("M.S. Angle 50*6",               "Raw Material", "kg",   71.00,    71.00),
    ("M.S. Angle 65,50",              "Raw Material", "kg",   71.00,    71.00),
    ("M.S. Chanel",                   "Raw Material", "kg",   71.00,    71.00),
    ("M.S. Flat",                     "Raw Material", "kg",   71.00,    71.00),
    ("M.S. Plate 16mm*5mm",           "Raw Material", "kg",   72.00,    72.00),
    ("M.S. Plate 16mm*10mm",          "Raw Material", "kg",   72.00,    72.00),
    ("M.S. Plate 5mm",                "Raw Material", "kg",   72.00,    72.00),
    ("M.S. Plate 8mm",                "Raw Material", "kg",   72.00,    72.00),
    ("M.S. Round",                    "Raw Material", "kg",   72.00,    72.00),
    ("M.S. Sheet 2mm",                "Raw Material", "kg",   77.00,    77.00),
    ("M.S. Sheet 3mm",                "Raw Material", "kg",   77.00,    77.00),
    ("M.S. Sheet 4mm",                "Raw Material", "kg",   72.00,    72.00),
    ("M.S. Sheet 5mm",                "Raw Material", "kg",   72.00,    72.00),
    ("M.S. Sheet 8mm",                "Raw Material", "kg",   72.00,    72.00),
    ("M.S. Tube B Class 1.5 in",      "Raw Material", "kg",  135.00,   135.00),
    ("M.S. Tube C Class 1.5 in",      "Raw Material", "kg",  135.00,   135.00),
    ("S.S. Sheet 3mm",                "Raw Material", "kg",  270.00,   270.00),
    ("Ceramic Fiber",                 "Raw Material", "kg", 2000.00,  1450.00),
    ("SS Pipe 304 60x3mm",            "Raw Material", "kg",  330.00,   330.00),
    ("SS Pipe 304 76x3mm",            "Raw Material", "kg",  330.00,   330.00),
    ("SS Pipe 304 100x3mm",           "Raw Material", "kg",  330.00,   330.00),
    ("Plumber Block with Bearing",    "Raw Material", "nos",13000.00, 11000.00),
    ("Pulley with V Belt",            "Raw Material", "nos", 5000.00,  3000.00),
    # Pipe prices per metre
    ("SS Pipe 304 60x3mm (per mtr)",  "Raw Material", "mtr",1391.65,  1391.65),
    ("SS Pipe 304 76x3mm (per mtr)",  "Raw Material", "mtr",1782.29,  1782.29),
    ("SS Pipe 304 100x3mm (per mtr)", "Raw Material", "mtr",2368.24,  2368.24),
    # Elbows
    ("M.S. Elbow 60 OD",              "Raw Material", "nos",  200.00,   200.00),
    ("M.S. Elbow 76 OD",              "Raw Material", "nos",  300.00,   300.00),
    ("M.S. Elbow 100 OD",             "Raw Material", "nos",  400.00,   400.00),
    ("SS Elbow 60 OD",                "Raw Material", "nos",  600.00,   500.00),
    ("SS Elbow 76 OD",                "Raw Material", "nos",  700.00,   600.00),
    ("SS Elbow 100 OD",               "Raw Material", "nos", 1000.00,   900.00),

    # ── BOUGHT OUT ITEMS ──────────────────────────────────────────────────────
    ("Casting Burner Parts",                      "Bought Out", "kg",  100.00,    90.00),
    ("SS Assly 2A/3A",                            "Bought Out", "nos", 2500.00,  2300.00),
    ("SS Assly 4A",                               "Bought Out", "nos", 2600.00,  2400.00),
    ("SS Assly 5A/6A",                            "Bought Out", "nos", 5000.00,  4500.00),
    ("SS Assly 7A",                               "Bought Out", "nos",26000.00, 23000.00),
    ("Micro Valve 2A/3A",                         "Bought Out", "nos", 2300.00,  2075.00),
    ("Micro Valve 4A",                            "Bought Out", "nos", 2300.00,  2075.00),
    ("Micro Valve 5A/6A",                         "Bought Out", "nos", 2300.00,  2075.00),
    ("Micro Valve 7A",                            "Bought Out", "nos", 2500.00,  2125.00),
    ("FLEXIBLE HOSE-15NB*750MM (OIL)",            "Bought Out", "nos",  590.00,   590.00),
    ("FLEXIBLE HOSE-15NB*1000MM (OIL)",           "Bought Out", "nos",  750.00,   750.00),
    ("FLEXIBLE HOSE-20NB*1000MM (OIL)",           "Bought Out", "nos",  800.00,   800.00),
    ("FLEXIBLE HOSE-25NB*1000MM (AIR)",           "Bought Out", "nos", 1100.00,  1100.00),
    ("FLEXIBLE HOSE-40NB*750MM WITH ADOPTER",     "Bought Out", "nos", 1500.00,  1500.00),
    ("ADOPTER 15NB*20NB (OIL)",                   "Bought Out", "nos",   30.00,    30.00),
    ("ADOPTER 15NB*15NB (AIR)",                   "Bought Out", "nos",   30.00,    30.00),
    ("Butterfly Valve 2.5\"",                     "Bought Out", "nos", 1500.00,  1500.00),
    ("Butterfly Valve 4\"",                       "Bought Out", "nos", 1650.00,  1650.00),
    ("Butterfly Valve 6\"",                       "Bought Out", "nos", 4500.00,  4500.00),
    ("ENCON-Y-STRAINER 20NB",                     "Bought Out", "nos",  450.00,   360.00),
    ("Whyteheat K",                               "Bought Out", "nos",   70.00,    50.00),
    ("Ball valve 20 NB",                          "Bought Out", "nos", 1718.00,  1718.00),
    ("Ball valve 25 NB",                          "Bought Out", "nos", 2500.00,  2500.00),
    ("Ball valve 32 NB",                          "Bought Out", "nos", 2929.00,  2929.00),
    ("Ball valve 40 NB",                          "Bought Out", "nos", 5000.00,  5000.00),

    # ── ENCON PURCHASE PRICES ─────────────────────────────────────────────────
    ("SEQUENCE",                  "ENCON Purchase", "nos", 10640.00,  5935.40),
    ("VACUUM SWITCH",             "ENCON Purchase", "nos",  1200.00,   882.00),
    ("ARE-13 Burner Alone",       "ENCON Purchase", "nos", 30000.00, 21550.00),
    ("ARE-22 Burner Alone",       "ENCON Purchase", "nos", 34400.00, 32623.00),
    ("ARE-35 Burner Alone",       "ENCON Purchase", "nos", 49000.00, 39547.05),
    ("ARE-50 Burner Alone",       "ENCON Purchase", "nos", 50000.00, 38000.00),
    ("ID FAN (ARE 13)",           "ENCON Purchase", "nos", 14000.00, 10278.46),
    ("ID FAN (ARE 35)",           "ENCON Purchase", "nos", 15247.00, 10877.70),
    ("ID FAN (HEAVY DUTY)",       "ENCON Purchase", "nos", 25550.00, 19254.00),
    ("ID FAN (TYPE O FAN)",       "ENCON Purchase", "nos", 45000.00, 41236.00),
    ("SOLENOID VALVE UNIT",       "ENCON Purchase", "nos",  8000.00,  4744.00),
    ("ELECTRODE ASSEMBLY",        "ENCON Purchase", "nos",  2200.00,   560.00),
    ("NOISE FILTER",              "ENCON Purchase", "nos",   400.00,   250.00),
    ("INPUT SOCKET",              "ENCON Purchase", "nos",  1070.00,   280.00),

    # ── BOM ITEMS (used directly by VLPH/HLPH builders) ──────────────────────
    ("COMPENSATOR",                       "Bought Out", "nos",  3500.00,  3000.00),
    ("PRESSURE GAUGE WITH TNV",           "Bought Out", "nos",   850.00,   750.00),
    ("PRESSURE GAUGE WITH NV",            "Bought Out", "nos",   750.00,   650.00),
    ("PRESSURE SWITCH LOW",               "Bought Out", "nos",  1200.00,  1000.00),
    ("PRESSURE SWITCH HIGH + LOW",        "Bought Out", "nos",  2200.00,  1800.00),
    ("SOLENOID VALVE",                    "Bought Out", "nos",  2500.00,  2000.00),
    ("PRESSURE REGULATING VALVE",         "Bought Out", "nos",  4500.00,  4000.00),
    ("FLEXIBLE HOSE (Pilot Burner)",      "Bought Out", "nos",   800.00,   750.00),
    ("FLEXIBLE HOSE (UV LINE)",           "Bought Out", "nos",   590.00,   590.00),
    ("FLEXIBLE HOSE PIPE",                "Bought Out", "nos",   590.00,   590.00),
    ("BALL VALVE",                        "Bought Out", "nos",  1718.00,  1718.00),
    ("BALL VALVE (Pilot Burner)",         "Bought Out", "nos",  2500.00,  2500.00),
    ("BALL VALVE (UV LINE)",              "Bought Out", "nos",  1718.00,  1718.00),
    ("THERMOCOUPLE",                      "Bought Out", "nos",  3500.00,  3000.00),
    ("COMPENSATING LEAD",                 "Bought Out", "mtr",   120.00,   100.00),
    ("LIMIT SWITCHES",                    "Bought Out", "nos",  2200.00,  1800.00),
    ("CONTROL PANEL",                     "Bought Out", "nos", 85000.00, 75000.00),
    ("HYDRAULIC POWER PACK & CYLINDER",   "Bought Out", "nos",125000.00,110000.00),
    ("CABLE FOR IGNITION TRANSFORMER",    "Bought Out", "mtr",   180.00,   150.00),
    ("TEMPERATURE TRANSMITTER",           "Bought Out", "nos",  4500.00,  4000.00),
    ("P.PID",                             "Bought Out", "nos",  8500.00,  7500.00),
    ("ENCON-PB (NG/LPG) - 100 KW",       "ENCON Purchase", "nos", 35000.00, 30000.00),
    ("Ignition Transformer",              "ENCON Purchase", "nos",  4744.60,  4000.00),
    ("Burner Control Unit",               "ENCON Purchase", "nos", 10640.00,  5935.40),
    ("UV Sensor with Air Jacket",         "ENCON Purchase", "nos",  3500.00,  3000.00),
]


def seed():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS component_price_master (
            item           TEXT PRIMARY KEY,
            category       TEXT,
            unit           TEXT,
            price          REAL,
            previous_price REAL
        )
    """)

    inserted = 0
    updated  = 0
    for item, category, unit, price, prev in PRICES:
        existing = conn.execute(
            "SELECT price FROM component_price_master WHERE item = ?", (item,)
        ).fetchone()

        if existing is None:
            conn.execute(
                "INSERT INTO component_price_master (item, category, unit, price, previous_price) VALUES (?,?,?,?,?)",
                (item, category, unit, price, prev)
            )
            inserted += 1
        else:
            conn.execute(
                "UPDATE component_price_master SET category=?, unit=?, price=?, previous_price=? WHERE item=?",
                (category, unit, price, prev, item)
            )
            updated += 1

    conn.commit()
    conn.close()
    print(f"Done — {inserted} inserted, {updated} updated  ({len(PRICES)} total rows)")


if __name__ == "__main__":
    seed()
