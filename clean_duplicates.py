# coding: utf-8
"""
clean_duplicates.py
-------------------
Merges duplicate/near-duplicate items in component_price_master.
Keeps the canonical name, deletes the alias rows.
Run once: python clean_duplicates.py
"""
import sqlite3, os

DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), "vlph.db")

# (canonical_name, [aliases_to_delete])
MERGES = [
    # Flexible hoses — trailing space variants
    ("FLEXIBLE HOSE-15NB*1000MM (OIL)", ["FLEXIBLE HOSE-15NB*1000MM (OIL )"]),
    ("FLEXIBLE HOSE-15NB*750MM (OIL)",  ["FLEXIBLE HOSE-15NB*750MM (OIL )"]),
    ("FLEXIBLE HOSE-20NB*1000MM (OIL)", ["FLEXIBLE HOSE-20NB*1000MM (OIL )"]),
    ("FLEXIBLE HOSE-25NB*1000MM (AIR)", ["FLEXIBLE HOSE-25NB*1000MM (AIR )"]),

    # MS sheets — double space
    ("M.S. Sheet 2mm",  ["M.S. Sheet  2mm"]),
    ("M.S. Sheet 3mm",  ["M.S. Sheet  3mm"]),

    # MS plate
    ("M.S. Plate 16mm*5mm", ["M.S. Plate 16mm* 5mm"]),

    # MS tube — quotes vs no quotes
    ('M.S. Tube B Class 1.5 in', ['M.S. Tube "B" Class 1.5 in']),
    ('M.S. Tube C Class 1.5 in', ['M.S. Tube "C" Class 1.5 in']),

    # MS Chanel — with/without dot-space
    ("M.S. Chanel", ["M.S.Chanel"]),

    # Plumber block / Pulley — case
    ("Plumber block with Bearing", ["Plumber Block with Bearing"]),
    ("Pulley with V belt",         ["Pulley with V Belt"]),

    # SS Pipe — keep the (per mtr) version, remove plain and space-X variant
    ("SS Pipe 304 60x3mm (per mtr)",  ["SS Pipe 304 60x3mm",  "SS Pipe 304 60 X 3mm"]),
    ("SS Pipe 304 76x3mm (per mtr)",  ["SS Pipe 304 76x3mm",  "SS Pipe 304 76 X 3mm"]),
    ("SS Pipe 304 100x3mm (per mtr)", ["SS Pipe 304 100x3mm", "SS Pipe 304 100 X 3mm"]),

    # ID FAN double space
    ("ID FAN (ARE 35)", ["ID FAN  (ARE 35)"]),

    # SEQUENCE — keep full name
    ("SEQUENCE CONTROLLER", ["SEQUENCE"]),
]

conn = sqlite3.connect(DB)

for canonical, aliases in MERGES:
    for alias in aliases:
        # Check alias exists
        row = conn.execute(
            "SELECT rowid, price FROM component_price_master WHERE item=?", (alias,)
        ).fetchone()
        if not row:
            print(f"  SKIP (not found): {alias!r}")
            continue
        alias_price = row[1]

        # Check canonical exists
        can_row = conn.execute(
            "SELECT rowid, price FROM component_price_master WHERE item=?", (canonical,)
        ).fetchone()
        if not can_row:
            # Rename alias to canonical
            conn.execute(
                "UPDATE component_price_master SET item=? WHERE item=?", (canonical, alias)
            )
            print(f"  RENAME: {alias!r} -> {canonical!r}")
        else:
            # Delete alias
            conn.execute("DELETE FROM component_price_master WHERE item=?", (alias,))
            print(f"  DELETE: {alias!r} (kept canonical {canonical!r} @ {can_row[1]})")

conn.commit()

# Verify
remaining = conn.execute(
    "SELECT COUNT(*) FROM component_price_master"
).fetchone()[0]
print(f"\nDone. {remaining} items remain in component_price_master.")
conn.close()
