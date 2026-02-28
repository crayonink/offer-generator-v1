import sqlite3
import os

# Always points to vlph.db in project root, regardless of where you run from
BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DB_PATH  = os.path.join(BASE_DIR, "vlph.db")
def select_agr(nb: int, connection: str, ratio: str, compact: str) -> dict:
    """
    AGR selector using database (vlph.db)
    Selection based on line size (NB), connection type, ratio and compact version.

    Args:
        nb         : Line size in mm (e.g. 15, 20, 25, 32, 40, 50, 65, 80, 100)
        connection : "Threaded" or "Flanged"
        ratio      : "1:1" or "1:1 to 1:10"
        compact    : "Yes" or "No"
    """

    conn = sqlite3.connect(DB_PATH)

    cursor = conn.cursor()

    cursor.execute("""
        SELECT enag, item_code, nb, connection, ratio, compact, list_price, pmax_mbar
        FROM agr_master
        WHERE nb         = ?
          AND connection = ?
          AND ratio      = ?
          AND compact    = ?
        LIMIT 1
    """, (nb, connection, ratio, compact))

    row = cursor.fetchone()
    conn.close()

    if not row:
        raise ValueError(
            f"No AGR found for NB={nb}, connection={connection}, "
            f"ratio={ratio}, compact={compact}"
        )

    return {
        "enag":       row[0],
        "item_code":  row[1],
        "nb":         row[2],
        "connection": row[3],
        "ratio":      row[4],
        "compact":    row[5],
        "price":      row[6],
        "pmax_mbar":  row[7],
    }


# Example usage
if __name__ == "__main__":
    # Print whole database
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute("SELECT * FROM agr_master")
    rows = cursor.fetchall()
    
    # Print header
    cursor.execute("PRAGMA table_info(agr_master)")
    columns = [col[1] for col in cursor.fetchall()]
    print(" | ".join(columns))
    print("-" * 100)
    
    # Print rows
    for row in rows:
        print(" | ".join(str(v) for v in row))
    
    print(f"\nTotal: {len(rows)} rows")
    conn.close()
