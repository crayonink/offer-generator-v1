import sqlite3
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "vlph.db")

print("Using DB:", DB_PATH)

conn = sqlite3.connect(DB_PATH)
cursor = conn.cursor()

cursor.execute("""
SELECT airflow, price_basic 
FROM blower_master 
WHERE airflow = 2000
""")

rows = cursor.fetchall()
print("Result:", rows)

conn.close()