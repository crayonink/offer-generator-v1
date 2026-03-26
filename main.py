from fastapi import FastAPI, UploadFile, File, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse
import pandas as pd
import sqlite3
import shutil
import os
from datetime import datetime

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "vlph.db")
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

VALID_TABLES = None


def ensure_log_table():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS table_update_log (
            table_name  TEXT PRIMARY KEY,
            updated_at  TEXT,
            uploaded_by TEXT
        )
    """)
    conn.commit()
    conn.close()


ensure_log_table()


# =================================================
# DB VIEWER
# =================================================

@app.get("/viewer", response_class=HTMLResponse)
def db_viewer():
    html_path = os.path.join(BASE_DIR, "db_viewer.html")
    with open(html_path, "r", encoding="utf-8") as f:
        content = f.read()
    # Replace localhost API with relative path so it works on any domain
    content = content.replace(
        'const API = "http://127.0.0.1:8000"',
        'const API = ""'
    )
    return HTMLResponse(content=content)


# =================================================
# UPLOAD EXCEL
# =================================================

@app.post("/upload-excel/")
async def upload_excel(request: Request, file: UploadFile = File(...)):
    try:
        filename = file.filename
        table_name = filename.replace(".xlsx", "").strip().lower()

        if VALID_TABLES and table_name not in VALID_TABLES:
            return {"error": f"Invalid table name: {table_name}"}

        file_path = os.path.join(UPLOAD_FOLDER, filename)

        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        df = pd.read_excel(file_path)

        if df.empty:
            return {"error": "Excel file is empty"}

        df.columns = [col.strip().lower() for col in df.columns]

        client_ip = request.client.host if request.client else "unknown"
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        conn = sqlite3.connect(DB_PATH)
        df.to_sql(table_name, conn, if_exists="replace", index=False)

        conn.execute("""
            INSERT INTO table_update_log (table_name, updated_at, uploaded_by)
            VALUES (?, ?, ?)
            ON CONFLICT(table_name) DO UPDATE SET
                updated_at = excluded.updated_at,
                uploaded_by = excluded.uploaded_by
        """, (table_name, now, client_ip))

        conn.commit()
        conn.close()

        return {
            "message": f"{table_name} updated successfully!",
            "updated_at": now
        }

    except Exception as e:
        return {"error": str(e)}


# =================================================
# GET TABLE LIST
# =================================================

@app.get("/db/tables")
def get_tables():
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()

        cursor.execute("""
            SELECT name FROM sqlite_master
            WHERE type='table'
              AND name NOT LIKE 'sqlite_%'
              AND name != 'table_update_log'
            ORDER BY name
        """)

        tables = [row[0] for row in cursor.fetchall()]

        cursor.execute("SELECT table_name, updated_at, uploaded_by FROM table_update_log")
        log = {
            row[0]: {"updated_at": row[1], "uploaded_by": row[2]}
            for row in cursor.fetchall()
        }

        result = []
        for table in tables:
            cursor.execute(f"SELECT COUNT(*) FROM [{table}]")
            count = cursor.fetchone()[0]
            entry = log.get(table, {})
            result.append({
                "name": table,
                "rows": count,
                "updated_at": entry.get("updated_at"),
                "uploaded_by": entry.get("uploaded_by")
            })

        conn.close()
        return {"tables": result}

    except Exception as e:
        return {"error": str(e)}


# =================================================
# GET TABLE DATA
# =================================================

@app.get("/db/table/{table_name}")
def get_table_data(table_name: str):

    if VALID_TABLES and table_name not in VALID_TABLES:
        return {"error": "Invalid table name"}

    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()

        cursor.execute(f"SELECT * FROM [{table_name}]")
        rows = cursor.fetchall()
        columns = [desc[0] for desc in cursor.description]

        cursor.execute(
            "SELECT updated_at, uploaded_by FROM table_update_log WHERE table_name = ?",
            (table_name,)
        )
        log_row = cursor.fetchone()

        conn.close()

        return {
            "table": table_name,
            "columns": columns,
            "rows": [list(r) for r in rows],
            "total": len(rows),
            "updated_at": log_row[0] if log_row else None,
            "uploaded_by": log_row[1] if log_row else None,
        }

    except Exception as e:
        return {"error": str(e)}