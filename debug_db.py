import sqlite3
import os

DB_PATH = "database.db"  # Assuming standard path, adjust if necessary.
# Based on app.py: from db import init_db, get_conn.
# I'll try to find where the DB is. Usually 'database.db' in current dir?

def check_schema():
    if not os.path.exists("database.db"):
        print("database.db not found")
        return

    conn = sqlite3.connect("database.db")
    cur = conn.cursor()
    
    # Check tables
    cur.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = [r[0] for r in cur.fetchall()]
    print("Tables:", tables)

    if "instrumentos" in tables:
        cur.execute("PRAGMA table_info(instrumentos)")
        cols = [r[1] for r in cur.fetchall()]
        print("instrumentos cols:", cols)

    if "envios" in tables:
        cur.execute("PRAGMA table_info(envios)")
        cols = [r[1] for r in cur.fetchall()]
        print("envios cols:", cols)
        
    if "instrumento_informes" in tables:
        print("instrumento_informes exists")
    else:
        print("instrumento_informes MISSING!")

    conn.close()

if __name__ == "__main__":
    check_schema()
