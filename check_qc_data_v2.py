import sqlite3
import os

DB_PATH = "sat.db"

def check_db():
    if not os.path.exists(DB_PATH):
        print("DB not found")
        return
    
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    
    print("--- Tables ---")
    cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
    for table in cur.fetchall():
        print(table['name'])
    
    print("\n--- Rows in instrumento_informes ---")
    try:
        cur.execute("SELECT * FROM instrumento_informes")
        rows = cur.fetchall()
        for row in rows:
            print(dict(row))
    except Exception as e:
        print(f"Error reading instrumento_informes: {e}")

    print("\n--- Rows in instrumentos ---")
    try:
        cur.execute("SELECT id, case estado when 'Reparado' then 1 else 0 end as rep FROM instrumentos LIMIT 10")
        rows = cur.fetchall()
        for row in rows:
            print(dict(row))
    except Exception as e:
        print(f"Error reading instrumentos: {e}")

    conn.close()

if __name__ == "__main__":
    check_db()
