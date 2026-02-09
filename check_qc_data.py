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
    
    print("--- instrument_qc_optica ---")
    cur.execute("SELECT * FROM instrumento_qc_optica")
    rows = cur.fetchall()
    for row in rows:
        print(dict(row))
        
    print("\n--- instrumento_informes ---")
    cur.execute("SELECT * FROM instrumento_informes")
    rows = cur.fetchall()
    for row in rows:
        print(dict(row))
        
    conn.close()

if __name__ == "__main__":
    check_db()
