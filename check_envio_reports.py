import sqlite3
import os

DB_PATH = "sat.db"

def check_envio_reports(envio_id):
    if not os.path.exists(DB_PATH):
        print("DB not found")
        return
    
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    
    print(f"--- Instruments for Envio {envio_id} ---")
    cur.execute("SELECT id, estado, denominacion FROM instrumentos WHERE envio_id=?", (envio_id,))
    instruments = cur.fetchall()
    for inst in instruments:
        inst_id = inst['id']
        cur.execute("SELECT count(*) FROM instrumento_informes WHERE instrumento_id=?", (inst_id,))
        count = cur.fetchone()[0]
        print(f"ID: {inst_id} | Status: {inst['estado']} | Name: {inst['denominacion']} | Reports count: {count}")
        
        if count > 0:
            cur.execute("SELECT * FROM instrumento_informes WHERE instrumento_id=?", (inst_id,))
            reports = cur.fetchall()
            for r in reports:
                print(f"  Report: {dict(r)}")

    print("\n--- Any global reports? ---")
    cur.execute("SELECT count(*) FROM instrumento_informes")
    print(f"Total reports in DB: {cur.fetchone()[0]}")
    
    cur.execute("SELECT * FROM instrumento_informes ORDER BY id DESC LIMIT 5")
    for r in cur.fetchall():
        print(dict(r))

    conn.close()

if __name__ == "__main__":
    check_envio_reports(1)
