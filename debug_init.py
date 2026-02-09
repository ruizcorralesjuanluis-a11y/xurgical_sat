import sqlite3
import os
from pathlib import Path
from db import init_db, get_connection, BASE_DIR, DB_PATH

def check_sat_db():
    print(f"Checking DB at: {DB_PATH}")
    if not os.path.exists(DB_PATH):
        print("sat.db not found (will be created by init_db)")
    
    try:
        init_db()
        print("init_db() executed successfully.")
    except Exception as e:
        print(f"init_db() FAILED: {e}")
        return

    conn = get_connection()
    cur = conn.cursor()
    # PGCursorWrapper vs sqlite3
    
    try:
        # Check table
        # If it is wrapper, execute vs sqlite3 directly is safer for verification?
        # get_connection returns wrapper if PG, or sqlite3 conn if local.
        # Assuming local sqlite3 for user.
        
        # Verify if table exists
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='instrumento_informes'")
        if cur.fetchone():
            print("Table 'instrumento_informes' EXISTS.")
        else:
            print("Table 'instrumento_informes' DOES NOT EXIST.")

        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='instrumento_qc_optica'")
        if cur.fetchone():
            print("Table 'instrumento_qc_optica' EXISTS.")
        else:
            print("Table 'instrumento_qc_optica' DOES NOT EXIST.")
            
    except Exception as e:
        print(f"Error checking validation: {e}")
    finally:
        conn.close()

if __name__ == "__main__":
    check_sat_db()
