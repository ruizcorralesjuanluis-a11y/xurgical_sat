
import sqlite3
import os

db_path = "sat.db"
print(f"Checking access to: {os.path.abspath(db_path)}")

if not os.path.exists(db_path):
    print("ERROR: File does not exist!")
else:
    try:
        conn = sqlite3.connect(db_path)
        cur = conn.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cur.fetchall()
        print("Success! Tables found:", [t[0] for t in tables])
        conn.close()
    except Exception as e:
        print(f"ERROR: Failed to connect. {e}")
