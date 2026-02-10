import sqlite3
import os

db_path = "sat.db"
if os.path.exists(db_path):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute("SELECT * FROM instrumentos ORDER BY id DESC LIMIT 5")
    rows = cur.fetchall()
    for r in rows:
        d = dict(r)
        print(f"ID: {d['id']}")
        for i in range(1, 7):
            print(f"  foto_entrada_{i}: {d.get(f'foto_entrada_{i}')}")
    conn.close()
else:
    print("DB not found")
