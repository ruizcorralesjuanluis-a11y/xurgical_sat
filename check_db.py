from db import get_conn

c = get_conn()
cur = c.cursor()

cur.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name;")
tables = [r[0] for r in cur.fetchall()]
print("TABLAS:", tables)

for t in ["envios", "instrumentos", "parts", "instruments"]:
    try:
        cur.execute(f"SELECT COUNT(*) FROM {t};")
        print(t, cur.fetchone()[0])
    except Exception:
        print(t, "NO")

c.close()
