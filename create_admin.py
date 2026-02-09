# create_admin.py
from db import init_db, get_conn
from auth_utils import hash_password

init_db()
conn = get_conn()
cur = conn.cursor()

username = "admin"
password = "admin123"
role = "admin"

cur.execute("SELECT id FROM users WHERE username=?", (username,))
exists = cur.fetchone()

if not exists:
    cur.execute(
        "INSERT INTO users (username, password_hash, role, is_active) VALUES (?, ?, ?, 1)",
        (username, hash_password(password), role),
    )
    conn.commit()
    print("✅ Admin creado: admin / admin123")
else:
    print("ℹ️ Admin ya existe")

conn.close()
