
import os
import sqlite3
from pathlib import Path

# -------------------------------------------------
# RUTA BASE DEL PROYECTO
# -------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent

# -------------------------------------------------
# RUTA DE LA BASE DE DATOS
# -------------------------------------------------
DB_PATH = Path(
    os.environ.get(
        "XURGICAL_DB_PATH",
        str(BASE_DIR / "sat.db")
    )
)

# -------------------------------------------------
# CONEXIÓN A SQLITE
# -------------------------------------------------
def get_connection():
    # Asegura que el directorio exista (útil si se usa una ruta custom)
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def get_conn():
    return get_connection()

# -------------------------------------------------
# INICIALIZACIÓN DE LA BD
# -------------------------------------------------
def init_db():
    conn = get_connection()
    cursor = conn.cursor()

    # 1. Usuarios
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        role TEXT NOT NULL,
        is_active INTEGER NOT NULL DEFAULT 1,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP
    )
    """)

    # 2. Clientes
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS clientes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT UNIQUE NOT NULL,
        prefijo TEXT,
        prefijo_nombre TEXT,
        ultimo_numero INTEGER DEFAULT 0
    )
    """)

    # 3. Envíos (Partes de trabajo)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS envios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ot_num TEXT UNIQUE,
        nombre_archivo TEXT,
        cliente TEXT,
        cliente_id INTEGER,
        tipo_trabajo TEXT DEFAULT 'REPARACION',
        fecha TEXT,
        creado_en DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (cliente_id) REFERENCES clientes(id)
    )
    """)

    # 4. Instrumentos (Ítems de cada envío)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS instrumentos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        envio_id INTEGER NOT NULL,
        codigo_producto TEXT,
        fabricante TEXT,
        num_serie TEXT,
        denominacion TEXT,
        observaciones TEXT,
        codigo_datamatrix TEXT,
        nombre_trazabilidad TEXT,
        estado TEXT DEFAULT 'Pendiente',
        grabado INTEGER DEFAULT 0,
        grabado_por INTEGER,
        grabado_en TEXT,
        creado_en DATETIME DEFAULT CURRENT_TIMESTAMP,
        actualizado_en DATETIME DEFAULT CURRENT_TIMESTAMP,
        foto_entrada_1 TEXT,
        foto_entrada_2 TEXT,
        FOREIGN KEY (envio_id) REFERENCES envios(id) ON DELETE CASCADE
    )
    """)

    # 5. Permisos granulares
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS user_permissions (
      user_id INTEGER NOT NULL,
      action TEXT NOT NULL,
      allowed INTEGER NOT NULL DEFAULT 0,
      PRIMARY KEY (user_id, action),
      FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
    )
    """)

    # 6. Catálogo de Checklist
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS checklist_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT NOT NULL,
        orden INTEGER DEFAULT 0,
        activo INTEGER DEFAULT 1,
        tipo_trabajo TEXT DEFAULT 'REPARACION'
    )
    """)

    # 7. Checklist por Instrumento
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS instrumento_checklist (
        instrumento_id INTEGER NOT NULL,
        item_id INTEGER NOT NULL,
        hecho INTEGER DEFAULT 0,
        hecho_por TEXT,
        hecho_en TEXT,
        PRIMARY KEY (instrumento_id, item_id),
        FOREIGN KEY (instrumento_id) REFERENCES instrumentos(id) ON DELETE CASCADE,
        FOREIGN KEY (item_id) REFERENCES checklist_items(id) ON DELETE CASCADE
    )
    """)

    conn.commit()
    conn.close()
