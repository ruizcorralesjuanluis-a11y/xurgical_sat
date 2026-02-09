import os
import sqlite3
from pathlib import Path

# -------------------------------------------------
# RUTA BASE DEL PROYECTO
# -------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent

# -------------------------------------------------
# RUTA DE LA BASE DE DATOS
# - Por defecto: sat.db junto a este archivo
# - Opcional: variable de entorno XURGICAL_DB_PATH
#   (para BD compartida en red en el futuro)
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
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


# -------------------------------------------------
# ALIAS PARA COMPATIBILIDAD CON app.py
# -------------------------------------------------
def get_conn():
    return get_connection()


# -------------------------------------------------
# INICIALIZACIÓN DE LA BD
# (no borra datos existentes)
# -------------------------------------------------
def init_db():
    conn = get_connection()
    cursor = conn.cursor()

    # Tabla de usuarios
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        role TEXT NOT NULL
    )
    """)

    # Tabla de instrumentos
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS instruments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        datamatrix TEXT UNIQUE,
        nombre TEXT,
        estado TEXT,
        observaciones TEXT
    )
    """)

    # Tabla de partes de trabajo
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS work_orders (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        numero_parte TEXT,
        cliente TEXT,
        estado TEXT
    )
    """)

    conn.commit()
    conn.close()
