
import os
import sqlite3
from pathlib import Path

# -------------------------------------------------
# RUTA BASE DEL PROYECTO
# -------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent

# -------------------------------------------------
# RUTA DE LA BASE DE DATOS (PRO o LOCAL)
# -------------------------------------------------
DATABASE_URL = os.environ.get("DATABASE_URL")
DB_PATH = Path(os.environ.get("XURGICAL_DB_PATH", str(BASE_DIR / "sat.db")))

# -------------------------------------------------
# WRAPPER PARA COMPATIBILIDAD (SQLite vs Postgres)
# -------------------------------------------------
class PGRowWrapper(dict):
    """Permite acceso por atributo o por clave como sqlite3.Row"""
    def __getattr__(self, name):
        return self.get(name)

class PGCursorWrapper:
    def __init__(self, cur):
        self.cur = cur
    
    def execute(self, sql, params=None):
        # Convertimos ? a %s para Postgres
        if params:
            sql = sql.replace('?', '%s')
        # Limpieza de SQL específico de SQLite
        sql = sql.replace("BEGIN IMMEDIATE", "BEGIN")
        return self.cur.execute(sql, params)

    def fetchone(self):
        row = self.cur.fetchone()
        if row: return PGRowWrapper(row)
        return None

    def fetchall(self):
        return [PGRowWrapper(r) for r in self.cur.fetchall()]

    def __iter__(self):
        return iter(self.fetchall())
    
    @property
    def lastrowid(self):
        return getattr(self.cur, "lastrowid", None) # psycopg2 no tiene lastrowid igual, se suele usar RETURNING

def get_table_columns(cur, table_name: str) -> list[str]:
    """Obtiene nombres de columnas de forma compatible con SQLite y Postgres"""
    # Detectamos si es Postgres por el tipo de cursor o conexión
    is_pg = False
    # Si viene del wrapper, cur es PGCursorWrapper
    if isinstance(cur, PGCursorWrapper):
        is_pg = True
    
    if is_pg:
        cur.execute(
            "SELECT column_name FROM information_schema.columns WHERE table_name=%s",
            (table_name,)
        )
        return [r["column_name"] for r in cur.fetchall()]
    else:
        cur.execute(f"PRAGMA table_info({table_name})")
        return [r["name"] for r in cur.fetchall()]

class PGConnWrapper:
    def __init__(self, conn):
        self.conn = conn
    def cursor(self):
        import psycopg2.extras
        return PGCursorWrapper(self.conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor))
    def commit(self):
        self.conn.commit()
    def rollback(self):
        self.conn.rollback()
    def close(self):
        self.conn.close()

# -------------------------------------------------
# CONEXIÓN
# -------------------------------------------------
def get_connection():
    if DATABASE_URL:
        import psycopg2
        # Limpieza de la URL para evitar errores comunes
        url = DATABASE_URL.strip()
        if url.startswith("postgres://"):
            url = url.replace("postgres://", "postgresql://", 1)
        
        try:
            conn = psycopg2.connect(url)
            print(">>> [DATABASE] ESTADO: Conectado a PostgreSQL (Neon/Producción)", flush=True)
            return PGConnWrapper(conn)
        except Exception as e:
            # Si falla, imprimimos el error para depurar en logs de Render
            print(f">>> [DATABASE] ESTADO: ERROR CRITICO - Falló la conexión a Postgres: {e}", flush=True)
            raise
    
    # Fallback a SQLite
    print(f">>> [DATABASE] ESTADO: ATENCION - Usando SQLite local. LOS DATOS SE PERDERAN al actualizar.", flush=True)
    try:
        if not str(DB_PATH).startswith("\\\\"):
            DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass
    
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
    cur = conn.cursor()

    is_pg = DATABASE_URL is not None
    pk = "SERIAL PRIMARY KEY" if is_pg else "INTEGER PRIMARY KEY AUTOINCREMENT"
    text = "TEXT"
    dt = "TIMESTAMP DEFAULT CURRENT_TIMESTAMP" if is_pg else "DATETIME DEFAULT CURRENT_TIMESTAMP"

    # 1. Usuarios
    cur.execute(f"""
    CREATE TABLE IF NOT EXISTS users (
        id {pk},
        username TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        role TEXT NOT NULL,
        is_active INTEGER NOT NULL DEFAULT 1,
        created_at {dt}
    )
    """)

    # 2. Clientes
    cur.execute(f"""
    CREATE TABLE IF NOT EXISTS clientes (
        id {pk},
        nombre TEXT UNIQUE NOT NULL,
        prefijo TEXT,
        prefijo_nombre TEXT,
        ultimo_numero INTEGER DEFAULT 0
    )
    """)

    # 3. Envíos (Partes de trabajo)
    cur.execute(f"""
    CREATE TABLE IF NOT EXISTS envios (
        id {pk},
        ot_num TEXT UNIQUE,
        nombre_archivo TEXT,
        cliente TEXT,
        cliente_id INTEGER,
        tipo_trabajo TEXT DEFAULT 'REPARACION',
        fecha TEXT,
        creado_en {dt}
    )
    """)

    # 4. Instrumentos (Ítems de cada envío)
    cur.execute(f"""
    CREATE TABLE IF NOT EXISTS instrumentos (
        id {pk},
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
        creado_en {dt},
        actualizado_en {dt},
        foto_entrada_1 TEXT,
        foto_entrada_2 TEXT,
        foto_salida_1 TEXT,
        foto_salida_2 TEXT,
        tecnico_reparacion TEXT,
        tecnico_reparacion_en TEXT
    )
    """)

    # Migración: asegurar que existen foto_salida_1 y foto_salida_2
    cols_inst = get_table_columns(cur, "instrumentos")
    if "foto_salida_1" not in cols_inst:
        cur.execute("ALTER TABLE instrumentos ADD COLUMN foto_salida_1 TEXT")
    if "foto_salida_2" not in cols_inst:
        cur.execute("ALTER TABLE instrumentos ADD COLUMN foto_salida_2 TEXT")

    # 5. Permisos granulares
    cur.execute(f"""
    CREATE TABLE IF NOT EXISTS user_permissions (
      user_id INTEGER NOT NULL,
      action TEXT NOT NULL,
      allowed INTEGER NOT NULL DEFAULT 0,
      PRIMARY KEY (user_id, action)
    )
    """)

    # 6. Catálogo de Checklist
    cur.execute(f"""
    CREATE TABLE IF NOT EXISTS checklist_items (
        id {pk},
        nombre TEXT NOT NULL,
        orden INTEGER DEFAULT 0,
        activo INTEGER DEFAULT 1,
        tipo_trabajo TEXT DEFAULT 'REPARACION'
    )
    """)

    # 7. Checklist por Instrumento
    cur.execute(f"""
    CREATE TABLE IF NOT EXISTS instrumento_checklist (
        instrumento_id INTEGER NOT NULL,
        item_id INTEGER NOT NULL,
        hecho INTEGER DEFAULT 0,
        hecho_por TEXT,
        hecho_en TEXT,
        PRIMARY KEY (instrumento_id, item_id)
    )
    """)

    # 7b. Tabla de Informes PDF general
    cur.execute("""
    CREATE TABLE IF NOT EXISTS instrumento_informes (
        id {pk},
        instrumento_id INTEGER NOT NULL,
        filename TEXT,
        path TEXT,
        filepath TEXT,
        uploaded_at TEXT,
        uploaded_by TEXT
    )
    """.format(pk=pk))

    # 8. Control de Calidad Ópticas Rígidas
    cur.execute(f"""
    CREATE TABLE IF NOT EXISTS instrumento_qc_optica (
        instrumento_id INTEGER PRIMARY KEY,
        parte_trabajo_cliente TEXT,
        observaciones_cliente TEXT,
        observaciones_previas TEXT,
        
        -- Diagnostico choices: CORRECTO, INCORRECTO, SUSTITUCION, REPARACION
        diag_ventana TEXT,
        diag_fibra TEXT,
        diag_objetivo TEXT,
        diag_lentes TEXT,
        diag_camisa TEXT,
        diag_ocular TEXT,
        diag_pieza_ojo TEXT,
        diag_contaminacion TEXT,
        
        reparable INTEGER DEFAULT 1,
        
        -- Datos técnicos
        campo_vision_val TEXT,
        campo_vision_ok INTEGER DEFAULT 1,
        direccion_vision_val TEXT,
        direccion_vision_ok INTEGER DEFAULT 1,
        resolucion_val TEXT,
        resolucion_ok INTEGER DEFAULT 1,
        desviacion_val TEXT,
        desviacion_ok INTEGER DEFAULT 1,
        luz_val TEXT,
        luz_ok INTEGER DEFAULT 1,
        
        observaciones_finales TEXT,
        fecha_salida TEXT,
        firma_tecnico TEXT,
        firma_responsable TEXT,
        
        -- Fotos específicas para el informe CC
        qc_foto_entrada_1 TEXT,
        qc_foto_entrada_2 TEXT,
        qc_foto_salida_1 TEXT,
        qc_foto_salida_2 TEXT,
        
        creado_en {dt}
    )
    """)

    # Migración: asegurar columnas en instrumento_qc_optica
    cols_qc = get_table_columns(cur, "instrumento_qc_optica")
    diag_items = ['ventana', 'fibra', 'objetivo', 'lentes', 'camisa', 'ocular', 'pieza_ojo', 'contaminacion']
    for item in diag_items:
        if f"diag_{item}_estado" not in cols_qc:
            cur.execute(f"ALTER TABLE instrumento_qc_optica ADD COLUMN diag_{item}_estado TEXT")
        if f"diag_{item}_accion" not in cols_qc:
            cur.execute(f"ALTER TABLE instrumento_qc_optica ADD COLUMN diag_{item}_accion TEXT")

    for col in ["qc_foto_entrada_1", "qc_foto_entrada_2", "qc_foto_salida_1", "qc_foto_salida_2"]:
        if col not in cols_qc:
            cur.execute(f"ALTER TABLE instrumento_qc_optica ADD COLUMN {col} TEXT")

    conn.commit()
    conn.close()
