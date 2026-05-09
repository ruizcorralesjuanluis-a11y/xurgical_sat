import os
import sqlite3
import threading
from pathlib import Path
try:
    from psycopg2.pool import ThreadedConnectionPool
    import psycopg2.extras
    import psycopg2
except ImportError:
    ThreadedConnectionPool = None
    psycopg2 = None
from fastapi import HTTPException

import sys

# -------------------------------------------------
# RUTA BASE DEL PROYECTO
# -------------------------------------------------
if getattr(sys, 'frozen', False):
    # Si estamos en un ejecutable (PyInstaller), BASE_DIR es la carpeta del .exe
    BASE_DIR = Path(sys.executable).resolve().parent
    INTERNAL_DIR = Path(getattr(sys, "_MEIPASS", str(BASE_DIR)))
else:
    # Si estamos ejecutando con python normal
    BASE_DIR = Path(__file__).resolve().parent
    INTERNAL_DIR = BASE_DIR

# -------------------------------------------------
# RUTA DE LA BASE DE DATOS (PRO o LOCAL)
# -------------------------------------------------
DATABASE_URL = os.environ.get("DATABASE_URL")

# Si estamos en modo ejecutable (Local Agent), forzamos base de datos LOCAL (sat.db)
# para evitar problemas de permisos con rutas de red (UNC) como \\NONI\...
if getattr(sys, 'frozen', False):
    DB_PATH = BASE_DIR / "sat.db"
    if os.environ.get("XURGICAL_DB_PATH"):
         print(f">>> [DATABASE] AVISO: Ignorando XURGICAL_DB_PATH en modo ejecutable para usar DB local: {DB_PATH}")
else:
    DB_PATH = Path(os.environ.get("XURGICAL_DB_PATH", str(BASE_DIR / "sat.db")))

print(f">>> [DATABASE] BASE_DIR: {BASE_DIR}", flush=True)
print(f">>> [DATABASE] DB_PATH: {DB_PATH}", flush=True)

# -------------------------------------------------
# WRAPPER PARA COMPATIBILIDAD (SQLite vs Postgres)
# -------------------------------------------------
class PGRowWrapper(dict):
    """Permite acceso por atributo o por clave como sqlite3.Row"""
    def __getattr__(self, name):
        return self.get(name)

# Variable global para el pool de conexiones
_pool = None
_pool_lock = threading.Lock()

def get_pool():
    global _pool
    env_url = os.environ.get("DATABASE_URL")
    if not env_url:
        return None
    
    with _pool_lock:
        if _pool is None:
            if psycopg2 is None or ThreadedConnectionPool is None:
                print(">>> [DATABASE] AVISO: psycopg2 no instalado. No se puede usar Postgres.")
                return None
            url = env_url.strip()
            if url.startswith("postgres://"):
                url = url.replace("postgres://", "postgresql://", 1)
            # Creamos un pool de 1 a 10 conexiones
            _pool = ThreadedConnectionPool(1, 15, url, connect_timeout=10)
            print(f">>> [DATABASE] Pool de conexiones inicializado (1-15 conexiones)", flush=True)
        return _pool

class PGCursorWrapper:
    def __init__(self, cur):
        self.cur = cur
    
    def execute(self, sql, params=None):
        # Convertimos ? a %s para Postgres
        if params is not None:
            sql = sql.replace('?', '%s')
        # Limpieza de SQL específico de SQLite
        sql = sql.replace("BEGIN IMMEDIATE", "BEGIN")
        sql = sql.replace("datetime('now')", "CURRENT_TIMESTAMP")
        return self.cur.execute(sql, params)

    def executemany(self, sql, seq_params):
        # Convertimos ? a %s para Postgres
        sql = sql.replace('?', '%s')
        return self.cur.executemany(sql, seq_params)

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
    def __init__(self, conn, pool=None):
        self.conn = conn
        self.pool = pool
    def cursor(self):
        if psycopg2 and hasattr(psycopg2, "extras"):
            return PGCursorWrapper(self.conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor))
        return PGCursorWrapper(self.conn.cursor())
    def commit(self):
        self.conn.commit()
    def rollback(self):
        self.conn.rollback()
    def close(self):
        if self.pool:
            self.pool.putconn(self.conn)
        else:
            self.conn.close()

# -------------------------------------------------
# CONEXIÓN
# -------------------------------------------------
def get_connection():
    # Intentamos obtener la URL del entorno de nuevo por si se cargó tarde
    env_url = os.environ.get("DATABASE_URL")
    
    if env_url:
        try:
            pool = get_pool()
            if pool:
                conn = pool.getconn()
                return PGConnWrapper(conn, pool=pool)
        except Exception as e:
            print(f">>> [DATABASE] ERROR: No se pudo obtener conexión del pool. Reintentando conexión directa... {e}", flush=True)
            try:
                # Si falla el pool, intentamos conexión directa una vez
                if psycopg2 is None: raise ImportError("psycopg2 no disponible")
                url = env_url.strip().replace("postgres://", "postgresql://", 1)
                conn = psycopg2.connect(url, connect_timeout=10)
                return PGConnWrapper(conn)
            except Exception as e2:
                print(f">>> [DATABASE] ERROR: No se pudo conectar a Postgres de forma directa. Error: {e2}", flush=True)
                # Si falla Postgres, no hacemos fallback silencioso a SQLite si la URL existe
                raise HTTPException(status_code=500, detail=f"Error de conexión a la base de datos segura: {e2}")
    
    # Fallback a SQLite (solo si NO hay DATABASE_URL configurada)
    print(">>> [DATABASE] AVISO: Usando SQLite (Temporal). Configure DATABASE_URL en Render para persistencia.", flush=True)
    try:
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
        cliente_id INTEGER,
        created_at {dt}
    )
    """)

    # Migración: asegurar que existe cliente_id en users
    cols_users = get_table_columns(cur, "users")
    if "cliente_id" not in cols_users:
        cur.execute("ALTER TABLE users ADD COLUMN cliente_id INTEGER")

    # 2. Clientes
    cur.execute(f"""
    CREATE TABLE IF NOT EXISTS clientes (
        id {pk},
        numero_cliente INTEGER,
        nombre TEXT UNIQUE NOT NULL,
        prefijo TEXT,
        prefijo_nombre TEXT,
        email TEXT,
        ultimo_numero INTEGER DEFAULT 0
    )
    """)

    # Migración: asegurar que existe email y numero_cliente en clientes
    cols_clientes = get_table_columns(cur, "clientes")
    if "email" not in cols_clientes:
        cur.execute("ALTER TABLE clientes ADD COLUMN email TEXT")
    if "numero_cliente" not in cols_clientes:
        cur.execute("ALTER TABLE clientes ADD COLUMN numero_cliente INTEGER")

    # Retro-compatibilidad: auto-numerar clientes antiguos que tengan numero_cliente a NULL
    cur.execute("SELECT id FROM clientes WHERE numero_cliente IS NULL ORDER BY id ASC")
    filas_sin_numero = cur.fetchall()
    if filas_sin_numero:
        cur.execute("SELECT MAX(numero_cliente) as vmax FROM clientes WHERE numero_cliente IS NOT NULL")
        row_vmax = cur.fetchone()
        vmax_actual = 0
        if row_vmax and getattr(row_vmax, "vmax", None) is not None:
            vmax_actual = int(row_vmax.vmax)
        elif row_vmax and (isinstance(row_vmax, dict) or hasattr(row_vmax, "keys")) and getattr(row_vmax, "get", None) and row_vmax.get("vmax") is not None:
            vmax_actual = int(row_vmax.get("vmax"))
        else:
            try:
                if row_vmax and row_vmax["vmax"] is not None:
                    vmax_actual = int(row_vmax["vmax"])
            except:
                pass

        for f_cli in filas_sin_numero:
            vmax_actual += 1
            cur.execute("UPDATE clientes SET numero_cliente=? WHERE id=?", (vmax_actual, getattr(f_cli, "id", f_cli.get("id", f_cli["id"]) if isinstance(f_cli, dict) else f_cli[0])))


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
        aceptado INTEGER DEFAULT 0,
        aviso_enviado INTEGER DEFAULT 0,
        creado_en {dt},
        observaciones TEXT
    )
    """)

    # Migración: asegurar que existe aceptado y observaciones en envios
    cols_envios = get_table_columns(cur, "envios")
    if "aceptado" not in cols_envios:
        cur.execute("ALTER TABLE envios ADD COLUMN aceptado INTEGER DEFAULT 0")
    if "aviso_enviado" not in cols_envios:
        cur.execute("ALTER TABLE envios ADD COLUMN aviso_enviado INTEGER DEFAULT 0")
    if "observaciones" not in cols_envios:
        cur.execute("ALTER TABLE envios ADD COLUMN observaciones TEXT")

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
        foto_entrada_3 TEXT,
        foto_entrada_4 TEXT,
        foto_entrada_5 TEXT,
        foto_entrada_6 TEXT,
        foto_salida_1 TEXT,
        foto_salida_2 TEXT,
        foto_salida_3 TEXT,
        foto_salida_4 TEXT,
        foto_salida_5 TEXT,
        foto_salida_6 TEXT,
        tecnico_reparacion TEXT,
        tecnico_reparacion_en TEXT,
        repuesto_info TEXT,
        repuesto_precio REAL,
        revisado INTEGER DEFAULT 0,
        revisado_por TEXT,
        revisado_en TEXT,
        recomendada_sustitucion INTEGER DEFAULT 0,
        codigo_cliente TEXT
    )
    """)

    # Migración: asegurar columnas en instrumentos
    cols_instrumentos = get_table_columns(cur, "instrumentos")
    if "tecnico_reparacion" not in cols_instrumentos:
        cur.execute("ALTER TABLE instrumentos ADD COLUMN tecnico_reparacion TEXT")
    if "tecnico_reparacion_en" not in cols_instrumentos:
        cur.execute("ALTER TABLE instrumentos ADD COLUMN tecnico_reparacion_en TEXT")
    if "repuesto_info" not in cols_instrumentos:
        cur.execute("ALTER TABLE instrumentos ADD COLUMN repuesto_info TEXT")
    if "repuesto_precio" not in cols_instrumentos:
        cur.execute("ALTER TABLE instrumentos ADD COLUMN repuesto_precio REAL")
    if "recomendada_sustitucion" not in cols_instrumentos:
        cur.execute("ALTER TABLE instrumentos ADD COLUMN recomendada_sustitucion INTEGER DEFAULT 0")

    cur.execute(f"""
    CREATE TABLE IF NOT EXISTS articulos_catalogo (
        id {pk},
        codigo TEXT UNIQUE NOT NULL,
        descripcion TEXT,
        fabricante TEXT
    )
    """)
    cur.execute("CREATE INDEX IF NOT EXISTS idx_articulos_codigo ON articulos_catalogo(codigo)")

    cur.execute(f"""
    CREATE TABLE IF NOT EXISTS repuestos_catalogo (
        id {pk},
        nombre TEXT NOT NULL,
        precio REAL DEFAULT 0,
        activo INTEGER DEFAULT 1,
        creado_en {dt}
    )
    """)

    cur.execute(f"""
    CREATE TABLE IF NOT EXISTS instrumento_repuestos (
        instrumento_id INTEGER NOT NULL,
        repuesto_id INTEGER NOT NULL,
        precio_aplicado REAL,
        cantidad INTEGER DEFAULT 1,
        PRIMARY KEY (instrumento_id, repuesto_id)
    )
    """)

    # Migración: asegurar que existe la columna cantidad
    cols_ir = get_table_columns(cur, "instrumento_repuestos")
    if "cantidad" not in cols_ir:
        try:
            cur.execute("ALTER TABLE instrumento_repuestos ADD COLUMN cantidad INTEGER DEFAULT 1")
        except Exception:
            if is_pg: conn.rollback()

    # Migración: asegurar que existen las columnas de fotos (1-6) tanto entrada como salida
    cols_inst = get_table_columns(cur, "instrumentos")
    if "repuesto_info" not in cols_inst:
        cur.execute("ALTER TABLE instrumentos ADD COLUMN repuesto_info TEXT")
    if "repuesto_precio" not in cols_inst:
        cur.execute("ALTER TABLE instrumentos ADD COLUMN repuesto_precio REAL")
    if "revisado" not in cols_inst:
        cur.execute("ALTER TABLE instrumentos ADD COLUMN revisado INTEGER DEFAULT 0")
    if "revisado_por" not in cols_inst:
        cur.execute("ALTER TABLE instrumentos ADD COLUMN revisado_por TEXT")
    if "revisado_en" not in cols_inst:
        cur.execute("ALTER TABLE instrumentos ADD COLUMN revisado_en TEXT")
    if "recomendada_sustitucion" not in cols_inst:
        cur.execute("ALTER TABLE instrumentos ADD COLUMN recomendada_sustitucion INTEGER DEFAULT 0")
    if "codigo_cliente" not in cols_inst:
        cur.execute("ALTER TABLE instrumentos ADD COLUMN codigo_cliente TEXT")
    for i in range(1, 7):
        col_e = f"foto_entrada_{i}"
        if col_e not in cols_inst:
            cur.execute(f"ALTER TABLE instrumentos ADD COLUMN {col_e} TEXT")
        col_s = f"foto_salida_{i}"
        if col_s not in cols_inst:
            cur.execute(f"ALTER TABLE instrumentos ADD COLUMN {col_s} TEXT")

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

    # 9. ÍNDICES DE RENDIMIENTO (OPTIMIZADOS)
    # Estos índices aceleran drásticamente las búsquedas en tablas con muchos datos
    # Aceleramos el Dashboard y filtrados por cliente/estado
    cur.execute("CREATE INDEX IF NOT EXISTS idx_instrumentos_envio_id ON instrumentos(envio_id)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_instrumentos_estado ON instrumentos(estado)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_envios_cliente_id ON envios(cliente_id)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_envios_ot_num ON envios(ot_num)")
    
    # Aceleramos la búsqueda de trazabilidad (DM y S/N)
    cur.execute("CREATE INDEX IF NOT EXISTS idx_instrumentos_dm ON instrumentos(codigo_datamatrix)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_instrumentos_sn ON instrumentos(num_serie)")
    
    # Índices para checklist e informes
    cur.execute("CREATE INDEX IF NOT EXISTS idx_checklist_instrumento_id ON instrumento_checklist(instrumento_id)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_informes_instrumento_id ON instrumento_informes(instrumento_id)")
    
    if is_pg:
        # Postgres específicos si hiciera falta (de momento con los de arriba sobra)
        pass

    # 10. PETICIONES DE RECOGIDA
    cur.execute(f"""
    CREATE TABLE IF NOT EXISTS peticiones_recogida (
        id {pk},
        num_peticion TEXT UNIQUE,
        cliente_id INTEGER NOT NULL,
        usuario_id INTEGER,
        n_instrumentos INTEGER DEFAULT 0,
        contacto TEXT,
        telefono TEXT,
        observaciones TEXT,
        estado TEXT DEFAULT 'Pendiente',
        creado_en {dt}
    )
    """)

    # 11. CONSULTAS TÉCNICAS (NUEVO)
    cur.execute(f"""
    CREATE TABLE IF NOT EXISTS consultas (
        id {pk},
        cliente_id INTEGER NOT NULL,
        usuario_id INTEGER NOT NULL,
        titulo TEXT,
        descripcion TEXT,
        estado TEXT DEFAULT 'Abierta', -- Abierta, Respondida, Cerrada
        foto_1 TEXT,
        foto_2 TEXT,
        foto_3 TEXT,
        creado_en {dt},
        actualizado_en {dt}
    )
    """)

    # 12. MENSAJES DE CONSULTAS (CHAT)
    cur.execute(f"""
    CREATE TABLE IF NOT EXISTS consultas_mensajes (
        id {pk},
        consulta_id INTEGER NOT NULL,
        usuario_id INTEGER NOT NULL,
        mensaje TEXT,
        creado_en {dt},
        FOREIGN KEY (consulta_id) REFERENCES consultas(id) ON DELETE CASCADE
    )
    """)

    # Migraciones para peticiones_recogida
    cols_pr = get_table_columns(cur, "peticiones_recogida")
    if "contacto" not in cols_pr:
        cur.execute("ALTER TABLE peticiones_recogida ADD COLUMN contacto TEXT")
    if "telefono" not in cols_pr:
        cur.execute("ALTER TABLE peticiones_recogida ADD COLUMN telefono TEXT")
    if "num_peticion" not in cols_pr:
        cur.execute("ALTER TABLE peticiones_recogida ADD COLUMN num_peticion TEXT")
        cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_peticiones_num ON peticiones_recogida(num_peticion)")

    # Migración: Vincular automáticamente envíos con cliente_id si coinciden por nombre
    cur.execute("""
        UPDATE envios 
        SET cliente_id = (SELECT id FROM clientes WHERE clientes.nombre = envios.cliente)
        WHERE cliente_id IS NULL 
          AND (SELECT id FROM clientes WHERE clientes.nombre = envios.cliente) IS NOT NULL
    """)

    conn.commit()
    conn.close()
