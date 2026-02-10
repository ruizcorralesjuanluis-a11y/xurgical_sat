
def _clean_trz(s):
    if not s:
        return ""
    s = str(s).strip()
    # quita backslashes que vienen escapando comillas
    s = s.replace("\\'", "").replace('\\"', "").replace("\\", "")
    # quita comillas envolventes repetidas
    while len(s) >= 2 and ((s[0] == "'" and s[-1] == "'") or (s[0] == '"' and s[-1] == '"')):
        s = s[1:-1].strip()
    # quita comillas sueltas en extremos
    s = s.strip().strip("'").strip('"').strip()
    return s

import re
import time
import uuid

def _last5_digits_from_dm(dm: str) -> str:
    """Extrae los últimos 5 dígitos del DataMatrix. Si hay menos de 5 dígitos, usa los últimos 5 caracteres."""
    s = (dm or "").strip()
    digits = re.findall(r"\d", s)
    if len(digits) >= 5:
        return "".join(digits[-5:])
    s2 = re.sub(r"\s+", "", s)
    return (s2[-5:] if len(s2) >= 5 else s2)

def _build_nombre_trazabilidad(prefijo_nombre: str, codigo_datamatrix: str) -> str:
    pref = (prefijo_nombre or "").strip()
    suf = _last5_digits_from_dm(codigo_datamatrix)
    if not pref and not suf:
        return ""
    return f"{pref}{suf}"

import os
import base64
import sqlite3
import io
import csv
from datetime import datetime

# Etiquetas (PDF) + código de barras / QR
from reportlab.pdfgen import canvas
from reportlab.graphics.barcode import code128, qr
from reportlab.lib.units import mm

# Para autocompletar artículos (Articulos.xlsx)
try:
    import pandas as pd
except Exception:  # pragma: no cover
    pd = None

try:
    from openpyxl import load_workbook
except Exception:  # pragma: no cover
    load_workbook = None

from fastapi import FastAPI, Request, UploadFile, File, Form, Depends, HTTPException
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import sys
from pathlib import Path

# Rutas compatibles con PyInstaller (sys._MEIPASS) y ejecución normal
BASE_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))

from db import init_db, get_conn
from excel_import import leer_excel_envio

from auth_utils import verify_password, hash_password, make_serializer, sign_session
from security import get_current_user, require_roles

try:
    from openpyxl import Workbook
except Exception:  # pragma: no cover
    Workbook = None

app = FastAPI(title="Xurgical SAT")

# Asegura el esquema incluso si el servidor no dispara el evento startup
# (algunas ejecuciones/entornos pueden evitarlo en ciertas rutas).
try:
    init_db()
except Exception:
    pass

app.state.secret_key = os.environ.get("XURGICAL_SECRET_KEY", "dev-secret-change-me")
app.state.serializer = make_serializer(app.state.secret_key)

UPLOAD_DIR = os.environ.get("XURGICAL_UPLOAD_DIR", "uploads")
# Priorizamos que las fotos estén dentro del disco persistente si estamos en Render
FOTOS_DIR = os.environ.get("XURGICAL_FOTOS_DIR", os.path.join(UPLOAD_DIR, "fotos"))

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(FOTOS_DIR, exist_ok=True)

# Monta FOTOS_DIR en /static/fotos PRIORITARIAMENTE para soportar discos externos
app.mount("/static/fotos", StaticFiles(directory=FOTOS_DIR), name="fotos_externas")
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")

templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))

def format_fecha(value):
    if not value or value == "-":
        return "-"
    try:
        # Detectar si es un string y tratar de parsearlo
        s = str(value).strip()
        if not s: return "-"
        
        # Intentar varios formatos comunes
        dt = None
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
            try:
                dt = datetime.strptime(s, fmt)
                break
            except:
                continue
        
        if not dt:
            try:
                dt = datetime.fromisoformat(s.replace("Z", "+00:00"))
            except:
                return s # Si falla todo, devuelve el original
        
        return dt.strftime("%d-%m-%Y")
    except:
        return value

templates.env.filters["fecha"] = format_fecha



# -----------------------------
# CLIENTES (prefijo + contador)
# -----------------------------
def _list_clientes(cur) -> list[dict]:
    """Listado de clientes.

    Nota: en instalaciones existentes puede faltar la tabla por no haberse ejecutado
    aún init_db() sobre una BD antigua. Recuperamos automáticamente.
    """
    # Detectar si estamos en Postgres o SQLite para capturar errores de tabla inexistente
    import sqlite3
    try:
        import psycopg2
        PG_ERR = psycopg2.Error
    except ImportError:
        PG_ERR = Exception

    sql = "SELECT id, nombre, prefijo, prefijo_nombre, ultimo_numero FROM clientes ORDER BY LOWER(nombre) ASC"
    try:
        cur.execute(sql)
        return [dict(r) for r in cur.fetchall()]
    except (sqlite3.OperationalError, PG_ERR) as e:
        err_msg = str(e).lower()
        if "no such table" in err_msg or "does not exist" in err_msg:
            init_db()
            cur.execute(sql)
            return [dict(r) for r in cur.fetchall()]
        raise


def _envios_has_column(cur, col: str) -> bool:
    try:
        from db import get_table_columns
        cols = get_table_columns(cur, "envios")
        return col in cols
    except Exception:
        return False


def _get_cliente(cur, cliente_id: int) -> dict | None:
    cur.execute(
        "SELECT id, nombre, prefijo, prefijo_nombre, ultimo_numero FROM clientes WHERE id=?",
        (int(cliente_id),),
    )
    r = cur.fetchone()
    return dict(r) if r else None


def _reserve_numeros_cliente(cur, cliente_id: int, cantidad: int) -> tuple[str, str, list[int]]:
    """Reserva un rango de numeros (por cliente) de forma transaccional.

    Devuelve: (prefijo_dm, prefijo_nombre, [n1, n2, ...])
    """
    if cantidad <= 0:
        return "", []

    # Bloqueo transaccional (en sqlite3, basta con ejecutar dentro del mismo conn
    # con BEGIN IMMEDIATE para evitar carreras).
    cur.execute("BEGIN IMMEDIATE")

    cur.execute("SELECT prefijo, prefijo_nombre, ultimo_numero FROM clientes WHERE id=?", (int(cliente_id),))
    row = cur.fetchone()
    if not row:
        raise HTTPException(status_code=400, detail="Cliente no encontrado")

    row = dict(row)

    prefijo_dm = (row["prefijo"] or "").strip()
    prefijo_nombre = (row.get("prefijo_nombre") or "").strip()
    ultimo = int(row["ultimo_numero"] or 0)
    nums = list(range(ultimo + 1, ultimo + cantidad + 1))
    cur.execute("UPDATE clientes SET ultimo_numero=? WHERE id=?", (nums[-1], int(cliente_id)))

    return prefijo_dm, prefijo_nombre, nums

# moved to top

# -----------------------------
# ARTICULOS (catálogo para autocompletar)
# -----------------------------
ARTICULOS_XLSX = str(BASE_DIR / 'Articulos.xlsx')
ARTICULOS_XLS  = str(BASE_DIR / 'Articulos.xls')
_articulos_map = None  # dict normalizado: codigo -> {descripcion, fabricante?}

def _norm_codigo(x: str) -> str:
    """Normaliza códigos de artículo.

    Reglas:
    - Trim + upper (búsqueda insensible a mayúsculas)
    - Elimina prefijos del catálogo (p.ej. RP, MT), ignorando mayúsculas/minúsculas para que el usuario pueda
      buscar sin ellos y el sistema siempre devuelva el código "limpio".
    """
    s = (x or '').strip().upper()
    # En el Excel algunos códigos vienen con prefijos como RP/MT.
    # Los quitamos para unificar y evitar que el usuario tenga que escribirlos.
    for pfx in ("RP", "MT"):
        if s.startswith(pfx):
            s = s[len(pfx):]
            break
    # Quita separadores habituales después del prefijo (espacio, guión, etc.)
    s = s.lstrip(" -_\t")
    return s


def _codigo_variants(codigo_norm: str) -> list:
    """Genera variantes de búsqueda para un código normalizado.

    En el catálogo hay códigos que son equivalentes con o sin una "R" final.
    Esta función devuelve una lista de claves candidatas para encontrar el artículo.
    """
    c = (codigo_norm or '').strip().upper()
    if not c:
        return []

    out = []
    def add(x: str):
        x = (x or '').strip().upper()
        if x and x not in out:
            out.append(x)

    add(c)

    # Variante sin espacios internos
    add(c.replace(' ', ''))

    # Equivalencia con/sin R final
    if c.endswith('R') and len(c) > 1:
        add(c[:-1])
        add(c[:-1].replace(' ', ''))
    else:
        add(c + 'R')
        add((c + 'R').replace(' ', ''))

    return out

def load_articulos_map() -> dict:
    """Carga el catálogo de Articulos para autocompletar."""
    global _articulos_map
    if _articulos_map is not None:
        return _articulos_map

    path = ARTICULOS_XLS if os.path.exists(ARTICULOS_XLS) else ARTICULOS_XLSX
    if not os.path.exists(path):
        print(f">>> [ARTICULOS] Archivo no encontrado: {path}", flush=True)
        _articulos_map = {}
        return _articulos_map

    def _put(m, code, desc, fab=None):
        code = _norm_codigo(str(code))
        if not code or code == "NAN":
            return
        desc = ("" if desc is None else str(desc)).strip()
        if not desc or desc.lower() == "nan":
            return
        fab_s = None
        if fab is not None:
            fab_s = str(fab).strip()
            if not fab_s or fab_s.lower() == "nan":
                fab_s = None

        if fab_s is None:
            try:
                d = desc.strip()
                if d.upper().endswith(" AS"):
                    fab_s = "Aescula"
            except Exception:
                pass
        m[code] = {"descripcion": desc, "fabricante": fab_s}
        if code.endswith('R') and len(code) > 1:
            code2 = code[:-1].strip()
            if code2 and code2 not in m:
                m[code2] = {"descripcion": desc, "fabricante": fab_s}

    m = {}
    df = None
    if pd is not None:
        try:
            engine = 'openpyxl' if path.lower().endswith('.xlsx') else None
            df = pd.read_excel(path, engine=engine)
            print(f">>> [ARTICULOS] Cargado {path} ({len(df)} filas)", flush=True)
        except Exception as e:
            print(f">>> [ARTICULOS] Error {path}: {e}", flush=True)
            if path.lower().endswith('.xls') and os.path.exists(ARTICULOS_XLSX):
                try:
                    path = ARTICULOS_XLSX
                    df = pd.read_excel(path, engine='openpyxl')
                except Exception:
                    pass

    if df is not None:
        cols_raw = [str(c).strip() for c in df.columns]
        cols_map = {c.lower(): c for c in cols_raw}
        col_codigo = cols_map.get("código") or cols_map.get("codigo") or cols_map.get("cod") or (cols_raw[0] if cols_raw else None)
        col_desc = cols_map.get("descripción") or cols_map.get("descripcion") or cols_map.get("denominacion") or (cols_raw[1] if len(cols_raw) >= 2 else None)
        col_fab = None
        if len(cols_raw) >= 4:
            col_fab = cols_raw[3]
        else:
            col_fab = cols_map.get("fabricante") or cols_map.get("marca") or cols_map.get("manufacturer")

        if col_codigo and col_desc:
            for _, row in df.iterrows():
                try:
                    _put(m, row[col_codigo], row[col_desc], row[col_fab] if col_fab else None)
                except Exception:
                    continue
    _articulos_map = m
    return _articulos_map

# -----------------------------
# Permisos por acción (además de roles)
# -----------------------------
ACTIONS = [
    ("dashboard_ver", "Ver dashboard"),
    ("envio_crear", "Crear parte"),
    ("envio_borrar", "Borrar parte"),
    ("instrumento_crear", "Añadir instrumento"),
    ("instrumento_editar", "Editar instrumento"),
    ("instrumento_borrar", "Borrar instrumento"),
    ("fotos_gestionar", "Subir/Borrar fotos"),
    ("estado_cambiar", "Cambiar estado instrumento"),
    ("excel_importar", "Importar Excel"),
    ("usuarios_gestionar", "Gestionar usuarios/roles"),
]

def _user_role(user):
    """Devuelve el role de user soportando dict/sqlite3.Row/objeto."""
    if user is None:
        return None
    # dict-like
    try:
        if hasattr(user, "keys"):
            return user.get("role") if hasattr(user, "get") else user["role"]
    except Exception:
        pass
    # objeto
    try:
        return getattr(user, "role", None)
    except Exception:
        return None

def _user_id(user):
    """Devuelve el id de user soportando dict/sqlite3.Row/objeto."""
    if user is None:
        return None
    try:
        if hasattr(user, "keys"):
            return int(user.get("id") if hasattr(user, "get") else user["id"])
    except Exception:
        pass
    try:
        return int(getattr(user, "id", None))
    except Exception:
        return None

def _default_allowed_by_role(role: str, action: str) -> int:
    if role == "admin":
        return 1
    if role == "recepcion":
        return 1 if action in {"dashboard_ver","envio_crear","instrumento_crear","fotos_gestionar","excel_importar"} else 0
    if role == "tecnico":
        return 1 if action in {"dashboard_ver","instrumento_editar","fotos_gestionar","estado_cambiar"} else 0
    if role == "grabado":
        return 1 if action in {"dashboard_ver"} else 0
    return 0

def _get_user_permissions_map(cur, user_id: int) -> dict:
    cur.execute("SELECT action, allowed FROM user_permissions WHERE user_id=?", (user_id,))
    return {r["action"]: int(r["allowed"] or 0) for r in cur.fetchall()}

def can_action(user, action: str, cur=None) -> bool:
    # Admin siempre
    if _user_role(user) == "admin":
        return True
    role = _user_role(user) or ""
    if cur is None or user is None:
        return _default_allowed_by_role(role, action) == 1
    try:
        cur.execute("SELECT allowed FROM user_permissions WHERE user_id=? AND action=?", (int(_user_id(user)), action))
        row = cur.fetchone()
        if row is None:
            return _default_allowed_by_role(role, action) == 1
        return int(row["allowed"] or 0) == 1
    except Exception:
        # Fallback to default if table missing or other DB error
        return _default_allowed_by_role(role, action) == 1


@app.on_event("startup")
def on_startup():
    init_db()

    # Tabla de permisos por acción (granularidad extra a roles)
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
    CREATE TABLE IF NOT EXISTS user_permissions (
      user_id INTEGER NOT NULL,
      action TEXT NOT NULL,
      allowed INTEGER NOT NULL DEFAULT 0,
      PRIMARY KEY (user_id, action),
      FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
    );
    """)
    conn.commit()
    conn.close()
    # -----------------------------
    # AUTO-RECOVERY ADMIN (solo para demo)
    # Si NO hay ningún admin activo, fuerza:
    #   usuario: admin
    #   password: admin123
    # -----------------------------
    conn = get_conn()
    cur = conn.cursor()

    # Detecta columnas reales en users
    from db import get_table_columns
    cols = get_table_columns(cur, "users")
    colset = set(cols)

    pw_col = "password_hash" if "password_hash" in colset else ("password" if "password" in colset else None)
    has_is_active = "is_active" in colset
    has_created_at = "created_at" in colset
    has_created_at_at = "created_at_at" in colset

    if pw_col:
        if has_is_active:
            cur.execute("SELECT COUNT(*) AS n FROM users WHERE role='admin' AND is_active=1")
        else:
            cur.execute("SELECT COUNT(*) AS n FROM users WHERE role='admin'")
        n_admin = int(cur.fetchone()["n"] or 0)

        if n_admin == 0:
            # No hay admin activo -> forzar admin/admin123
            admin_hash = hash_password("admin123")

            cur.execute("SELECT id FROM users WHERE username='admin'")
            row = cur.fetchone()
            if row:
                # Update existente
                if has_is_active:
                    cur.execute(f"UPDATE users SET role='admin', is_active=1, {pw_col}=? WHERE username='admin'", (admin_hash,))
                else:
                    cur.execute(f"UPDATE users SET role='admin', {pw_col}=? WHERE username='admin'", (admin_hash,))
            else:
                # Insert nuevo compatible con columnas
                cols_ins = ["username", pw_col, "role"]
                vals = ["admin", admin_hash, "admin"]

                if has_is_active:
                    cols_ins.append("is_active")
                    vals.append(1)

                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                if has_created_at:
                    cols_ins.append("created_at")
                    vals.append(now)
                elif has_created_at_at:
                    cols_ins.append("created_at_at")
                    vals.append(now)

                sql = f"INSERT INTO users ({', '.join(cols_ins)}) VALUES ({', '.join(['?']*len(cols_ins))})"
                cur.execute(sql, tuple(vals))

            conn.commit()

    conn.close()


@app.get("/init_db")
def manual_init_db():
    try:
        on_startup()
        return {"ok": True, "message": "Base de datos inicializada y admin creado (admin/admin123)"}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@app.get('/articulos_lookup')
def articulos_lookup(codigo: str, user=Depends(get_current_user)):
    """Lookup de artículo para autorrelleno en alta manual de instrumentos."""
    key = _norm_codigo(codigo)
    if not key:
        return {'found': False}

    m = load_articulos_map()
    if not m:
        return {'found': False, 'error': 'Catálogo Articulos no disponible'}
    
    for k in _codigo_variants(key):
        hit = m.get(k)
        if hit:
            resp = {
                'found': True,
                'codigo': k,
                'descripcion': (hit.get('descripcion') or '').strip(),
            }
            fab = (hit.get('fabricante') or '').strip()
            if fab:
                resp['fabricante'] = fab
            return resp

    return {'found': False}




@app.get('/articulos_lookup')
def articulos_lookup(codigo: str, user=Depends(get_current_user)):
    """Lookup de artículo para autorrelleno en alta manual de instrumentos.

    Devuelve JSON con:
      - found: bool
      - codigo: str (normalizado)
      - descripcion: str
      - fabricante: str (opcional)
    """
    key = _norm_codigo(codigo)
    if not key:
        return {'found': False}

    m = load_articulos_map()
    if not m:
        return {'found': False, 'error': 'Catálogo Articulos no disponible'}

    for k in _codigo_variants(key):
        hit = m.get(k)
        if hit:
            resp = {
                'found': True,
                'codigo': k,
                'descripcion': (hit.get('descripcion') or '').strip(),
            }
            fab = (hit.get('fabricante') or '').strip()
            if fab:
                resp['fabricante'] = fab
            return resp

    return {'found': False}


@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    import traceback
    return HTMLResponse(
        f"<h1>Error Global del Servidor</h1><pre>{traceback.format_exc()}</pre>",
        status_code=500
    )

@app.exception_handler(HTTPException)
async def http_exception_handler(request: Request, exc: HTTPException):
    if exc.status_code == 401:
        accept = request.headers.get("accept", "")
        if "text/html" in accept:
            return RedirectResponse(url="/login", status_code=303)
    return JSONResponse({"detail": exc.detail}, status_code=exc.status_code)


def _try_delete_public_photo(public_path: str | None) -> None:
    if not public_path:
        return
    if not public_path.startswith("/static/fotos/"):
        return
    filename = public_path.replace("/static/fotos/", "")
    path_fs = os.path.join(FOTOS_DIR, filename)
    try:
        if os.path.exists(path_fs):
            os.remove(path_fs)
    except Exception:
        pass


def _next_ot_num(cur) -> str:
    yy = datetime.now().year % 100
    prefix = f"{yy:02d}OT"
    cur.execute(
        "SELECT ot_num FROM envios WHERE ot_num LIKE ? ORDER BY ot_num DESC LIMIT 1;",
        (prefix + "%",),
    )
    row = cur.fetchone()
    max_seq = 0
    if row and row["ot_num"]:
        try:
            max_seq = int(row["ot_num"][-5:])
        except Exception:
            max_seq = 0
    return f"{prefix}{(max_seq + 1):05d}"


# -----------------------------
# AUTH
# -----------------------------
@app.get("/login", response_class=HTMLResponse)
def login_form(request: Request):
    return templates.TemplateResponse("login.html", {"request": request, "error": None})


@app.post("/login")
async def login(request: Request):
    form = await request.form()
    username = (form.get("username") or "").strip()
    password = form.get("password") or ""

    conn = get_conn()
    cur = conn.cursor()

    schema = _users_schema(cur)
    pw_col = schema.get("password_col")
    if not pw_col:
        conn.close()
        return templates.TemplateResponse("login.html", {"request": request, "error": "Credenciales inválidas"})

    if schema.get("has_is_active"):
        cur.execute(f"SELECT id, {pw_col} AS pw, is_active FROM users WHERE username=?", (username,))
    else:
        cur.execute(f"SELECT id, {pw_col} AS pw, 1 AS is_active FROM users WHERE username=?", (username,))

    u = cur.fetchone()
    conn.close()

    if not u:
        return templates.TemplateResponse("login.html", {"request": request, "error": "Credenciales inválidas"})

    if int(u["is_active"] or 0) == 0:
        return templates.TemplateResponse("login.html", {"request": request, "error": "Usuario desactivado (contacta con el administrador)"})

    if not verify_password(password, u["pw"]):
        return templates.TemplateResponse("login.html", {"request": request, "error": "Credenciales inválidas"})

    response = RedirectResponse(url="/", status_code=303)
    token = sign_session(app.state.serializer, user_id=int(u["id"]))
    response.set_cookie("xurgical_session", token, httponly=True, samesite="lax")
    return response




@app.post("/logout")
def logout():
    resp = RedirectResponse(url="/login", status_code=303)
    resp.delete_cookie("xurgical_session")
    return resp


# -----------------------------
# DASHBOARD
# -----------------------------
@app.get("/", response_class=HTMLResponse)
def dashboard(request: Request, user=Depends(get_current_user)):
    # Comprobamos el entorno en tiempo real para el indicador
    is_pg = bool(os.environ.get("DATABASE_URL"))
    is_disk = bool(os.environ.get("XURGICAL_DB_PATH"))
    
    if is_pg:
        db_type = "PRO PG" 
    elif is_disk:
        db_type = "PRO DISK"
    else:
        db_type = "TEMP"
    conn = get_conn()
    cur = conn.cursor()

    q = (request.query_params.get('q') or '').strip()

    # --- 1. KPIs Globales (Instrumentos) en UNA sola consulta ---
    cur.execute("""
        SELECT
            COUNT(*) as total,
            SUM(CASE WHEN estado='Pendiente' THEN 1 ELSE 0 END) as pendientes,
            SUM(CASE WHEN estado='En proceso' THEN 1 ELSE 0 END) as en_proceso,
            SUM(CASE WHEN estado='Reparado' THEN 1 ELSE 0 END) as reparados,
            SUM(CASE WHEN estado='Baja' THEN 1 ELSE 0 END) as baja
        FROM instrumentos
    """)
    row_kpi = cur.fetchone()
    if row_kpi:
        kpis = {
            "total": row_kpi["total"],
            "pendientes": row_kpi["pendientes"],
            "en_proceso": row_kpi["en_proceso"],
            "reparado": row_kpi["reparados"],
            "baja": row_kpi["baja"],
        }
    else:
        kpis = {"total":0, "pendientes":0, "en_proceso":0, "reparado":0, "baja":0}

    # --- 2. KPIs Partes (Abiertos/Cerrados) en UNA sola consulta ---
    cur.execute("SELECT COUNT(*) AS n FROM envios")
    total_partes = cur.fetchone()["n"]

    has_tipo = _envios_has_column(cur, "tipo_trabajo")
    tipo_col = "e.tipo_trabajo" if has_tipo else "'REPARACION'"
    
    # Esta consulta agrupa por envío y calcula si está 'cerrado' o no
    # Logica de cierre:
    #   - TRAZABILIDAD: cerrado si total > 0 Y todos grabados (grabado=1)
    #   - REPARACION: cerrado si total > 0 Y todos (Reparado o Baja)
    cur.execute(f"""
        SELECT 
            e.id,
            {tipo_col} as tipo,
            COUNT(i.id) as total_inst,
            SUM(CASE WHEN COALESCE(i.grabado,0)=1 THEN 1 ELSE 0 END) as n_grabados,
            SUM(CASE WHEN i.estado IN ('Reparado','Baja') THEN 1 ELSE 0 END) as n_terminados
        FROM envios e
        LEFT JOIN instrumentos i ON i.envio_id = e.id
        GROUP BY e.id, {tipo_col}
    """)
    
    abiertos = 0
    cerrados = 0
    
    rows = cur.fetchall() # Traemos todos de golpe (mucho más rápido que 1 query por row)
    for r in rows:
        t_inst = r["total_inst"]
        if t_inst == 0:
            abiertos += 1 # Parte vacío = abierto (o pendiente de hacer algo)
            continue
            
        tipo = (r["tipo"] or "REPARACION").upper()
        if tipo == "TRAZABILIDAD":
            # Cerrado si todos grabados
            if r["n_grabados"] == t_inst:
                cerrados += 1
            else:
                abiertos += 1
        else:
            # REPARACION / OPTICA
            # Cerrado si todos terminados
            if r["n_terminados"] == t_inst:
                cerrados += 1
            else:
                abiertos += 1

    kpis_partes = {"total": total_partes, "abiertos": abiertos, "cerrados": cerrados}

    # ✅ AÑADIDO: n_fotos_completas y n_con_alguna_foto para calcular el punto por parte
    # Compatibilidad: en BDs antiguas puede no existir e.tipo_trabajo. En ese caso asumimos REPARACION.
    select_tipo = "e.tipo_trabajo" if _envios_has_column(cur, "tipo_trabajo") else "'REPARACION' AS tipo_trabajo"

    # --- Buscador (OT / Cliente / DataMatrix) ---
    where_q = ""
    params_q: list = []
    if q:
        where_q = "WHERE (e.ot_num LIKE ? OR e.cliente LIKE ? OR COALESCE(i.codigo_datamatrix,'') LIKE ?)"
        params_q = [f"%{q}%", f"%{q}%", f"%{q}%"]

    cur.execute(f"""
        SELECT
            e.id, e.ot_num, e.nombre_archivo, e.cliente, e.fecha,
            {select_tipo},
            COUNT(i.id) AS n_instrumentos,
            SUM(CASE WHEN i.estado IN ('Pendiente','En proceso') THEN 1 ELSE 0 END) AS n_pendientes,
            SUM(CASE WHEN i.foto_entrada_1 IS NOT NULL AND i.foto_entrada_2 IS NOT NULL AND i.foto_entrada_3 IS NOT NULL AND i.foto_entrada_4 IS NOT NULL AND i.foto_entrada_5 IS NOT NULL AND i.foto_entrada_6 IS NOT NULL THEN 1 ELSE 0 END) AS n_fotos_completas,
            SUM(CASE WHEN i.foto_entrada_1 IS NOT NULL OR  i.foto_entrada_2 IS NOT NULL OR i.foto_entrada_3 IS NOT NULL OR i.foto_entrada_4 IS NOT NULL OR i.foto_entrada_5 IS NOT NULL OR i.foto_entrada_6 IS NOT NULL THEN 1 ELSE 0 END) AS n_con_alguna_foto,
            SUM(CASE WHEN COALESCE(i.grabado,0)=1 THEN 1 ELSE 0 END) AS n_grabados
        FROM envios e
        LEFT JOIN instrumentos i ON i.envio_id = e.id
        {where_q}
        GROUP BY e.id
        ORDER BY e.id DESC
        LIMIT 200
    """, params_q)

    envios = []
    for r in cur.fetchall():
        d = dict(r)

        # Cierre:
        # - REPARACION: cerrado si todos están en Reparado o Baja
        # - TRAZABILIDAD: cerrado si todos están grabados (grabado=1)
        if (d.get("tipo_trabajo") or "REPARACION") == "TRAZABILIDAD":
            cur.execute(
                """
                SELECT
                  COUNT(*) AS total,
                  SUM(CASE WHEN COALESCE(grabado,0)=1 THEN 1 ELSE 0 END) AS done
                FROM instrumentos
                WHERE envio_id=?
                """,
                (d["id"],),
            )
            rr = cur.fetchone()
            total = int(rr["total"] or 0) if rr else 0
            done = int(rr["done"] or 0) if rr else 0
            d["is_closed"] = (total > 0 and done == total)
            # En trazabilidad, los "pendientes" son los no grabados
            d["n_pendientes"] = max(total - done, 0)
        else:
            cur.execute(
                """
                SELECT
                  COUNT(*) AS total,
                  SUM(CASE WHEN COALESCE(estado,'') IN ('Reparado','Baja') THEN 1 ELSE 0 END) AS done
                FROM instrumentos
                WHERE envio_id=?
                """,
                (d["id"],),
            )
            rr = cur.fetchone()
            total = int(rr["total"] or 0) if rr else 0
            done = int(rr["done"] or 0) if rr else 0
            d["is_closed"] = (total > 0 and done == total)
            d["n_pendientes"] = max(total - done, 0)

        # Color de la OT en dashboard: verde si cerrada, rojo si abierta
        d["color"] = "green" if d.get("is_closed") else "red"

        n_inst = int(d.get("n_instrumentos") or 0)
        n_fotos_completas = int(d.get("n_fotos_completas") or 0)
        n_con_alguna_foto = int(d.get("n_con_alguna_foto") or 0)

        # 🔴 ninguno tiene fotos
        if n_inst == 0 or n_con_alguna_foto == 0:
            d["foto_dot"] = "red"
        # 🟢 todos los instrumentos tienen las 2 fotos
        elif n_fotos_completas == n_inst:
            d["foto_dot"] = "green"
        # 🟡 mezcla
        else:
            d["foto_dot"] = "yellow"

        envios.append(d)

    # Mostrar primero las OTs abiertas (rojo) y enviar las cerradas al final.
    # Dentro de cada grupo, más recientes primero.
    envios.sort(key=lambda x: (1 if x.get('is_closed') else 0, -int(x.get('id') or 0)))

    # --- Usuarios para modal (solo admin) ---
    open_users_modal = (request.query_params.get("users") == "1")
    users_list = []
    perms_by_user = {}
    if _user_role(user) == "admin":
        schema = _users_schema(cur)
        cur.execute(_select_users_sql(schema))
        users_list = [dict(r) for r in cur.fetchall()]
        for u in users_list:
            perms_by_user[int(u["id"])] = _get_user_permissions_map(cur, int(u["id"]))

    conn.close()
    context = {
        "request": request,
        "user": user,
        "kpis": kpis,
        "kpis_partes": kpis_partes,
        "envios": envios,
        "q": q,
        "open_users_modal": open_users_modal,
        "users_list": users_list,
        "actions": ACTIONS,
        "perms_by_user": perms_by_user,
        "db_type": db_type,
    }

    return templates.TemplateResponse(
        "dashboard.html",
        context,
    )


# -----------------------------
# EXPORTACIÓN
# -----------------------------
def _bool_param(v: str | None) -> bool:
    if v is None:
        return False
    return str(v).strip().lower() in {"1", "true", "t", "yes", "y", "on"}


def _build_instrumentos_where(estado: str | None, solo_pendientes: bool, solo_grabados: bool | None):
    where = []
    params: list = []

    if solo_pendientes:
        where.append("i.estado IN ('Pendiente','En proceso')")
    elif estado and estado != "TODOS":
        where.append("i.estado = ?")
        params.append(estado)

    if solo_grabados is True:
        where.append("COALESCE(i.grabado,0)=1")
    elif solo_grabados is False:
        where.append("COALESCE(i.grabado,0)=0")

    sql = " AND ".join(where) if where else "1=1"
    return sql, params


@app.get("/export", response_class=HTMLResponse)
def export_home(request: Request, user=Depends(require_roles("admin", "recepcion"))):
    # Página con opciones de exportación
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT id, ot_num, cliente, fecha FROM envios ORDER BY id DESC LIMIT 400")
    envios = [dict(r) for r in cur.fetchall()]
    conn.close()
    return templates.TemplateResponse(
        "export.html",
        {
            "request": request,
            "user": user,
            "envios": envios,
            "estados": ["TODOS", "Pendiente", "En proceso", "Reparado", "Baja"],
        },
    )


@app.get("/export/download")
def export_download(
    request: Request,
    user=Depends(require_roles("admin", "recepcion")),
    scope: str = "partes",  # partes | instrumentos | parte
    envio_id: int | None = None,
    estado: str | None = None,
    solo_pendientes: str | None = None,
    grabado: str | None = None,  # todos|si|no
    fmt: str = "xlsx",  # xlsx|csv
):
    fmt = (fmt or "xlsx").lower()
    if fmt not in {"xlsx", "csv"}:
        raise HTTPException(status_code=400, detail="Formato no soportado")

    scope = (scope or "partes").lower()
    if scope not in {"partes", "instrumentos", "parte"}:
        raise HTTPException(status_code=400, detail="Scope no soportado")

    solo_pend = _bool_param(solo_pendientes)
    grabado = (grabado or "todos").lower()
    solo_grabados: bool | None = None
    if grabado == "si":
        solo_grabados = True
    elif grabado == "no":
        solo_grabados = False

    conn = get_conn()
    cur = conn.cursor()

    # ---- DATOS ----
    partes_rows: list[dict] = []
    inst_rows: list[dict] = []

    if scope in {"partes", "parte"}:
        if scope == "parte" and not envio_id:
            raise HTTPException(status_code=400, detail="Falta envio_id")

        where_env = "1=1"
        params_env: list = []
        if envio_id:
            where_env = "e.id=?"
            params_env = [int(envio_id)]

        cur.execute(
            f"""
            SELECT
                e.id,
                e.ot_num,
                e.cliente,
                e.fecha,
                e.creado_en,
                COUNT(i.id) AS n_instrumentos,
                SUM(CASE WHEN i.estado IN ('Pendiente','En proceso') THEN 1 ELSE 0 END) AS n_pendientes,
                SUM(CASE WHEN i.estado='Reparado' THEN 1 ELSE 0 END) AS n_reparados,
                SUM(CASE WHEN i.estado='Baja' THEN 1 ELSE 0 END) AS n_baja,
                SUM(CASE WHEN COALESCE(i.grabado,0)=1 THEN 1 ELSE 0 END) AS n_grabados,
                SUM(CASE WHEN i.foto_entrada_1 IS NOT NULL AND i.foto_entrada_2 IS NOT NULL AND i.foto_entrada_3 IS NOT NULL AND i.foto_entrada_4 IS NOT NULL AND i.foto_entrada_5 IS NOT NULL AND i.foto_entrada_6 IS NOT NULL THEN 1 ELSE 0 END) AS n_fotos_completas
            FROM envios e
            LEFT JOIN instrumentos i ON i.envio_id = e.id
            WHERE {where_env}
            GROUP BY e.id
            ORDER BY e.id DESC
            """,
            tuple(params_env),
        )
        for r in cur.fetchall():
            d = dict(r)
            n_inst = int(d.get("n_instrumentos") or 0)
            n_pend = int(d.get("n_pendientes") or 0)
            d["cerrado"] = 1 if (n_inst > 0 and n_pend == 0) else 0
            partes_rows.append(d)

    if scope in {"instrumentos", "parte"}:
        where_i, params_i = _build_instrumentos_where(estado, solo_pend, solo_grabados)
        where_envio = "1=1"
        params_envio: list = []
        if scope == "parte":
            where_envio = "i.envio_id=?"
            params_envio = [int(envio_id)]

        cur.execute(
            f"""
            SELECT
                i.id,
                i.envio_id,
                e.ot_num,
                e.cliente,
                e.fecha,
                i.codigo_datamatrix,
                i.fabricante,
                i.codigo_producto,
                i.denominacion,
                i.num_serie,
                i.estado,
                COALESCE(i.grabado,0) AS grabado,
                i.grabado_en,
                i.grabado_por,
                i.foto_entrada_1,
                i.foto_entrada_2,
                i.observaciones,
                i.creado_en,
                i.actualizado_en
            FROM instrumentos i
            JOIN envios e ON e.id = i.envio_id
            WHERE {where_envio} AND {where_i}
            ORDER BY e.id DESC, i.id ASC
            """,
            tuple(params_envio + params_i),
        )
        inst_rows = [dict(r) for r in cur.fetchall()]

    conn.close()

    # ---- RESPUESTA (CSV/XLSX) ----
    now_tag = datetime.now().strftime("%Y%m%d_%H%M")
    base_name = f"xurgical_export_{scope}_{now_tag}"

    if fmt == "csv":
        # Para CSV exportamos una sola tabla (prioridad instrumentos si aplica)
        rows = inst_rows if scope in {"instrumentos", "parte"} else partes_rows
        if not rows:
            rows = []
        headers = list(rows[0].keys()) if rows else []
        sio = io.StringIO()
        w = csv.writer(sio, delimiter=";")
        w.writerow(headers)
        for r in rows:
            w.writerow([r.get(h, "") for h in headers])

        data = sio.getvalue().encode("utf-8-sig")
        return StreamingResponse(
            io.BytesIO(data),
            media_type="text/csv; charset=utf-8",
            headers={"Content-Disposition": f"attachment; filename={base_name}.csv"},
        )

    # XLSX
    if Workbook is None:
        raise HTTPException(status_code=500, detail="openpyxl no está disponible")

    wb = Workbook()
    # elimina la hoja por defecto
    ws0 = wb.active
    wb.remove(ws0)

    def add_sheet(title: str, rows: list[dict]):
        ws = wb.create_sheet(title=title)
        headers = list(rows[0].keys()) if rows else []
        ws.append(headers)
        for r in rows:
            ws.append([r.get(h, "") for h in headers])
        # ancho columnas
        for col_idx, h in enumerate(headers, start=1):
            maxlen = len(str(h))
            for rr in rows[:200]:
                maxlen = max(maxlen, len(str(rr.get(h, "")) or ""))
            ws.column_dimensions[chr(64 + col_idx)].width = min(max(10, maxlen + 2), 55)

    if partes_rows:
        add_sheet("Partes", partes_rows)
    if inst_rows:
        add_sheet("Instrumentos", inst_rows)

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return StreamingResponse(
        bio,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={base_name}.xlsx"},
    )
# -----------------------------
# ENVÍOS
# -----------------------------

# -----------------------------
# CLIENTES (admin/recepcion)
# -----------------------------
@app.get("/clientes", response_class=HTMLResponse)
def clientes_list(request: Request, user=Depends(require_roles("admin", "recepcion"))):
    conn = get_conn()
    cur = conn.cursor()
    clientes = _list_clientes(cur)
    conn.close()
    return templates.TemplateResponse(
        "clientes_list.html",
        {"request": request, "user": user, "clientes": clientes},
    )


@app.get("/clientes/nuevo", response_class=HTMLResponse)
def clientes_nuevo_form(request: Request, user=Depends(require_roles("admin", "recepcion"))):
    return templates.TemplateResponse(
        "clientes_form.html",
        {"request": request, "user": user, "mode": "new", "cliente": None},
    )


@app.post("/clientes/nuevo")
def clientes_nuevo_crear(
    nombre: str = Form(""),
    prefijo: str = Form(""),
    prefijo_nombre: str = Form(""),
    ultimo_numero: int = Form(0),
    user=Depends(require_roles("admin", "recepcion")),
):
    nombre = (nombre or "").strip()
    if not nombre:
        return RedirectResponse(url="/clientes?err=nombre", status_code=303)

    conn = get_conn()
    cur = conn.cursor()

    # Prevenir duplicados (validación manual además del índice UNIQUE)
    cur.execute("SELECT id FROM clientes WHERE nombre = ?", (nombre,))
    if cur.fetchone():
        conn.close()
        return RedirectResponse(url="/clientes?err=exists", status_code=303)

    cur.execute(
        "INSERT INTO clientes (nombre, prefijo, prefijo_nombre, ultimo_numero) VALUES (?, ?, ?, ?)",
        (nombre, (prefijo or "").strip(), (prefijo_nombre or "").strip(), int(ultimo_numero or 0)),
    )
    conn.commit()
    conn.close()
    return RedirectResponse(url="/clientes", status_code=303)


@app.get("/clientes/{cliente_id}/editar", response_class=HTMLResponse)
def clientes_editar_form(request: Request, cliente_id: int, user=Depends(require_roles("admin", "recepcion"))):
    conn = get_conn()
    cur = conn.cursor()
    cli = _get_cliente(cur, cliente_id)
    conn.close()
    if not cli:
        return HTMLResponse("Cliente no encontrado", status_code=404)
    return templates.TemplateResponse(
        "clientes_form.html",
        {"request": request, "user": user, "mode": "edit", "cliente": cli},
    )


@app.post("/clientes/{cliente_id}/editar")
def clientes_editar_guardar(
    cliente_id: int,
    nombre: str = Form(""),
    prefijo: str = Form(""),
    prefijo_nombre: str = Form(""),
    ultimo_numero: int = Form(0),
    user=Depends(require_roles("admin", "recepcion")),
):
    nombre = (nombre or "").strip()
    if not nombre:
        return RedirectResponse(url=f"/clientes/{cliente_id}/editar?err=nombre", status_code=303)
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "UPDATE clientes SET nombre=?, prefijo=?, prefijo_nombre=?, ultimo_numero=? WHERE id=?",
        (nombre, (prefijo or "").strip(), (prefijo_nombre or "").strip(), int(ultimo_numero or 0), int(cliente_id)),
    )
    conn.commit()
    conn.close()
    return RedirectResponse(url="/clientes", status_code=303)

@app.get("/envios/nuevo", response_class=HTMLResponse)
def nuevo_envio_form(request: Request, user=Depends(require_roles("admin", "recepcion"))):
    conn = get_conn()
    cur = conn.cursor()
    clientes = _list_clientes(cur)
    conn.close()
    return templates.TemplateResponse(
        "envio_nuevo.html",
        {"request": request, "user": user, "clientes": clientes},
    )


@app.post("/envios/nuevo")
def nuevo_envio_crear(
    referencia: str = Form(""),
    cliente_id: str = Form(""),
    cliente: str = Form(""),
    tipo_trabajo: str = Form("REPARACION"),
    fecha: str = Form(""),
    user=Depends(require_roles("admin", "recepcion")),
):
    try:
        # Validación de fecha obligatoria
        fecha = (fecha or "").strip()
        if not fecha:
            return RedirectResponse(url="/envios/nuevo?err=fecha", status_code=303)

        # Permiso granular (además del rol)
        conn_perm = get_conn()
        cur_perm = conn_perm.cursor()
        if not can_action(user, "envio_crear", cur_perm):
            conn_perm.close()
            return RedirectResponse(url="/?err=perm", status_code=303)
        conn_perm.close()

        conn = get_conn()
        cur = conn.cursor()

        ot_num = _next_ot_num(cur)

        tipo_trabajo = (tipo_trabajo or "REPARACION").strip().upper()
        if tipo_trabajo not in ("REPARACION", "TRAZABILIDAD", "OPTICA_RIGIDA"):
            tipo_trabajo = "REPARACION"

        cli_id_val = None
        cli_nombre = (cliente or "").strip()
        if cliente_id:
            try:
                cli_id_val = int(cliente_id)
            except Exception:
                cli_id_val = None

        # Si viene un cliente_id, lo resolvemos a nombre (y garantizamos que existe)
        if cli_id_val:
            cli = _get_cliente(cur, cli_id_val)
            if not cli:
                conn.close()
                return RedirectResponse(url="/envios/nuevo?err=cliente", status_code=303)
            cli_nombre = cli["nombre"]
        else:
            # Si no hay ID, debe haber un nombre manual
            if not cli_nombre:
                conn.close()
                return RedirectResponse(url="/envios/nuevo?err=cliente", status_code=303)
            
            # Si es trazabilidad, requerimos cliente registrado (ID).
            if tipo_trabajo == "TRAZABILIDAD":
                conn.close()
                return RedirectResponse(url="/envios/nuevo?err=cliente", status_code=303)

        sql = """
            INSERT INTO envios (ot_num, nombre_archivo, cliente, cliente_id, tipo_trabajo, fecha)
            VALUES (?, ?, ?, ?, ?, ?)
        """
        vals = [
            ot_num,
            (referencia or "").strip(),
            cli_nombre,
            cli_id_val,
            tipo_trabajo,
            (fecha or "").strip(),
        ]

        is_pg = bool(os.environ.get("DATABASE_URL"))
        if is_pg:
            sql += " RETURNING id"
            cur.execute(sql, tuple(vals))
            row = cur.fetchone()
            if row:
                envio_id = int(row["id"])
            else:
                # Fallback extremo: buscar por OT exacta que acabamos de meter
                cur.execute("SELECT id FROM envios WHERE ot_num=?", (ot_num,))
                row_f = cur.fetchone()
                if row_f:
                    envio_id = int(row_f["id"])
                else:
                     raise Exception("No se pudo obtener el ID del nuevo envío en Postgres")
        else:
            cur.execute(sql, tuple(vals))
            envio_id = cur.lastrowid

        conn.commit()
        conn.close()
        # Al crear una OT, vamos al detalle para añadir instrumentos.
        return RedirectResponse(url=f"/envios/{envio_id}", status_code=303)

    except Exception as e:
        import traceback
        return HTMLResponse(f"<h1>Error al crear envio</h1><pre>{traceback.format_exc()}</pre>", status_code=500)


@app.get("/ot/{ot_num}")
def ver_ot_directa(ot_num: str, user=Depends(get_current_user)):
    """Busca una OT por su numero y redirige al detalle."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT id FROM envios WHERE ot_num=?", (ot_num,))
    e = cur.fetchone()
    conn.close()
    if not e:
        return HTMLResponse("OT no encontrada", status_code=404)
    
    # Si es técnico, le llevamos a la vista móvil directamente? O lo hacemos por rol en /envios/{id}?
    # Mejor centralizar en /envios/{id} o crear una ruta específica.
    return RedirectResponse(url=f"/envios/{e['id']}", status_code=303)


@app.get("/envios/{envio_id}", response_class=HTMLResponse)
def ver_envio(request: Request, envio_id: int, user=Depends(get_current_user)):
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("SELECT * FROM envios WHERE id=?", (envio_id,))
    envio = cur.fetchone()
    if not envio:
        conn.close()
        return HTMLResponse("Envío no encontrado", status_code=404)

    cur.execute("""
        SELECT i.*, 
               (SELECT 1 FROM instrumento_informes ii WHERE ii.instrumento_id = i.id LIMIT 1) as has_archived,
               (SELECT 1 FROM instrumento_qc_optica qc WHERE qc.instrumento_id = i.id LIMIT 1) as has_qc_data
        FROM instrumentos i 
        WHERE i.envio_id=? 
        ORDER BY i.id DESC
    """, (envio_id,))
    instrumentos = [dict(r) for r in cur.fetchall()]

    if _user_role(user) == "tecnico":
        # Aseguramos que los técnicos vean todo para poder acceder al PDF
        pass

    for r in instrumentos:
        # Limpieza de trazabilidad
        if r.get("nombre_trazabilidad"):
            r["nombre_trazabilidad"] = _clean_trz(r["nombre_trazabilidad"])
        
        # Bandera para mostrar icono PDF (informe)
        # Se muestra si hay un archivo archivado O si al menos hay datos de QC guardados
        r["has_informe"] = bool(r.get("has_archived") or r.get("has_qc_data"))

    # DEF_TRZ_NOMBRE_OK: prepara nombre_trazabilidad SOLO para pantalla de grabación (copiar/pegar)
    envio_dict = dict(envio)
    prefijo = ""
    for k in ("prefijo_nombre", "prefijo_trazabilidad", "prefijo_traz", "prefijo_etiqueta", "prefijo"):
        v = envio_dict.get(k)
        if v:
            prefijo = str(v).strip()
            break

    for r in instrumentos:
        nt = (r.get("nombre_trazabilidad") or "").strip()
        # limpia comillas accidentales tipo "'00015'"
        if nt:
            nt = nt.strip().strip("'").strip('"')
        # si es solo 5 dígitos y hay prefijo, lo completamos
        if nt and prefijo and re.fullmatch(r"\d{5}", nt):
            nt = f"{prefijo}{nt}"
        # si no hay nt pero hay DM y prefijo, lo calculamos
        if (not nt) and prefijo:
            dm = (r.get("codigo_datamatrix") or "").strip()
            if dm:
                try:
                    nt = _build_nombre_trazabilidad(prefijo, dm)
                except Exception:
                    nt = ""
        r["nombre_trazabilidad"] = _clean_trz(nt)

    conn.close()

    # Si es técnico, le mostramos la vista simplificada/móvil
    if _user_role(user) == "tecnico":
        return templates.TemplateResponse(
            "tecnico_parte.html",
            {"request": request, "user": user, "envio": dict(envio), "instrumentos": instrumentos},
        )

    return templates.TemplateResponse(
        "envio_detalle.html",
        {"request": request, "user": user, "envio": dict(envio), "instrumentos": instrumentos},
    )


# -----------------------------
# API CHECKLIST (para modo técnico móvil)
# -----------------------------
@app.get("/api/instrumentos/{instrumento_id}/checklist")
def api_get_checklist(instrumento_id: int, user=Depends(get_current_user)):
    conn = get_conn()
    cur = conn.cursor()
    
    # Obtener tipo de OT
    cur.execute("""
        SELECT e.tipo_trabajo 
        FROM instrumentos i 
        JOIN envios e ON e.id = i.envio_id 
        WHERE i.id=?
    """, (instrumento_id,))
    row = cur.fetchone()
    tipo_ot = (row["tipo_trabajo"] if row else "REPARACION") or "REPARACION"
    
    cur.execute("""
        SELECT ci.id AS item_id, ci.nombre, 
               COALESCE(ic.hecho,0) AS hecho
        FROM checklist_items ci
        LEFT JOIN instrumento_checklist ic ON ic.item_id = ci.id AND ic.instrumento_id = ?
        WHERE ci.activo = 1 AND ci.tipo_trabajo = ?
        ORDER BY ci.orden
    """, (instrumento_id, tipo_ot))
    
    checklist = [dict(r) for r in cur.fetchall()]
    conn.close()
    return {"checklist": checklist}

@app.post("/api/instrumentos/{instrumento_id}/checklist")
async def api_save_checklist(instrumento_id: int, request: Request, user=Depends(require_roles("admin", "tecnico"))):
    data = await request.json()
    items_hechos = data.get("items", []) # lista de IDs de items marcados
    
    conn = get_conn()
    cur = conn.cursor()
    
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    username = user.get("username")
    
    # 1. Desmarcar todos para este instrumento (o marcar como 0)
    cur.execute("DELETE FROM instrumento_checklist WHERE instrumento_id=?", (instrumento_id,))
    
    # 2. Insertar los marcados
    for item_id in items_hechos:
        cur.execute("""
            INSERT INTO instrumento_checklist (instrumento_id, item_id, hecho, hecho_por, hecho_en)
            VALUES (?, ?, 1, ?, ?)
        """, (instrumento_id, item_id, username, now))
    
    conn.commit()
    conn.close()
    return {"ok": True}


# -----------------------------
# ETIQUETA (pegatina) OT
# -----------------------------
def _build_etiqueta_pdf(ot_num: str, cliente: str, fecha: str, n_instrumentos: int) -> bytes:
    """Genera una etiqueta PDF con texto + código de barras Code128.

    Contenido del barcode: OT|CLIENTE|FECHA|N
    """
    # Etiqueta térmica 70x40 mm (requerimiento).
    # Al imprimir: usar "tamaño real" / 100% (sin ajustar).
    w, h = 70 * mm, 40 * mm
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(w, h))

    # Texto
    # Márgenes
    x0 = 3 * mm
    y_top = h - 3 * mm

    # Texto (compacto para 70x40)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(x0, y_top - 6 * mm, f"OT: {ot_num}")

    c.setFont("Helvetica", 10)
    # Recorta cliente para que no rompa el ancho
    cli = (cliente or "").strip()
    if len(cli) > 35:
        cli = cli[:32] + "…"
    c.drawString(x0, y_top - 11 * mm, f"Cliente: {cli}")
    c.drawString(x0, y_top - 16 * mm, f"Fecha: {fecha}")
    c.drawString(x0, y_top - 21 * mm, f"Nº Instrumentos: {n_instrumentos}")

    # Barcode
    # Payload solo con el OT para que el código de barras sea más sencillo y legible
    payload = str(ot_num)

    # QR Code ajustado para que entre en 70x40.
    # El QR es mucho más fácil de leer con móviles (iPhone) que el Barcode 1D.
    from reportlab.graphics.shapes import Drawing
    from reportlab.graphics import renderPDF
    
    qr_code = qr.QrCodeWidget(payload)
    bounds = qr_code.getBounds()
    qr_w = bounds[2] - bounds[0]
    qr_h = bounds[3] - bounds[1]
    
    # Queremos que mida unos 25mm de lado
    size = 25 * mm
    d = Drawing(size, size, transform=[size/qr_w, 0, 0, size/qr_h, 0, 0])
    d.add(qr_code)
    
    # Posicionamos el QR a la derecha (ajustando coordenadas)
    renderPDF.draw(d, c, w - size - 2*mm, 5 * mm)
    
    c.showPage()
    c.save()
    return buf.getvalue()


@app.get("/envios/{envio_id}/etiqueta.pdf")
def etiqueta_envio(envio_id: int, user=Depends(get_current_user)):
    """Devuelve una pegatina PDF para la OT."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT id, ot_num, cliente, fecha FROM envios WHERE id=?", (envio_id,))
    e = cur.fetchone()
    if not e:
        conn.close()
        return HTMLResponse("OT no encontrada", status_code=404)

    cur.execute("SELECT COUNT(*) AS n FROM instrumentos WHERE envio_id=?", (envio_id,))
    n_inst = int(cur.fetchone()["n"] or 0)
    conn.close()

    # --- Fecha para la etiqueta (dd/mm/aaaa) ---
    fecha_raw = (e["fecha"] or "").strip()
    if fecha_raw:
        try:
            dt = datetime.fromisoformat(fecha_raw)
            fecha = dt.strftime("%d/%m/%Y")
        except ValueError:
            fecha = fecha_raw
    else:
        fecha = datetime.now().strftime("%d/%m/%Y")

    pdf = _build_etiqueta_pdf(
        str(e["ot_num"]),
        str(e["cliente"] or ""),
        fecha,
        n_inst
    )
    filename = f"OT_{e['ot_num']}_etiqueta.pdf"
    return StreamingResponse(
        io.BytesIO(pdf),
        media_type="application/pdf",
        headers={"Content-Disposition": f"inline; filename={filename}"},
    )




# -----------------------------
# GRABACION (mosaico por parte)
# -----------------------------
@app.get("/envios/{envio_id}/grabacion", response_class=HTMLResponse)
def grabacion_envio(request: Request, envio_id: int, user=Depends(require_roles("admin", "grabado"))):
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("SELECT * FROM envios WHERE id=?", (envio_id,))
    envio = cur.fetchone()
    if not envio:
        conn.close()
        return HTMLResponse("Envío no encontrado", status_code=404)

    cur.execute(
        """
        SELECT
          i.*,
          COALESCE(i.grabado,0) AS grabado
        FROM instrumentos i
        WHERE i.envio_id=?
        ORDER BY i.id DESC
        """,
        (envio_id,),
    )
    instrumentos = [dict(r) for r in cur.fetchall()]

    # Conteo
    total = len(instrumentos)
    grabados = sum(1 for r in instrumentos if int(r.get("grabado") or 0) == 1)

    conn.close()
    return templates.TemplateResponse(
        "envio_grabacion.html",
        {
            "request": request,
            "user": user,
            "envio": dict(envio),
            "instrumentos": instrumentos,
            "total": total,
            "grabados": grabados,
        },
    )


@app.post("/instrumentos/{instrumento_id}/grabar")
def grabar_instrumento(instrumento_id: int, user=Depends(require_roles("admin", "grabado"))):
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("SELECT envio_id, COALESCE(grabado,0) AS grabado FROM instrumentos WHERE id=?", (instrumento_id,))
    row = cur.fetchone()
    if not row:
        conn.close()
        return HTMLResponse("Instrumento no encontrado", status_code=404)

    envio_id = int(row["envio_id"])

    cur.execute(
        """
        UPDATE instrumentos
        SET grabado=1,
            grabado_por=?,
            grabado_en=CURRENT_TIMESTAMP
        WHERE id=?
        """,
        (int(user["id"]), instrumento_id),
    )

    conn.commit()
    conn.close()

    return RedirectResponse(url=f"/envios/{envio_id}/grabacion", status_code=303)


@app.post("/instrumentos/{instrumento_id}/desgrabar")
def desgrabar_instrumento(instrumento_id: int, user=Depends(require_roles("admin"))):
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("SELECT envio_id FROM instrumentos WHERE id=?", (instrumento_id,))
    row = cur.fetchone()
    if not row:
        conn.close()
        return HTMLResponse("Instrumento no encontrado", status_code=404)

    envio_id = int(row["envio_id"])

    cur.execute(
        """
        UPDATE instrumentos
        SET grabado=0,
            grabado_por=NULL,
            grabado_en=NULL
        WHERE id=?
        """,
        (instrumento_id,),
    )

    conn.commit()
    conn.close()

    return RedirectResponse(url=f"/envios/{envio_id}/grabacion", status_code=303)

# -----------------------------
# NUEVO INSTRUMENTO (admin/recepcion)
# -----------------------------
@app.get("/envios/{envio_id}/instrumentos/nuevo", response_class=HTMLResponse)
def instrumento_nuevo_form(request: Request, envio_id: int, user=Depends(require_roles("admin", "recepcion"))):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT * FROM envios WHERE id=?", (envio_id,))
    envio = cur.fetchone()
    conn.close()
    if not envio:
        return HTMLResponse("Envío no encontrado", status_code=404)

    return templates.TemplateResponse(
        "instrumento_nuevo.html",
        {
            "request": request,
            "user": user,
            "mode": "new",
            "envio": dict(envio),
            "inst": None,
        },
    )


@app.post("/envios/{envio_id}/instrumentos/nuevo")
def instrumento_nuevo_crear(
    envio_id: int,
    codigo_producto: str = Form(""),
    fabricante: str = Form(""),
    num_serie: str = Form(""),
    denominacion: str = Form(""),
    observaciones: str = Form(""),
    codigo_datamatrix: str = Form(""),
    user=Depends(require_roles("admin", "recepcion")),
):
    # Permiso granular (además del rol)
    conn_perm = get_conn()
    cur_perm = conn_perm.cursor()
    if not can_action(user, "instrumento_crear", cur_perm):
        conn_perm.close()
        return RedirectResponse(url="/?err=perm", status_code=303)
    conn_perm.close()

    conn = get_conn()
    cur = conn.cursor()

    cur.execute("SELECT id, cliente_id, tipo_trabajo FROM envios WHERE id=?", (envio_id,))
    e = cur.fetchone()
    if not e:
        conn.close()
        return HTMLResponse("Envío no encontrado", status_code=404)

    dm_auto = ""
    nombre_trz_auto = ""
    if (e["tipo_trabajo"] or "").upper() == "TRAZABILIDAD":
        if not e["cliente_id"]:
            conn.close()
            return HTMLResponse("OT de trazabilidad sin cliente registrado", status_code=400)
        prefijo_dm, prefijo_nombre, nums = _reserve_numeros_cliente(cur, int(e["cliente_id"]), 1)
        dm_auto = f"{prefijo_dm}{str(nums[0]).zfill(5)}"
        nombre_trz_auto = _build_nombre_trazabilidad(prefijo_nombre, dm_auto)

    vals = [
        envio_id,
        (codigo_producto or "").strip(),
        (fabricante or "").strip(),
        (num_serie or "").strip(),
        (denominacion or "").strip(),
        (observaciones or "").strip(),
        (dm_auto or (codigo_datamatrix or "").strip()),
        (nombre_trz_auto or ""),
    ]

    sql = """
        INSERT INTO instrumentos
        (envio_id, codigo_producto, fabricante, num_serie, denominacion, observaciones, codigo_datamatrix, nombre_trazabilidad, estado, creado_en)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'Pendiente', CURRENT_TIMESTAMP)
    """

    is_pg = bool(os.environ.get("DATABASE_URL"))
    inst_id = None
    
    if is_pg:
        sql += " RETURNING id"
        cur.execute(sql, tuple(vals))
        row = cur.fetchone()
        if row:
            inst_id = int(row["id"])
    else:
        cur.execute(sql, tuple(vals))
        inst_id = cur.lastrowid
        
    # Fallback por si acaso (ej: driver antiguo)
    if inst_id is None:
        cur.execute("SELECT MAX(id) as mid FROM instrumentos WHERE envio_id=?", (envio_id,))
        row_id = cur.fetchone()
        if row_id and row_id["mid"]:
            inst_id = row_id["mid"]

    conn.commit()
    conn.close()

    if (e["tipo_trabajo"] or "").upper() == "TRAZABILIDAD":
        return RedirectResponse(url=f"/instrumentos/{inst_id}", status_code=303)
    return RedirectResponse(url=f"/instrumentos/{inst_id}/editar", status_code=303)


# -----------------------------
# EDITAR INSTRUMENTO (admin/recepcion)
# -----------------------------
@app.get("/instrumentos/{instrumento_id}/editar", response_class=HTMLResponse)
def instrumento_editar_form(request: Request, instrumento_id: int, user=Depends(require_roles("admin", "recepcion"))):
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("SELECT * FROM instrumentos WHERE id=?", (instrumento_id,))
    inst = cur.fetchone()
    if not inst:
        conn.close()
        return HTMLResponse("Instrumento no encontrado", status_code=404)

    cur.execute("SELECT * FROM envios WHERE id=?", (inst["envio_id"],))
    envio = cur.fetchone()
    conn.close()

    return templates.TemplateResponse(
        "instrumento_nuevo.html",
        {
            "request": request,
            "user": user,
            "mode": "edit",
            "envio": dict(envio) if envio else None,
            "inst": dict(inst),
        },
    )


@app.post("/instrumentos/{instrumento_id}/editar")
def instrumento_editar_guardar(
    instrumento_id: int,
    codigo_producto: str = Form(""),
    fabricante: str = Form(""),
    num_serie: str = Form(""),
    denominacion: str = Form(""),
    observaciones: str = Form(""),
    codigo_datamatrix: str = Form(""),
    user=Depends(require_roles("admin", "recepcion")),
):
    # Permiso granular (además del rol)
    conn_perm = get_conn()
    cur_perm = conn_perm.cursor()
    if not can_action(user, "envio_borrar", cur_perm):
        conn_perm.close()
        return RedirectResponse(url="/?err=perm", status_code=303)
    conn_perm.close()

    dm = (codigo_datamatrix or "").strip()

    conn = get_conn()
    cur = conn.cursor()

    # Obtener OT/cliente para saber si es trazabilidad y su prefijo_nombre
    cur.execute(
        """
        SELECT e.id AS envio_id, COALESCE(e.tipo_trabajo,'REPARACION') AS tipo_trabajo, e.cliente_id,
               c.prefijo_nombre
        FROM instrumentos i
        JOIN envios e ON e.id = i.envio_id
        LEFT JOIN clientes c ON c.id = e.cliente_id
        WHERE i.id=?
        """,
        (instrumento_id,),
    )
    meta = cur.fetchone()
    if not meta:
        conn.close()
        return HTMLResponse("Instrumento no encontrado", status_code=404)

    meta = dict(meta)

    nombre_trz = ""
    if (meta["tipo_trabajo"] or "").upper() == "TRAZABILIDAD":
        nombre_trz = _build_nombre_trazabilidad((meta.get("prefijo_nombre") or ""), dm)

    cur.execute(
        """
        UPDATE instrumentos
        SET codigo_producto=?,
            fabricante=?,
            num_serie=?,
            denominacion=?,
            observaciones=?,
            codigo_datamatrix=?,
            nombre_trazabilidad=?,
            actualizado_en=CURRENT_TIMESTAMP
        WHERE id=?
        """,
        (
            (codigo_producto or "").strip(),
            (fabricante or "").strip(),
            (num_serie or "").strip(),
            (denominacion or "").strip(),
            (observaciones or "").strip(),
            dm,
            nombre_trz,
            instrumento_id,
        ),
    )

    conn.commit()
    conn.close()

    return RedirectResponse(url=f"/instrumentos/{instrumento_id}/editar", status_code=303)


@app.post("/instrumentos/{instrumento_id}/finalizar_fotos")
def instrumento_finalizar_fotos(instrumento_id: int, user=Depends(require_roles("admin", "recepcion"))):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT envio_id FROM instrumentos WHERE id=?", (instrumento_id,))
    row = cur.fetchone()
    conn.close()
    if not row:
        return HTMLResponse("Instrumento no encontrado", status_code=404)
    return RedirectResponse(url=f"/envios/{row['envio_id']}", status_code=303)


# -----------------------------
# ABRIR INSTRUMENTO (todos)
# -----------------------------
@app.get("/instrumentos/{instrumento_id}", response_class=HTMLResponse)
def instrumento_detalle(request: Request, instrumento_id: int, user=Depends(get_current_user)):
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("SELECT * FROM instrumentos WHERE id=?", (instrumento_id,))
    inst = cur.fetchone()
    if not inst:
        conn.close()
        return HTMLResponse("Instrumento no encontrado", status_code=404)

    # Cargamos datos de la OT y del cliente (para modo trazabilidad)
    cur.execute("SELECT * FROM envios WHERE id=?", (inst["envio_id"],))
    envio = cur.fetchone()
    cliente = None
    if envio and envio["cliente_id"]:
        cur.execute("SELECT * FROM clientes WHERE id=?", (int(envio["cliente_id"]),))
        cliente = cur.fetchone()
    # Checklist configurable: por tipo de OT (REPARACION / OPTICA_RIGIDA / TRAZABILIDAD) y activo=1
    tipo_ot = "REPARACION"
    try:
        if envio is not None:
            # sqlite3.Row (o dict) -> tipo_trabajo
            if hasattr(envio, "keys") and ("tipo_trabajo" in envio.keys()):
                tipo_ot = envio["tipo_trabajo"] or "REPARACION"
            elif isinstance(envio, dict):
                tipo_ot = envio.get("tipo_trabajo") or "REPARACION"
    except Exception:
        tipo_ot = "REPARACION"
    tipo_ot = str(tipo_ot or "REPARACION").strip().upper()

    def _load_checklist(tipo: str):
        cur.execute("""
            SELECT ci.id AS item_id, ci.nombre, ci.orden,
                   COALESCE(ic.hecho,0) AS hecho,
                   ic.hecho_por, ic.hecho_en
            FROM checklist_items ci
            LEFT JOIN instrumento_checklist ic
              ON ic.item_id = ci.id AND ic.instrumento_id = ?
            WHERE COALESCE(ci.activo,1)=1
              AND COALESCE(ci.tipo_trabajo,'REPARACION') = ?
            ORDER BY ci.orden
        """, (instrumento_id, tipo))
        return [dict(r) for r in cur.fetchall()]

    checklist = _load_checklist(tipo_ot)

    # Fallback: si no hay checklist específico para ese tipo, usa el de REPARACION
    if (not checklist) and tipo_ot != "REPARACION":
        checklist = _load_checklist("REPARACION")
        # Informe PDF (óptica rígida): opcional
    cur.execute("SELECT * FROM instrumento_informes WHERE instrumento_id=? ORDER BY id DESC LIMIT 1", (instrumento_id,))
    informe = cur.fetchone()

    return templates.TemplateResponse(
        "instrumento_detalle.html",
        {
            "request": request,
            "user": user,
            "inst": dict(inst),
            "checklist": checklist,
            "envio": dict(envio) if envio else None,
            "cliente": dict(cliente) if cliente else None,
            "informe": dict(informe) if informe else None,
        },
    )
# -----------------------------
# CONTROL DE CALIDAD - OPTICAS RIGIDAS
# -----------------------------
@app.get("/instrumentos/{instrumento_id}/qc_optica", response_class=HTMLResponse)
async def qc_optica_view(request: Request, instrumento_id: int, user=Depends(require_roles("admin", "tecnico"))):
    conn = get_conn()
    cur = conn.cursor()
    
    cur.execute("SELECT * FROM instrumentos WHERE id=?", (instrumento_id,))
    inst = cur.fetchone()
    if not inst:
        conn.close()
        return HTMLResponse("Instrumento no encontrado", status_code=404)
        
    cur.execute("SELECT * FROM envios WHERE id=?", (inst["envio_id"],))
    envio = cur.fetchone()
    
    cur.execute("SELECT * FROM clientes WHERE id=?", (envio["cliente_id"],))
    cliente = cur.fetchone()
    
    cur.execute("SELECT * FROM instrumento_qc_optica WHERE instrumento_id=?", (instrumento_id,))
    qc = cur.fetchone()
    
    conn.close()
    
    return templates.TemplateResponse(
        "instrumento_qc_optica.html",
        {
            "request": request,
            "user": user,
            "inst": dict(inst),
            "envio": dict(envio),
            "cliente": dict(cliente) if cliente else None,
            "qc": dict(qc) if qc else None,
            "now_date": datetime.now().strftime("%Y-%m-%d")
        }
    )

@app.post("/instrumentos/{instrumento_id}/qc_optica")
async def qc_optica_save(
    instrumento_id: int,
    request: Request,
    user=Depends(require_roles("admin", "tecnico")),
    parte_trabajo_cliente: str = Form(""),
    observaciones_cliente: str = Form(""),
    observaciones_previas: str = Form(""),
    diag_ventana: str = Form("CORRECTO"),
    diag_fibra: str = Form("CORRECTO"),
    diag_objetivo: str = Form("CORRECTO"),
    diag_lentes: str = Form("CORRECTO"),
    diag_camisa: str = Form("CORRECTO"),
    diag_ocular: str = Form("CORRECTO"),
    diag_pieza_ojo: str = Form("CORRECTO"),
    diag_contaminacion: str = Form("CORRECTO"),
    reparable: int = Form(1),
    campo_vision_val: str = Form(""),
    campo_vision_ok: int = Form(1),
    direccion_vision_val: str = Form(""),
    direccion_vision_ok: int = Form(1),
    resolucion_val: str = Form(""),
    resolucion_ok: int = Form(1),
    desviacion_val: str = Form(""),
    desviacion_ok: int = Form(1),
    luz_val: str = Form(""),
    luz_ok: int = Form(1),
    observaciones_finales: str = Form(""),
    fecha_salida: str = Form(""),
    firma_tecnico: str = Form(""),
    firma_responsable: str = Form(""),
    qc_foto_entrada_1: UploadFile = File(None),
    qc_foto_entrada_2: UploadFile = File(None),
    qc_foto_salida_1: UploadFile = File(None),
    qc_foto_salida_2: UploadFile = File(None),
):
    conn = get_conn()
    cur = conn.cursor()
    
    # Manejo de fotos QC (entrada y salida)
    qc_fotos_vals = {}
    for key, f in [
        ("qc_foto_entrada_1", qc_foto_entrada_1), 
        ("qc_foto_entrada_2", qc_foto_entrada_2),
        ("qc_foto_salida_1", qc_foto_salida_1),
        ("qc_foto_salida_2", qc_foto_salida_2)
    ]:
        if f and f.filename:
            safe_name = re.sub(r"[^a-zA-Z0-9.-]", "_", f.filename)
            prefix = key.replace("qc_", "") # entrada_1, etc.
            fname = f"inst_{instrumento_id}_qc_{prefix}_{int(time.time())}_{safe_name}"
            path = os.path.join(FOTOS_DIR, fname)
            content = await f.read()
            
            # Intentar optimizar con Pillow, si falla guardar original
            try:
                from PIL import Image as PILImage
                import io
                
                img = PILImage.open(io.BytesIO(content))
                
                # Reducir si es muy grande (máx 1280px)
                max_size = 1290
                if img.width > max_size or img.height > max_size:
                    img.thumbnail((max_size, max_size), PILImage.LANCZOS)
                    
                # Convertir a RGB si es necesario (por si suben PNG con transp)
                if img.mode in ("RGBA", "P"):
                    img = img.convert("RGB")
                    
                # Guardar optimizada
                with open(path, "wb") as out:
                    img.save(out, format="JPEG", quality=85, optimize=True)
            except Exception as e:
                print(f"Error optimizando imagen {fname}: {e}")
                # Fallback: guardar tal cual
                with open(path, "wb") as out:
                    out.write(content)
            
            public_url = f"/static/fotos/{fname}"
            qc_fotos_vals[key] = public_url
            
            # Borrar anterior si existía en la tabla QC
            cur.execute(f"SELECT {key} FROM instrumento_qc_optica WHERE instrumento_id=?", (instrumento_id,))
            old = cur.fetchone()
            if old and old[key]:
                _try_delete_public_photo(old[key])

    # Guardar/Actualizar QC
    # Leemos todos los campos del form para manejar los dinámicos (split diagnosis)
    form_data = await request.form()
    
    cur.execute("SELECT 1 FROM instrumento_qc_optica WHERE instrumento_id=?", (instrumento_id,))
    exists = cur.fetchone()
    
    diag_items = ['ventana', 'fibra', 'objetivo', 'lentes', 'camisa', 'ocular', 'pieza_ojo', 'contaminacion']
    
    if exists:
        update_cols = [
            "parte_trabajo_cliente=?", "observaciones_cliente=?", "observaciones_previas=?",
            "reparable=?", "campo_vision_val=?", "campo_vision_ok=?", "direccion_vision_val=?", 
            "direccion_vision_ok=?", "resolucion_val=?", "resolucion_ok=?", "desviacion_val=?", 
            "desviacion_ok=?", "luz_val=?", "luz_ok=?", "observaciones_finales=?", "fecha_salida=?",
            "firma_tecnico=?", "firma_responsable=?"
        ]
        params = [
            parte_trabajo_cliente, observaciones_cliente, observaciones_previas,
            reparable, campo_vision_val, campo_vision_ok, direccion_vision_val, 
            direccion_vision_ok, resolucion_val, resolucion_ok, desviacion_val, 
            desviacion_ok, luz_val, luz_ok, observaciones_finales, fecha_salida,
            firma_tecnico, firma_responsable
        ]
        
        # Nuevas columnas de diagnóstico split
        for item in diag_items:
            update_cols.append(f"diag_{item}_estado=?")
            params.append(form_data.get(f"diag_{item}_estado", "CORRECTO"))
            update_cols.append(f"diag_{item}_accion=?")
            params.append(form_data.get(f"diag_{item}_accion", ""))
        
        for k, v in qc_fotos_vals.items():
            update_cols.append(f"{k}=?")
            params.append(v)
            
        params.append(instrumento_id)
        sql = f"UPDATE instrumento_qc_optica SET {', '.join(update_cols)} WHERE instrumento_id=?"
        cur.execute(sql, tuple(params))
    else:
        insert_cols = [
            "instrumento_id", "parte_trabajo_cliente", "observaciones_cliente", "observaciones_previas",
            "reparable", "campo_vision_val", "campo_vision_ok", "direccion_vision_val", "direccion_vision_ok",
            "resolucion_val", "resolucion_ok", "desviacion_val", "desviacion_ok",
            "luz_val", "luz_ok", "observaciones_finales", "fecha_salida",
            "firma_tecnico", "firma_responsable"
        ]
        params = [
            instrumento_id, parte_trabajo_cliente, observaciones_cliente, observaciones_previas,
            reparable, campo_vision_val, campo_vision_ok, direccion_vision_val, direccion_vision_ok,
            resolucion_val, resolucion_ok, desviacion_val, desviacion_ok,
            luz_val, luz_ok, observaciones_finales, fecha_salida,
            firma_tecnico, firma_responsable
        ]
        
        for item in diag_items:
            insert_cols.append(f"diag_{item}_estado")
            params.append(form_data.get(f"diag_{item}_estado", "CORRECTO"))
            insert_cols.append(f"diag_{item}_accion")
            params.append(form_data.get(f"diag_{item}_accion", ""))

        for k, v in qc_fotos_vals.items():
            insert_cols.append(k)
            params.append(v)
            
        placeholders = ", ".join(["?"] * len(params))
        sql = f"INSERT INTO instrumento_qc_optica ({', '.join(insert_cols)}) VALUES ({placeholders})"
        cur.execute(sql, tuple(params))

    conn.commit()
    conn.close()
    
    # --- NUEVOS PASOS PARA ARCHIVAR EL PDF AUTOMÁTICAMENTE ---
    conn = get_conn()
    
    try:
        cur = conn.cursor()
        # Recargar datos frescos para el generador
        pdf_bytes, filename = _generate_qc_optica_pdf_bytes(instrumento_id)
        if pdf_bytes:
            informes_dir = os.path.join(UPLOAD_DIR, "informes")
            os.makedirs(informes_dir, exist_ok=True)
            # Nombre único con timestamp para no machacar registros históricos si se quiere
            stored_name = f"auto_qc_{instrumento_id}_{int(time.time())}.pdf"
            full_path = os.path.join(informes_dir, stored_name)
            with open(full_path, "wb") as f_pdf:
                f_pdf.write(pdf_bytes)
                
            # Almacenar la ruta relativa para que sea accesible vía web
            db_path = f"uploads/informes/{stored_name}"
            
            # Insertar en instrumento_informes para que aparezca en el listado
            from db import get_table_columns
            cols = get_table_columns(cur, "instrumento_informes")
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            username = (user.get("username") if isinstance(user, dict) else getattr(user, "username", "SISTEMA"))
            
            insert_cols = ["instrumento_id", "filename"]
            insert_vals = [instrumento_id, filename]
            if "path" in cols:
                insert_cols.append("path")
                insert_vals.append(db_path)
            if "filepath" in cols:
                insert_cols.append("filepath")
                insert_vals.append(db_path)
            if "uploaded_at" in cols:
                insert_cols.append("uploaded_at")
                insert_vals.append(now_str)
            if "uploaded_by" in cols:
                insert_cols.append("uploaded_by")
                insert_vals.append(username)
                
            placeholders = ", ".join(["?"] * len(insert_cols))
            cur.execute(f"INSERT INTO instrumento_informes ({', '.join(insert_cols)}) VALUES ({placeholders})", tuple(insert_vals))
            conn.commit()
    except Exception as e:
        print(f"Error generando PDF automatico: {e}")
        # Continuamos con el resto (marcar reparado etc)
        
    # Marcar instrumento como REPARADO al guardar el QC
    try:
        cur = conn.cursor()
        cur.execute("UPDATE instrumentos SET estado='Reparado' WHERE id=?", (instrumento_id,))
        conn.commit()

        # Redirigir de vuelta al listado del parte (envio_id)
        cur.execute("SELECT envio_id FROM instrumentos WHERE id=?", (instrumento_id,))
        row_e = cur.fetchone()
    except Exception as e:
        print(f"Error actualizando estado o redireccionando: {e}")
        row_e = None
    finally:
        conn.close()
    
    dest_url = f"/envios/{row_e['envio_id']}" if row_e else "/"
    return RedirectResponse(url=dest_url, status_code=303)


def _generate_qc_optica_pdf_bytes(instrumento_id: int):
    conn = get_conn()
    cur = conn.cursor()
    
    cur.execute("SELECT * FROM instrumentos WHERE id=?", (instrumento_id,))
    inst = cur.fetchone()
    if not inst:
        conn.close()
        return None, None
        
    cur.execute("SELECT * FROM envios WHERE id=?", (inst["envio_id"],))
    envio = cur.fetchone()
    
    cur.execute("SELECT * FROM clientes WHERE id=?", (envio["cliente_id"],))
    cliente = cur.fetchone()
    
    cur.execute("SELECT * FROM instrumento_qc_optica WHERE instrumento_id=?", (instrumento_id,))
    qc_row = cur.fetchone()
    conn.close()
    
    if not qc_row:
        return None, None

    # Convertir a dict para poder usar .get() y evitar errores de sqlite3.Row
    inst = dict(inst)
    envio = dict(envio)
    cliente = dict(cliente) if cliente else None
    qc = dict(qc_row)

    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, HRFlowable, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=1.5*cm, leftMargin=1.5*cm, topMargin=1.5*cm, bottomMargin=1.5*cm)
    styles = getSampleStyleSheet()
    
    # Colores Corporativos
    c_navy = colors.HexColor('#002D62')
    c_light_bg = colors.HexColor('#F4F7F9')
    c_border = colors.HexColor('#D1D5DB')
    
    # Estilos Personalizados
    style_h1 = ParagraphStyle('H1', parent=styles['Heading1'], fontSize=16, textColor=c_navy, alignment=TA_CENTER, spaceAfter=20, fontName='Helvetica-Bold')
    style_sec = ParagraphStyle('Section', parent=styles['Heading2'], fontSize=11, textColor=colors.black, fontName='Helvetica-Bold', spaceBefore=10, spaceAfter=8, borderPadding=2)
    style_normal = ParagraphStyle('NormalSmall', parent=styles['Normal'], fontSize=9, leading=11)
    style_cell_label = ParagraphStyle('CellLabel', parent=styles['Normal'], fontSize=8, fontName='Helvetica-Bold', textColor=c_navy)

    elements = []

    # --- CABECERA ---
    logo_path = os.path.join("static", "logo-xurgical.png")
    logo = Image(logo_path, width=4.5*cm, height=1.35*cm) if os.path.exists(logo_path) else Paragraph("XURGICAL", styles["Normal"])
    
    header_right = [
        [Paragraph("<b>CERTIFICADO DE CONTROL DE CALIDAD</b>", ParagraphStyle('cc', fontSize=12, alignment=TA_RIGHT, textColor=c_navy))],
        [Paragraph(f"Informe Nº: CC-{24000 + instrumento_id}", ParagraphStyle('id', fontSize=10, alignment=TA_RIGHT))],
        [Paragraph(f"Fecha: {qc['fecha_salida'] or datetime.now().strftime('%d/%m/%Y')}", ParagraphStyle('dt', fontSize=10, alignment=TA_RIGHT))]
    ]
    header_tab = Table([[logo, Table(header_right, colWidths=[8*cm])]], colWidths=[8*cm, 10*cm])
    header_tab.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('ALIGN', (1,0), (1,0), 'RIGHT')]))
    elements.append(header_tab)
    elements.append(HRFlowable(width="100%", thickness=1, color=c_navy, spaceBefore=4, spaceAfter=15))

    # --- DATOS GENERALES ---
    elements.append(Paragraph("DATOS DEL EQUIPO Y CLIENTE", style_sec))
    
    dg_data = [
        [Paragraph("CLIENTE", style_cell_label), str(cliente["nombre"] if cliente else envio["cliente"])[:50]],
        [Paragraph("DENOMINACIÓN", style_cell_label), str(inst["denominacion"])[:50]],
        [Paragraph("MODELO / REF", style_cell_label), str(inst["codigo_producto"])[:30]],
        [Paragraph("Nº DE SERIE", style_cell_label), str(inst["num_serie"])[:30]],
        [Paragraph("ORDEN TRABAJO", style_cell_label), str(envio["ot_num"])[:20]],
        ["", ""]
    ]
    # Reorganizar en 2 columnas
    dg_tab_inner = [
        [dg_data[0][0], dg_data[0][1], dg_data[1][0], dg_data[1][1]],
        [dg_data[2][0], dg_data[2][1], dg_data[3][0], dg_data[3][1]],
        [dg_data[4][0], dg_data[4][1], "", ""]
    ]
    dg_tab = Table(dg_tab_inner, colWidths=[3.5*cm, 5.5*cm, 3.5*cm, 5.5*cm])
    dg_tab.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, c_border),
        ('BACKGROUND', (0,0), (0,-1), c_light_bg),
        ('BACKGROUND', (2,0), (2,-1), c_light_bg),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('FONTSIZE', (0,0), (-1,-1), 9),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
    ]))
    elements.append(dg_tab)
    
    # Observaciones adicionales
    obs_extra_data = [
        [Paragraph("OBS. CLIENTE", style_cell_label), Paragraph(qc["observaciones_cliente"] or "-", style_normal)],
        [Paragraph("OBS. PREVIAS", style_cell_label), Paragraph(qc["observaciones_previas"] or "-", style_normal)]
    ]
    obs_extra_tab = Table(obs_extra_data, colWidths=[3.5*cm, 14.5*cm])
    obs_extra_tab.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, c_border),
        ('BACKGROUND', (0,0), (0,-1), c_light_bg),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    elements.append(obs_extra_tab)
    elements.append(Spacer(1, 15))

    # --- DIAGNÓSTICO TÉCNICO ---
    elements.append(Paragraph("INSPECCIÓN Y DIAGNÓSTICO DE COMPONENTES", style_sec))
    diag_rows = [
        [
            Paragraph("ELEMENTO", style_cell_label),
            Paragraph("EVALUACIÓN", style_cell_label),
            "",
            Paragraph("ACCIÓN (SI INCORRECTO)", style_cell_label),
            ""
        ],
        [
            "",
            Paragraph("CORR.", style_cell_label),
            Paragraph("INCOR.", style_cell_label),
            Paragraph("SUST.", style_cell_label),
            Paragraph("REPAR.", style_cell_label)
        ]
    ]
    
    elementos_qc = [
        ('ventana', 'VENTANA'), ('fibra', 'FIBRA ILUMINACIÓN'), ('objetivo', 'OBJETIVO'),
        ('lentes', 'LENTES'), ('camisa', 'CAMISA EXTERIOR'), ('ocular', 'OCULAR'),
        ('pieza_ojo', 'PIEZA DE OJO'), ('contaminacion', 'CONTAMINACIÓN')
    ]
    
    for item, label in elementos_qc:
        est = qc.get(f"diag_{item}_estado", "CORRECTO")
        acc = qc.get(f"diag_{item}_accion", "")
        diag_rows.append([
            label, 
            "X" if est=="CORRECTO" else "", 
            "X" if est=="INCORRECTO" else "", 
            "X" if acc=="SUSTITUCION" else "", 
            "X" if acc=="REPARACION" else ""
        ])
        
    diag_tab = Table(diag_rows, colWidths=[6*cm, 3*cm, 3*cm, 3*cm, 3*cm])
    diag_tab.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, c_border),
        ('SPAN', (1,0), (2,0)), # Span Evaluación
        ('SPAN', (3,0), (4,0)), # Span Acción
        ('SPAN', (0,0), (0,1)), # Span Elemento header
        ('BACKGROUND', (0,0), (-1,1), c_light_bg),
        ('BACKGROUND', (3,0), (4,-1), colors.HexColor('#F9FAFB')), # Fondo sutil para acción
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('ALIGN', (1,1), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
    ]))
    elements.append(diag_tab)
    
    # Salto de página para que las fotos y lo demás vayan a la pág 2
    elements.append(PageBreak())

    # --- FOTOS ---
    elements.append(Paragraph("REGISTRO FOTOGRÁFICO DE DIAGNÓSTICO Y REPARACIÓN", style_sec))
    
    def _get_pdf_img(p):
        if not p: return Paragraph("<br/><br/><i>Sin imagen</i>", ParagraphStyle('si', alignment=TA_CENTER, fontSize=8, textColor=colors.gray))
        
        # Intentar resolver ruta. p suele ser '/static/fotos/...'
        fname = p.split("/")[-1]
        # Ruta en Render/Local suele ser BASE_DIR / static / fotos / fname
        loc = os.path.join(str(BASE_DIR), "static", "fotos", fname)
        
        if not os.path.exists(loc):
            # Fallback 2: ruta directa quitando el / inicial
            loc = p.lstrip("/")
            if not os.path.exists(loc):
                return "N/A"
        
        try:
            return Image(loc, width=7.5*cm, height=5.5*cm, kind='proportional')
        except:
            return "Error carga"

    foto_data = [
        [Paragraph("<b>ESTADO INICIAL / ENTRADA</b>", style_normal), Paragraph("<b>ESTADO FINAL / SALIDA</b>", style_normal)],
        [_get_pdf_img(qc["qc_foto_entrada_1"] or inst["foto_entrada_1"]), _get_pdf_img(qc["qc_foto_salida_1"] or inst["foto_salida_1"])],
        [_get_pdf_img(qc["qc_foto_entrada_2"] or inst["foto_entrada_2"]), _get_pdf_img(qc["qc_foto_salida_2"] or inst["foto_salida_2"])]
    ]
    foto_tab = Table(foto_data, colWidths=[8.5*cm, 8.5*cm])
    foto_tab.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,1), (-1,-1), 0.5, c_border),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
    ]))
    elements.append(foto_tab)

    # --- VERIFICACIÓN DE PARÁMETROS ---
    elements.append(Paragraph("VERIFICACIÓN DE ESPECIFICACIONES TÉCNICAS", style_sec))
    tec_rows = [[
        Paragraph("PARÁMETRO", style_cell_label),
        Paragraph("VALOR", style_cell_label),
        Paragraph("VÁLIDO (SÍ)", style_cell_label),
        Paragraph("VÁLIDO (NO)", style_cell_label)
    ]]
    for key, label in [('campo_vision', 'CAMPO DE VISIÓN'), ('direccion_vision', 'DIRECCIÓN DE VISIÓN'), 
                       ('resolucion', 'RESOLUCIÓN'), ('desviacion', 'DESVIACIÓN'), ('luz', 'LUZ')]:
        ok = qc[key+"_ok"]
        tec_rows.append([label, qc[key+"_val"] or "-", "X" if ok==1 else "", "X" if ok==0 else ""])
        
    tec_tab = Table(tec_rows, colWidths=[6.5*cm, 4.5*cm, 3.5*cm, 3.5*cm])
    tec_tab.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, c_border),
        ('BACKGROUND', (0,0), (-1,0), c_light_bg),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('ALIGN', (1,1), (-1,-1), 'CENTER'),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
    ]))
    elements.append(tec_tab)
    elements.append(Spacer(1, 15))

    # --- CIERRE ---
    elements.append(Paragraph("CONCLUSIONES Y OBSERVACIONES FINALES", style_sec))
    obs = qc["observaciones_finales"] or "El equipo ha sido sometido a las pruebas de control de calidad indicadas, resultando apto para su uso clínico tras la intervención realizada."
    elements.append(Paragraph(obs, style_normal))
    
    elements.append(Spacer(1, 20))
    
    # Firmas
    firma_data = [
        [Paragraph(f"<b>CONTROL DE CALIDAD</b><br/><br/><br/>{qc['firma_tecnico'] or 'Técnico Especialista'}", style_normal), 
         Paragraph(f"<b>RESPONSABLE TÉCNICO</b><br/><br/><br/>{qc['firma_responsable'] or 'Director Técnico'}", style_normal)]
    ]
    firma_tab = Table(firma_data, colWidths=[8.5*cm, 8.5*cm])
    firma_tab.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER')]))
    elements.append(firma_tab)

    # --- FOOTER (Se genera en cada página) ---
    def static_footer(canvas, doc):
        canvas.saveState()
        canvas.setFont('Helvetica', 8)
        canvas.setStrokeColor(c_navy)
        canvas.line(1.5*cm, 1.2*cm, 19.5*cm, 1.2*cm)
        footer_text = "XURGICAL - Soluciones en Instrumentación Quirúrgica | www.xurgical.com"
        canvas.drawCentredString(10.5*cm, 0.8*cm, footer_text)
        canvas.restoreState()

    doc.build(elements, onFirstPage=static_footer, onLaterPages=static_footer)
    pdf_bytes = buf.getvalue()
    clean_code = str(inst['codigo_producto'] or 'SREF').replace("/", "_").replace("\\", "_")
    filename = f"QC_OPTICA_{instrumento_id}_{clean_code}.pdf"
    return pdf_bytes, filename

@app.get("/instrumentos/{instrumento_id}/qc_optica/pdf")
async def qc_optica_pdf_gen(instrumento_id: int, user=Depends(get_current_user)):
    pdf_bytes, filename = _generate_qc_optica_pdf_bytes(instrumento_id)
    if not pdf_bytes:
        raise HTTPException(status_code=404, detail="No hay datos de QC para este instrumento")
    
    return StreamingResponse(
        io.BytesIO(pdf_bytes),
        media_type="application/pdf",
        headers={"Content-Disposition": f"inline; filename={filename}"}
    )


# -----------------------------
# CHECKLIST (admin/tecnico)
# Endpoint que usa el botón OK del checklist.
# Si no existe, el navegador devuelve {"detail":"No encontrado"}.
# -----------------------------
# -----------------------------
# INFORME PDF (óptica rígida)
# -----------------------------
@app.get("/instrumentos/{instrumento_id}/informe.pdf", response_class=FileResponse)
@app.get("/instrumentos/{instrumento_id}/informe", response_class=FileResponse)
def instrumento_informe_download(instrumento_id: int, user=Depends(get_current_user)):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT path, filename FROM instrumento_informes WHERE instrumento_id=? ORDER BY id DESC LIMIT 1", (instrumento_id,))
    row = cur.fetchone()
    conn.close()
    
    # Si existe el registro y el archivo en disco, lo servimos directamente
    if row and row["path"] and os.path.exists(row["path"]):
        return FileResponse(row["path"], filename=row["filename"], media_type="application/pdf")
    
    # FALLBACK INMINENTE: Si no hay archivo guardado, intentamos generarlo al vuelo
    pdf_bytes, filename = _generate_qc_optica_pdf_bytes(instrumento_id)
    if pdf_bytes:
        return StreamingResponse(
            io.BytesIO(pdf_bytes),
            media_type="application/pdf",
            headers={"Content-Disposition": f"inline; filename={filename}"}
        )
        
    raise HTTPException(status_code=404, detail="Informe no encontrado y no se puede generar (faltan datos QC)")


@app.post("/instrumentos/{instrumento_id}/informe/upload")
async def instrumento_informe_upload(
    instrumento_id: int,
    file: UploadFile = File(...),
    user=Depends(require_roles("admin", "recepcion", "tecnico")),
):
    """Sube un informe PDF para un instrumento.

    Esta versión es compatible con BDs antiguas que usaban la columna `filepath`
    (incluso si era NOT NULL). Inserta dinámicamente en `path` y/o `filepath`
    según existan en la tabla real.
    """
    if not (file.filename or "").lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Solo se permite PDF")

    informes_dir = os.path.join(UPLOAD_DIR, "informes")
    os.makedirs(informes_dir, exist_ok=True)

    safe_name = re.sub(r"[^a-zA-Z0-9._-]", "_", file.filename)
    stored_name = f"inst_{instrumento_id}_{int(time.time())}_{safe_name}"
    path = os.path.join(informes_dir, stored_name)

    with open(path, "wb") as f:
        f.write(await file.read())

    conn = get_conn()
    cur = conn.cursor()

    # Descubre columnas reales
    from db import get_table_columns
    cols = get_table_columns(cur, "instrumento_informes")

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # user puede ser dict (tu auth) o un objeto
    username = None
    try:
        if isinstance(user, dict):
            username = user.get("username")
        else:
            username = getattr(user, "username", None)
    except Exception:
        username = None

    insert_cols = ["instrumento_id", "filename"]
    insert_vals = [instrumento_id, file.filename]

    if "path" in cols:
        insert_cols.append("path")
        insert_vals.append(path)

    if "filepath" in cols:
        insert_cols.append("filepath")
        insert_vals.append(path)

    if "uploaded_at" in cols:
        insert_cols.append("uploaded_at")
        insert_vals.append(now)

    if "uploaded_by" in cols:
        insert_cols.append("uploaded_by")
        insert_vals.append(username)

    if ("path" not in cols) and ("filepath" not in cols):
        conn.close()
        raise HTTPException(status_code=500, detail="La tabla instrumento_informes no tiene columnas path/filepath")

    placeholders = ", ".join(["?"] * len(insert_cols))
    sql = f"INSERT INTO instrumento_informes ({', '.join(insert_cols)}) VALUES ({placeholders})"
    cur.execute(sql, tuple(insert_vals))

    conn.commit()
    conn.close()

    return RedirectResponse(url=f"/instrumentos/{instrumento_id}", status_code=303)


@app.post("/instrumentos/{instrumento_id}/check/{item_id}")
async def checklist_toggle(
    instrumento_id: int,
    item_id: int,
    request: Request,
    user=Depends(require_roles("admin", "tecnico")),
):
    """Marca/desmarca un item del checklist para un instrumento.

    Admite POST desde formulario (por ejemplo, botón OK) y devuelve redirect al
    detalle del instrumento.

    - Si el form incluye `hecho` ("1"/"0"/"on"), se usa ese valor.
    - Si no incluye `hecho`, se hace toggle del estado actual.
    """

    conn = get_conn()
    cur = conn.cursor()

    # Verifica instrumento
    cur.execute("SELECT id FROM instrumentos WHERE id=?", (instrumento_id,))
    if not cur.fetchone():
        conn.close()
        raise HTTPException(status_code=404, detail="No encontrado")

    # Verifica item
    cur.execute("SELECT id FROM checklist_items WHERE id=?", (item_id,))
    if not cur.fetchone():
        conn.close()
        raise HTTPException(status_code=404, detail="No encontrado")

    form = await request.form()
    raw = form.get("hecho")

    # Estado actual
    cur.execute(
        "SELECT hecho FROM instrumento_checklist WHERE instrumento_id=? AND item_id=?",
        (instrumento_id, item_id),
    )
    row = cur.fetchone()
    actual = int(row["hecho"]) if row and isinstance(row, dict) else (int(row[0]) if row else 0)

    if raw is None:
        nuevo = 0 if actual == 1 else 1
    else:
        s = str(raw).strip().lower()
        nuevo = 1 if s in ("1", "true", "on", "yes", "si") else 0

    hecho_por = user.get("username") if isinstance(user, dict) else None

    if row:
        cur.execute(
            """
            UPDATE instrumento_checklist
               SET hecho=?, hecho_por=?, hecho_en=CURRENT_TIMESTAMP
             WHERE instrumento_id=? AND item_id=?
            """,
            (nuevo, hecho_por, instrumento_id, item_id),
        )
    else:
        cur.execute(
            """
            INSERT INTO instrumento_checklist (instrumento_id, item_id, hecho, hecho_por, hecho_en)
            VALUES (?, ?, ?, ?, CURRENT_TIMESTAMP)
            """,
            (instrumento_id, item_id, nuevo, hecho_por),
        )

    conn.commit()
    conn.close()

    # Si viene por AJAX, devolvemos JSON
    accept = (request.headers.get("accept") or "").lower()
    xreq = (request.headers.get("x-requested-with") or "").lower()
    if "application/json" in accept or xreq == "xmlhttprequest":
        return {"ok": True, "instrumento_id": instrumento_id, "item_id": item_id, "hecho": nuevo}

    return RedirectResponse(url=f"/instrumentos/{instrumento_id}", status_code=303)


# -----------------------------
# CAMBIAR ESTADO (admin/tecnico)
# -----------------------------
@app.post("/instrumentos/{instrumento_id}/estado")
async def cambiar_estado(
    instrumento_id: int,
    request: Request,
    user=Depends(require_roles("admin", "tecnico")),
):
    """Cambia el estado del instrumento (Pendiente/En proceso/Reparado/Baja) y vuelve a la página anterior."""
    form = await request.form()
    estado = (form.get("estado") or "").strip()

    if estado not in ("Pendiente", "En proceso", "Reparado", "Baja"):
        return HTMLResponse("Estado inválido", status_code=400)

    conn = get_conn()
    cur = conn.cursor()

    # Si la OT es de TRAZABILIDAD, no se permite cambiar estado de reparación.
    cur.execute("SELECT envio_id FROM instrumentos WHERE id=?", (instrumento_id,))
    row = cur.fetchone()
    envio_id = int(row["envio_id"]) if row and row["envio_id"] is not None else None

    if envio_id is not None and _envios_has_column(cur, "tipo_trabajo"):
        cur.execute("SELECT tipo_trabajo FROM envios WHERE id=?", (envio_id,))
        e = cur.fetchone()
        tipo = (e["tipo_trabajo"] if e and e["tipo_trabajo"] else "REPARACION")
        if tipo == "TRAZABILIDAD":
            conn.close()
            return HTMLResponse("En OTs de trazabilidad no se cambia el estado de reparación.", status_code=400)

    # Actualiza estado
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    username = user.get("username")

    cur.execute("""
        UPDATE instrumentos 
        SET estado=?, 
            tecnico_reparacion=?, 
            tecnico_reparacion_en=? 
        WHERE id=?
    """, (estado, username, ahora, instrumento_id))
    
    conn.commit()
    conn.close()

    # Volver a la página de origen (envio_detalle o instrumento_detalle)
    back = request.headers.get("referer") or "/"
    return RedirectResponse(url=back, status_code=303)



# -----------------------------
# SUBIR FOTO DESDE WEBCAM (admin/recepcion)
# Body JSON: { image: "data:image/jpeg;base64,...." }
# -----------------------------
@app.post("/instrumentos/{instrumento_id}/foto_webcam/{slot}")
async def foto_webcam(instrumento_id: int, slot: int, request: Request, user=Depends(require_roles("admin", "recepcion"))):
    if slot not in range(1, 7):
        return JSONResponse({"ok": False, "error": "slot inválido"}, status_code=400)

    data = await request.json()
    image = (data.get("image") or "")
    if "," in image:
        image = image.split(",", 1)[1]

    try:
        raw = base64.b64decode(image)
    except Exception:
        return JSONResponse({"ok": False, "error": "imagen inválida"}, status_code=400)

    col = f"foto_entrada_{slot}"

    conn = get_conn()
    cur = conn.cursor()
    cur.execute(f"SELECT codigo_datamatrix, nombre_trazabilidad, {col} AS old FROM instrumentos WHERE id=?", (instrumento_id,))
    row = cur.fetchone()
    if not row:
        conn.close()
        return JSONResponse({"ok": False, "error": "instrumento no encontrado"}, status_code=404)

    row = dict(row)

    # nombre para la foto: prioriza nombre_trazabilidad (TRZ), si no usa datamatrix
    tag = (row.get("nombre_trazabilidad") or row.get("codigo_datamatrix") or "").strip()
    tag = re.sub(r"[^A-Za-z0-9_-]+", "_", tag)[:40] if tag else "SIN_CODIGO"

    filename = f"inst_{instrumento_id}_{tag}_f{slot}_{uuid.uuid4().hex[:8]}.jpg"
    path_fs = os.path.join(FOTOS_DIR, filename)
    with open(path_fs, "wb") as f:
        f.write(raw)

    public_path = f"/static/fotos/{filename}"

    # borrar anterior si existe
    old_path = row.get("old")
    if old_path:
        _try_delete_public_photo(old_path)

    cur.execute(f"UPDATE instrumentos SET {col}=? WHERE id=?", (public_path, instrumento_id))
    conn.commit()
    conn.close()

    return {"ok": True, "path": public_path}



# -----------------------------
# CHECKLIST ADMIN (configurable)
# -----------------------------
@app.get("/checklist_admin", response_class=HTMLResponse)
def checklist_admin(request: Request, tipo: str = "REPARACION", user=Depends(require_roles("admin"))):
    tipo = (tipo or "REPARACION").strip().upper()
    if tipo not in ("REPARACION", "TRAZABILIDAD", "OPTICA_RIGIDA"):
        tipo = "REPARACION"

    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        SELECT id, nombre, orden,
               COALESCE(activo,1) AS activo,
               COALESCE(tipo_trabajo,'REPARACION') AS tipo_trabajo
        FROM checklist_items
        WHERE COALESCE(tipo_trabajo,'REPARACION') = ?
        ORDER BY orden, id
    """, (tipo,))
    items = [dict(r) for r in cur.fetchall()]
    conn.close()

    return templates.TemplateResponse(
        "checklist_admin.html",
        {
            "request": request,
            "user": user,
            "tipo": tipo,
            "items": items,
            "tipos": ["REPARACION", "OPTICA_RIGIDA", "TRAZABILIDAD"],
        },
    )


@app.post("/checklist_admin/add")
async def checklist_admin_add(
    request: Request,
    user=Depends(require_roles("admin")),
):
    form = await request.form()
    nombre = (form.get("nombre") or "").strip()
    tipo = (form.get("tipo_trabajo") or "REPARACION").strip().upper()
    orden = form.get("orden") or "0"

    try:
        orden_i = int(str(orden).strip() or "0")
    except Exception:
        orden_i = 0

    if not nombre:
        return RedirectResponse(url=f"/checklist_admin?tipo={tipo}", status_code=303)

    if tipo not in ("REPARACION", "TRAZABILIDAD", "OPTICA_RIGIDA"):
        tipo = "REPARACION"

    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO checklist_items (nombre, orden, activo, tipo_trabajo) VALUES (?, ?, 1, ?)",
        (nombre, orden_i, tipo),
    )
    conn.commit()
    conn.close()
    return RedirectResponse(url=f"/checklist_admin?tipo={tipo}", status_code=303)


@app.post("/checklist_admin/{item_id}/update")
async def checklist_admin_update(
    item_id: int,
    request: Request,
    user=Depends(require_roles("admin")),
):
    form = await request.form()
    nombre = (form.get("nombre") or "").strip()
    tipo = (form.get("tipo_trabajo") or "REPARACION").strip().upper()
    orden = form.get("orden") or "0"
    try:
        orden_i = int(str(orden).strip() or "0")
    except Exception:
        orden_i = 0

    if tipo not in ("REPARACION", "TRAZABILIDAD", "OPTICA_RIGIDA"):
        tipo = "REPARACION"

    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "UPDATE checklist_items SET nombre=?, orden=?, tipo_trabajo=? WHERE id=?",
        (nombre, orden_i, tipo, item_id),
    )
    conn.commit()
    conn.close()
    return RedirectResponse(url=f"/checklist_admin?tipo={tipo}", status_code=303)


@app.post("/checklist_admin/{item_id}/toggle")
async def checklist_admin_toggle(
    item_id: int,
    request: Request,
    user=Depends(require_roles("admin")),
):
    form = await request.form()
    tipo = (form.get("tipo") or "REPARACION").strip().upper()

    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT COALESCE(activo,1) AS activo FROM checklist_items WHERE id=?", (item_id,))
    row = cur.fetchone()
    if row:
        activo = int(row["activo"] if isinstance(row, dict) else row[0])
        nuevo = 0 if activo == 1 else 1
        cur.execute("UPDATE checklist_items SET activo=? WHERE id=?", (nuevo, item_id))
        conn.commit()
    conn.close()
    return RedirectResponse(url=f"/checklist_admin?tipo={tipo}", status_code=303)


@app.post("/instrumentos/{instrumento_id}/foto_borrar/{slot}")
def foto_borrar(instrumento_id: int, slot: int, user=Depends(require_roles("admin", "recepcion", "tecnico", "grabado"))):
    if slot not in range(1, 7):
        return JSONResponse({"ok": False, "error": "slot inválido"}, status_code=400)

    col = f"foto_entrada_{slot}"

    conn = get_conn()
    cur = conn.cursor()
    cur.execute(f"SELECT {col} FROM instrumentos WHERE id=?", (instrumento_id,))
    row = cur.fetchone()
    if not row:
        conn.close()
        return JSONResponse({"ok": False, "error": "Instrumento no encontrado"}, status_code=404)

    old_path = row[col]
    cur.execute(f"UPDATE instrumentos SET {col}=NULL WHERE id=?", (instrumento_id,))
    conn.commit()
    conn.close()

    _try_delete_public_photo(old_path)
    return {"ok": True}


# -----------------------------
# BORRAR INSTRUMENTO (admin/recepcion)
# -----------------------------
@app.post("/instrumentos/{instrumento_id}/borrar")
def borrar_instrumento(instrumento_id: int, user=Depends(require_roles("admin", "recepcion"))):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT envio_id, " + ", ".join([f"foto_entrada_{i}" for i in range(1,7)]) + " FROM instrumentos WHERE id=?", (instrumento_id,))
    row = cur.fetchone()
    if not row:
        conn.close()
        return HTMLResponse("Instrumento no encontrado", status_code=404)

    envio_id = row["envio_id"]
    fotos_to_delete = [row[f"foto_entrada_{i}"] for i in range(1,7)]

    cur.execute("DELETE FROM instrumentos WHERE id=?", (instrumento_id,))
    conn.commit()
    conn.close()

    for f in fotos_to_delete:
        _try_delete_public_photo(f)

    return RedirectResponse(url=f"/envios/{envio_id}", status_code=303)


# IMPORTAR EXCEL (admin/recepcion)
# -----------------------------
@app.post("/envios/{envio_id}/importar")
async def envio_importar_excel(
    envio_id: int,
    file: UploadFile = File(...),
    user=Depends(require_roles("admin", "recepcion")),
):
    """Importa instrumentos desde un Excel a un parte existente."""
    path = os.path.join(UPLOAD_DIR, file.filename)
    with open(path, "wb") as f:
        f.write(await file.read())

    # Reutilizamos la lógica de lectura existente
    try:
        _, _, df = leer_excel_envio(path)
    except Exception as e:
        return RedirectResponse(url=f"/envios/{envio_id}?err=excel&msg={str(e)}", status_code=303)

    conn = get_conn()
    cur = conn.cursor()

    rows = []
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for _, r in df.iterrows():
        # Normalizamos un poco los campos
        codigo = str(r.get("codigo_producto") or "").strip()
        if not codigo: continue # si no hay código, saltamos? o permitimos?

        rows.append((
            envio_id,
            codigo,
            str(r.get("fabricante") or "").strip(),
            str(r.get("num_serie") or "").strip(),
            str(r.get("denominacion") or "").strip(),
            str(r.get("observaciones") or "").strip(),
            str(r.get("codigo_datamatrix") or "").strip(),
            "", # nombre_trazabilidad
            "Pendiente",
            now_str,
        ))

    if rows:
        cur.executemany("""
            INSERT INTO instrumentos
            (envio_id, codigo_producto, fabricante, num_serie, denominacion, observaciones, codigo_datamatrix, nombre_trazabilidad, estado, creado_en)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, rows)
        conn.commit()

    conn.close()
    return RedirectResponse(url=f"/envios/{envio_id}?ok=import", status_code=303)


# -----------------------------
# IMPORTAR EXCEL (admin/recepcion) - CREANDO NUEVA OT
# -----------------------------
@app.get("/importar", response_class=HTMLResponse)
def importar_form(request: Request, user=Depends(require_roles("admin", "recepcion"))):
    return templates.TemplateResponse("importar.html", {"request": request, "user": user})


@app.post("/importar")
async def importar_excel(
    tipo_trabajo: str = Form("REPARACION"),
    file: UploadFile = File(...),
    user=Depends(require_roles("admin", "recepcion")),
):
    path = os.path.join(UPLOAD_DIR, file.filename)
    with open(path, "wb") as f:
        f.write(await file.read())

    try:
        cliente, fecha, df = leer_excel_envio(path)
    except Exception as e:
        return RedirectResponse(url=f"/importar?err=excel&msg={str(e)}", status_code=303)

    if not cliente or not str(cliente).strip():
        return RedirectResponse(url="/importar?err=cliente", status_code=303)

    conn = get_conn()
    cur = conn.cursor()

    ot_num = _next_ot_num(cur)
    tipo_trabajo = (tipo_trabajo or "REPARACION").strip().upper()
    if tipo_trabajo not in ("REPARACION", "TRAZABILIDAD", "OPTICA_RIGIDA"):
        tipo_trabajo = "REPARACION"

    has_tipo = _envios_has_column(cur, "tipo_trabajo")
    if has_tipo:
        cur.execute(
            "INSERT INTO envios (ot_num, nombre_archivo, cliente, fecha, tipo_trabajo) VALUES (?, ?, ?, ?, ?)",
            (ot_num, file.filename, cliente, fecha, tipo_trabajo),
        )
    else:
        cur.execute(
            "INSERT INTO envios (ot_num, nombre_archivo, cliente, fecha) VALUES (?, ?, ?, ?)",
            (ot_num, file.filename, cliente, fecha),
        )
    envio_id = cur.lastrowid

    rows = []
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for _, r in df.iterrows():
        rows.append((
            envio_id,
            r.get("codigo_producto"),
            r.get("fabricante"),
            r.get("num_serie"),
            r.get("denominacion"),
            r.get("observaciones"),
            r.get("codigo_datamatrix"),
            "",  # nombre_trazabilidad (faltaba para que el INSERT tenga 10 valores)
            "Pendiente",
            now_str,
        ))

    cur.executemany("""
        INSERT INTO instrumentos
        (envio_id, codigo_producto, fabricante, num_serie, denominacion, observaciones, codigo_datamatrix, nombre_trazabilidad, estado, creado_en)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, rows)

    conn.commit()
    conn.close()
    return RedirectResponse(url=f"/envios/{envio_id}", status_code=303)


# -----------------------------
# USUARIOS (admin)
# -----------------------------
@app.get("/usuarios", response_class=HTMLResponse)
def usuarios_list(request: Request, user=Depends(require_roles("admin"))):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT id, username, role, is_active, created_at FROM users ORDER BY id DESC;")
    users = [dict(r) for r in cur.fetchall()]
    conn.close()
    return templates.TemplateResponse("users.html", {"request": request, "user": user, "users": users})


@app.get("/usuarios/nuevo", response_class=HTMLResponse)
def usuarios_new_form(request: Request, user=Depends(require_roles("admin"))):
    return templates.TemplateResponse(
        "user_form.html",
        {"request": request, "user": user, "mode": "new", "u": {"username": "", "role": "tecnico", "is_active": 1}, "error": None},
    )


@app.post("/usuarios/nuevo")
def usuarios_new(
    request: Request,
    username: str = Form(...),
    password: str = Form(...),
    role: str = Form(...),
    is_active: int = Form(1),
    user=Depends(require_roles("admin")),
):
    username = (username or "").strip()

    if role not in ("admin", "recepcion", "tecnico", "grabado"):
        return templates.TemplateResponse(
            "user_form.html",
            {"request": request, "user": user, "mode": "new", "u": {"username": username, "role": role, "is_active": is_active}, "error": "Rol inválido"},
            status_code=400,
        )

    if len(username) < 3:
        return templates.TemplateResponse(
            "user_form.html",
            {"request": request, "user": user, "mode": "new", "u": {"username": username, "role": role, "is_active": is_active}, "error": "El usuario debe tener al menos 3 caracteres"},
            status_code=400,
        )

    if not password or len(password) < 6:
        return templates.TemplateResponse(
            "user_form.html",
            {"request": request, "user": user, "mode": "new", "u": {"username": username, "role": role, "is_active": is_active}, "error": "La contraseña debe tener al menos 6 caracteres"},
            status_code=400,
        )

    conn = get_conn()
    cur = conn.cursor()

    schema = _users_schema(cur)
    pw_col = schema.get("password_col")
    if not pw_col:
        conn.close()
        return templates.TemplateResponse(
            "user_form.html",
            {"request": request, "user": user, "mode": "new", "u": {"username": username, "role": role, "is_active": is_active}, "error": "Error de esquema DB"},
            status_code=400,
        )

    try:
        cols = ["username", pw_col, "role"]
        vals = [username, hash_password(password), role]
        
        if schema.get("has_is_active"):
            cols.append("is_active")
            vals.append(int(is_active))
            
        if schema.get("has_created_at"):
            cols.append("created_at")
            vals.append(_now_str())
        elif schema.get("has_created_at_at"):
            cols.append("created_at_at")
            vals.append(_now_str())

        sql = f"INSERT INTO users ({', '.join(cols)}) VALUES ({', '.join(['?']*len(cols))})"
        cur.execute(sql, tuple(vals))
        
        # Insert permissions default
        # (Generic logic copied from dash_users_nuevo, simplified here without RETURNING id logic as we redirect anyway)
        # But wait, we need ID to insert permissions.
        # So we MUST fetch ID.
        
        is_pg = bool(os.environ.get("DATABASE_URL"))
        new_id = None
        if is_pg:
             # On PG we need RETURNING id, but we already executed insert without it above? No wait.
             # We should use RETURNING id in the SAME execute or fetch lastrowid is unreliable on PG.
             # Let's redo the execute properly.
             pass
        
    except Exception:
        pass 
        
    # Re-writing the block properly to match dash_users_nuevo logic completely
    
    try:
        cols = ["username", pw_col, "role"]
        vals = [username, hash_password(password), role]

        if schema.get("has_is_active"):
            cols.append("is_active")
            vals.append(int(is_active))

        if schema.get("has_created_at"):
            cols.append("created_at")
            vals.append(_now_str())
        elif schema.get("has_created_at_at"):
            cols.append("created_at_at")
            vals.append(_now_str())

        sql = f"INSERT INTO users ({', '.join(cols)}) VALUES ({', '.join(['?']*len(cols))})"
        
        is_pg = bool(os.environ.get("DATABASE_URL"))
        if is_pg:
            sql += " RETURNING id"
            cur.execute(sql, tuple(vals))
            row = cur.fetchone()
            if row:
                new_id = int(row["id"])
            else:
                 raise Exception("No ID returned")
        else:
            cur.execute(sql, tuple(vals))
            new_id = int(cur.lastrowid)

        # Default permissions
        for action, _label in ACTIONS:
            allowed = _default_allowed_by_role(role, action)
            cur.execute(
                """
                INSERT INTO user_permissions (user_id, action, allowed) VALUES (?,?,?)
                ON CONFLICT (user_id, action) DO UPDATE SET allowed=excluded.allowed
                """,
                (new_id, action, int(allowed)),
            )

        conn.commit()
    except sqlite3.IntegrityError:
        conn.close()
        return templates.TemplateResponse(
            "user_form.html",
            {"request": request, "user": user, "mode": "new", "u": {"username": username, "role": role, "is_active": is_active}, "error": "Ese usuario ya existe"},
            status_code=400,
        )
    except Exception as e:
        conn.close()
        print(f"Error creating user: {e}")
        return templates.TemplateResponse(
            "user_form.html",
            {"request": request, "user": user, "mode": "new", "u": {"username": username, "role": role, "is_active": is_active}, "error": "Error interno de base de datos"},
            status_code=400,
        )

    conn.close()
    return RedirectResponse(url="/usuarios", status_code=303)


def _users_schema(cur) -> dict:
    from db import get_table_columns
    cols = get_table_columns(cur, "users")
    colset = set(cols)

    # columna de password
    if "password_hash" in colset:
        password_col = "password_hash"
    elif "password" in colset:
        password_col = "password"
    else:
        password_col = None

    return {
        "cols": colset,
        "password_col": password_col,
        "has_is_active": ("is_active" in colset),
        "has_created_at": ("created_at" in colset),
        "has_created_at_at": ("created_at_at" in colset),
    }

def _select_users_sql(schema: dict) -> str:
    parts = ["id", "username", "role"]

    if schema.get("has_is_active"):
        parts.append("is_active")
    else:
        parts.append("1 AS is_active")

    if schema.get("has_created_at"):
        parts.append("created_at")
    elif schema.get("has_created_at_at"):
        parts.append("created_at_at AS created_at")
    else:
        parts.append("NULL AS created_at")

    return "SELECT " + ", ".join(parts) + " FROM users ORDER BY id ASC"

def _now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# -----------------------------
# USUARIOS desde MODAL en DASHBOARD (solo admin)
# No toca tu gestor /usuarios existente.
# -----------------------------
@app.post("/dash_users/nuevo")
def dash_users_nuevo(
    username: str = Form(...),
    password: str = Form(...),
    role: str = Form(...),
    user=Depends(require_roles("admin")),
):
    username = (username or "").strip()
    if not username:
        return RedirectResponse(url="/?users=1&uerr=username", status_code=303)

    if role not in ("admin", "recepcion", "tecnico", "grabado"):
        return RedirectResponse(url="/?users=1&uerr=role", status_code=303)

    if not (password or "").strip():
        return RedirectResponse(url="/?users=1&uerr=pw", status_code=303)

    conn = get_conn()
    cur = conn.cursor()

    cur.execute("SELECT id FROM users WHERE username=?", (username,))
    if cur.fetchone():
        conn.close()
        return RedirectResponse(url="/?users=1&uerr=exists", status_code=303)

    schema = _users_schema(cur)
    pw_col = schema.get("password_col")
    if not pw_col:
        conn.close()
        return RedirectResponse(url="/?users=1&uerr=db", status_code=303)

    try:
        cols = ["username", pw_col, "role"]
        vals = [username, hash_password(password), role]

        if schema.get("has_is_active"):
            cols.append("is_active")
            vals.append(1)

        if schema.get("has_created_at"):
            cols.append("created_at")
            vals.append(_now_str())
        elif schema.get("has_created_at_at"):
            cols.append("created_at_at")
            vals.append(_now_str())

        sql = f"INSERT INTO users ({', '.join(cols)}) VALUES ({', '.join(['?']*len(cols))})"
        
        # Postgres necesita RETURNING id porque cursor.lastrowid no es estándar
        is_pg = bool(os.environ.get("DATABASE_URL"))
        if is_pg:
            sql += " RETURNING id"
            cur.execute(sql, tuple(vals))
            row = cur.fetchone()
            if not row:
                raise Exception("No se obtuvo ID del nuevo usuario (Postgres)")
            new_id = int(row["id"])
        else:
            cur.execute(sql, tuple(vals))
            new_id = int(cur.lastrowid)

        for action, _label in ACTIONS:
            allowed = _default_allowed_by_role(role, action)
            cur.execute(
                """
                INSERT INTO user_permissions (user_id, action, allowed) VALUES (?,?,?)
                ON CONFLICT (user_id, action) DO UPDATE SET allowed=excluded.allowed
                """,
                (new_id, action, int(allowed)),
            )

        conn.commit()
    except Exception as e:
        print(f"ERROR creando usuario: {e}")
        conn.close()
        return RedirectResponse(url="/?users=1&uerr=db", status_code=303)

    conn.close()
    return RedirectResponse(url="/?users=1&uok=created", status_code=303)



@app.post("/dash_users/{user_id}/role")
def dash_users_set_role(
    user_id: int,
    role: str = Form(...),
    user=Depends(require_roles("admin")),
):
    if role not in ("admin", "recepcion", "tecnico", "grabado"):
        return RedirectResponse(url="/?users=1&uerr=role", status_code=303)

    conn = get_conn()
    cur = conn.cursor()
    cur.execute("UPDATE users SET role=? WHERE id=?", (role, user_id))

    # Ajusta permisos a defaults del rol (sin borrar nada extra)
    for action, _label in ACTIONS:
        allowed = _default_allowed_by_role(role, action)
        cur.execute(
            """
            INSERT INTO user_permissions (user_id, action, allowed) VALUES (?,?,?)
            ON CONFLICT (user_id, action) DO UPDATE SET allowed=excluded.allowed
            """,
            (int(user_id), action, int(allowed)),
        )

    conn.commit()
    conn.close()
    return RedirectResponse(url="/?users=1&uok=role", status_code=303)


@app.post("/dash_users/{user_id}/toggle")
def dash_users_toggle_active(
    user_id: int,
    user=Depends(require_roles("admin")),
):
    conn = get_conn()
    cur = conn.cursor()
    schema = _users_schema(cur)
    if not schema.get("has_is_active"):
        conn.close()
        return RedirectResponse(url="/?users=1&uerr=db", status_code=303)

    cur.execute("SELECT is_active FROM users WHERE id=?", (user_id,))
    row = cur.fetchone()
    if not row:
        conn.close()
        return RedirectResponse(url="/?users=1&uerr=notfound", status_code=303)

    new_val = 0 if int(row["is_active"] or 0) == 1 else 1
    cur.execute("UPDATE users SET is_active=? WHERE id=?", (new_val, user_id))
    conn.commit()
    conn.close()
    return RedirectResponse(url="/?users=1&uok=active", status_code=303)



@app.post("/dash_users/{user_id}/password")
def dash_users_set_password(
    user_id: int,
    password: str = Form(...),
    user=Depends(require_roles("admin")),
):
    if not (password or "").strip():
        return RedirectResponse(url="/?users=1&uerr=pw", status_code=303)

    conn = get_conn()
    cur = conn.cursor()
    schema = _users_schema(cur)
    pw_col = schema.get("password_col")
    if not pw_col:
        conn.close()
        return RedirectResponse(url="/?users=1&uerr=db", status_code=303)

    cur.execute(f"UPDATE users SET {pw_col}=? WHERE id=?", (hash_password(password), user_id))
    conn.commit()
    conn.close()
    return RedirectResponse(url="/?users=1&uok=pw", status_code=303)



@app.post("/dash_users/{user_id}/perms")
async def dash_users_set_perms(
    request: Request,
    user_id: int,
    user=Depends(require_roles("admin")),
):
    form = await request.form()

    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT id FROM users WHERE id=?", (user_id,))
    if not cur.fetchone():
        conn.close()
        return RedirectResponse(url="/?users=1&uerr=notfound", status_code=303)

    for action, _label in ACTIONS:
        allowed = 1 if form.get(f"perm_{action}") == "on" else 0
        cur.execute(
            """
            INSERT INTO user_permissions (user_id, action, allowed) VALUES (?,?,?)
            ON CONFLICT (user_id, action) DO UPDATE SET allowed=excluded.allowed
            """,
            (int(user_id), action, int(allowed)),
        )

    conn.commit()
    conn.close()
    return RedirectResponse(url="/?users=1&uok=perms", status_code=303)

@app.post("/dash_users/{user_id}/delete")
def dash_users_delete(
    user_id: int,
    user=Depends(require_roles("admin")),
):
    # Evitar que el admin se borre a sí mismo
    if _user_id(user) == int(user_id):
        return RedirectResponse(url="/?users=1&uerr=selfdelete", status_code=303)

    conn = get_conn()
    cur = conn.cursor()

    cur.execute("SELECT id FROM users WHERE id=?", (user_id,))
    if not cur.fetchone():
        conn.close()
        return RedirectResponse(url="/?users=1&uerr=notfound", status_code=303)

    cur.execute("DELETE FROM user_permissions WHERE user_id=?", (user_id,))
    cur.execute("DELETE FROM users WHERE id=?", (user_id,))

    conn.commit()
    conn.close()
    return RedirectResponse(url="/?users=1&uok=deleted", status_code=303)


# -----------------------------
# BORRAR PARTE / ENVÍO (admin/recepcion)


# -----------------------------
# USUARIOS desde MODAL en DASHBOARD (solo admin)
# No toca tu gestor /usuarios existente.
# -----------------------------




@app.post("/dash_users/{user_id}/role")
def dash_users_set_role(
    user_id: int,
    role: str = Form(...),
    user=Depends(require_roles("admin")),
):
    if role not in ("admin", "recepcion", "tecnico", "grabado"):
        return RedirectResponse(url="/?users=1&uerr=role", status_code=303)

    conn = get_conn()
    cur = conn.cursor()
    cur.execute("UPDATE users SET role=? WHERE id=?", (role, user_id))

    # Ajusta permisos a defaults del rol (sin borrar nada extra)
    for action, _label in ACTIONS:
        allowed = _default_allowed_by_role(role, action)
        cur.execute(
            "INSERT OR REPLACE INTO user_permissions (user_id, action, allowed) VALUES (?,?,?)",
            (int(user_id), action, int(allowed)),
        )

    conn.commit()
    conn.close()
    return RedirectResponse(url="/?users=1&uok=role", status_code=303)


@app.post("/dash_users/{user_id}/toggle")
def dash_users_toggle_active(
    user_id: int,
    user=Depends(require_roles("admin")),
):
    conn = get_conn()
    cur = conn.cursor()
    schema = _users_schema(cur)
    if not schema.get("has_is_active"):
        conn.close()
        return RedirectResponse(url="/?users=1&uerr=db", status_code=303)

    cur.execute("SELECT is_active FROM users WHERE id=?", (user_id,))
    row = cur.fetchone()
    if not row:
        conn.close()
        return RedirectResponse(url="/?users=1&uerr=notfound", status_code=303)

    new_val = 0 if int(row["is_active"] or 0) == 1 else 1
    cur.execute("UPDATE users SET is_active=? WHERE id=?", (new_val, user_id))
    conn.commit()
    conn.close()
    return RedirectResponse(url="/?users=1&uok=active", status_code=303)



@app.post("/dash_users/{user_id}/password")
def dash_users_set_password(
    user_id: int,
    password: str = Form(...),
    user=Depends(require_roles("admin")),
):
    if not (password or "").strip():
        return RedirectResponse(url="/?users=1&uerr=pw", status_code=303)

    conn = get_conn()
    cur = conn.cursor()
    schema = _users_schema(cur)
    pw_col = schema.get("password_col")
    if not pw_col:
        conn.close()
        return RedirectResponse(url="/?users=1&uerr=db", status_code=303)

    cur.execute(f"UPDATE users SET {pw_col}=? WHERE id=?", (hash_password(password), user_id))
    conn.commit()
    conn.close()
    return RedirectResponse(url="/?users=1&uok=pw", status_code=303)



@app.post("/dash_users/{user_id}/perms")
async def dash_users_set_perms(
    request: Request,
    user_id: int,
    user=Depends(require_roles("admin")),
):
    form = await request.form()

    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT id FROM users WHERE id=?", (user_id,))
    if not cur.fetchone():
        conn.close()
        return RedirectResponse(url="/?users=1&uerr=notfound", status_code=303)

    for action, _label in ACTIONS:
        allowed = 1 if form.get(f"perm_{action}") == "on" else 0
        cur.execute(
            "INSERT OR REPLACE INTO user_permissions (user_id, action, allowed) VALUES (?,?,?)",
            (int(user_id), action, int(allowed)),
        )

    conn.commit()
    conn.close()
    return RedirectResponse(url="/?users=1&uok=perms", status_code=303)

# -----------------------------
# BORRAR PARTE / ENVÍO (admin/recepcion)
#  - Bloquea borrado si el parte está "cerrado" (sin pendientes y con instrumentos)
#  - Requiere confirmación fuerte: escribir la OT exacta
# -----------------------------
@app.post("/envios/{envio_id}/borrar")
def borrar_envio(
    request: Request,
    envio_id: int,
    confirm_ot: str = Form(""),
    user=Depends(require_roles("admin", "recepcion")),
):
    # Permiso granular (además del rol)
    conn_perm = get_conn()
    cur_perm = conn_perm.cursor()
    if not can_action(user, "envio_borrar", cur_perm):
        conn_perm.close()
        return RedirectResponse(url="/?err=perm", status_code=303)
    conn_perm.close()

    conn = get_conn()
    cur = conn.cursor()

    cur.execute("SELECT id, ot_num FROM envios WHERE id=?", (envio_id,))
    envio = cur.fetchone()
    if not envio:
        conn.close()
        return RedirectResponse(url="/?err=envio_no_encontrado", status_code=303)

    ot_num = (envio["ot_num"] or "").strip()

    # Confirmación fuerte
    if (confirm_ot or "").strip() != ot_num:
        conn.close()
        return RedirectResponse(url="/?err=confirm_ot", status_code=303)

    # Bloquear borrado si está cerrado:
    # "cerrado" = tiene instrumentos y ninguno está en Pendiente/En proceso
    cur.execute("SELECT COUNT(*) AS n FROM instrumentos WHERE envio_id=?", (envio_id,))
    n_inst = int(cur.fetchone()["n"] or 0)

    cur.execute(
        "SELECT COUNT(*) AS n FROM instrumentos WHERE envio_id=? AND estado IN ('Pendiente','En proceso')",
        (envio_id,),
    )
    n_pend = int(cur.fetchone()["n"] or 0)

    if n_inst > 0 and n_pend == 0:
        conn.close()
        return RedirectResponse(url="/?err=cerrado", status_code=303)

    # Borrar fotos de instrumentos del envío (evita archivos huérfanos)
    cur.execute(
        "SELECT foto_entrada_1, foto_entrada_2 FROM instrumentos WHERE envio_id=?",
        (envio_id,)
    )
    for r in cur.fetchall():
        _try_delete_public_photo(r["foto_entrada_1"])
        _try_delete_public_photo(r["foto_entrada_2"])

    # Borrar instrumentos y envío
    cur.execute("DELETE FROM instrumentos WHERE envio_id=?", (envio_id,))
    cur.execute("DELETE FROM envios WHERE id=?", (envio_id,))

    conn.commit()
    conn.close()

    return RedirectResponse(url="/?ok=borrado", status_code=303)
