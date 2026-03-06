import os
import sys
from pathlib import Path
from fastapi import HTTPException
from db import get_conn, get_table_columns

# Configuración de rutas centralizada
BASE_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
UPLOAD_DIR = os.environ.get("XURGICAL_UPLOAD_DIR", "uploads")
FOTOS_DIR = os.environ.get("XURGICAL_FOTOS_DIR", os.path.join(UPLOAD_DIR, "fotos"))

# Asegurar existencia de directorios (especialmente en Render)
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(FOTOS_DIR, exist_ok=True)

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
    try:
        if hasattr(user, "keys"):
            return user.get("role") if hasattr(user, "get") else user["role"]
    except Exception:
        pass
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
    role = (role or "").strip().lower()
    if role == "admin":
        return 1
    if role == "socio":
        return 1 if action.endswith("_ver") or action == "dashboard_ver" else 0
    if role == "recepcion":
        return 1 if action in {"dashboard_ver","envio_crear","instrumento_crear","fotos_gestionar","excel_importar"} else 0
    if role == "tecnico":
        return 1 if action in {"dashboard_ver","instrumento_editar","fotos_gestionar","estado_cambiar"} else 0
    if role == "grabado":
        return 1 if action in {"dashboard_ver"} else 0
    if role == "cliente":
        return 1 if action in {"dashboard_ver"} else 0
    return 0

def _get_user_permissions_map(cur, user_id: int) -> dict:
    cur.execute("SELECT action, allowed FROM user_permissions WHERE user_id=?", (user_id,))
    return {r["action"]: int(r["allowed"] or 0) for r in cur.fetchall()}

def can_action(user, action: str, cur=None) -> bool:
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
        return _default_allowed_by_role(role, action) == 1

def _users_schema(cur) -> dict:
    """Detecta las columnas reales en la tabla users."""
    cols = get_table_columns(cur, "users")
    colset = set(cols)
    return {
        "password_col": "password_hash" if "password_hash" in colset else ("password" if "password" in colset else None),
        "has_is_active": "is_active" in colset
    }

def _select_users_sql(schema: dict) -> str:
    pw_col = schema.get("password_col") or "password"
    active_part = ", is_active" if schema.get("has_is_active") else ", 1 AS is_active"
    return f"SELECT id, username, {pw_col} AS password_hash, role {active_part} FROM users ORDER BY username ASC"

def _envios_has_column(cur, col: str) -> bool:
    try:
        from db import get_table_columns
        cols = get_table_columns(cur, "envios")
        return col in cols
    except Exception:
        return False

def _get_cliente(cur, cliente_id: int) -> dict | None:
    cur.execute(
        "SELECT id, numero_cliente, nombre, prefijo, prefijo_nombre, email, ultimo_numero FROM clientes WHERE id=?",
        (int(cliente_id),),
    )
    r = cur.fetchone()
    return dict(r) if r else None

def _reserve_numeros_cliente(cur, cliente_id: int, cantidad: int) -> tuple[str, str, list[int]]:
    if cantidad <= 0:
        return "", "", []
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

def _list_clientes(cur) -> list[dict]:
    import sqlite3
    try:
        import psycopg2
        PG_ERR = psycopg2.Error
    except ImportError:
        PG_ERR = Exception

    sql = "SELECT id, numero_cliente, nombre, prefijo, prefijo_nombre, email, ultimo_numero FROM clientes ORDER BY LOWER(nombre) ASC"
    try:
        cur.execute(sql)
        return [dict(r) for r in cur.fetchall()]
    except (sqlite3.OperationalError, PG_ERR) as e:
        err_msg = str(e).lower()
        if "no such table" in err_msg or "does not exist" in err_msg or "no such column" in err_msg or "column" in err_msg:
            from db import init_db
            init_db()
            cur.execute(sql)
            return [dict(r) for r in cur.fetchall()]
        raise
