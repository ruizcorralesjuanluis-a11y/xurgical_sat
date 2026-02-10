# security.py
from __future__ import annotations

from typing import Callable, Iterable, Optional, Dict, Any
from fastapi import Request, HTTPException, Depends
from db import get_conn


COOKIE_NAME = "xurgical_session"


def _redirect_to_login() -> None:
    """
    FastAPI permite redirección levantando HTTPException con status 303 + Location.
    Esto evita que el usuario vea {"detail":"No encontrado"} al entrar sin sesión.
    """
    raise HTTPException(status_code=303, headers={"Location": "/login"})


def _decode_session(request: Request) -> Optional[int]:
    """
    Devuelve user_id si el token es válido.
    Retorna None si falta token o si no puede decodificarse.

    Tu app crea el serializer en:
      app.state.serializer = make_serializer(app.state.secret_key)

    Y crea el token en login con:
      token = sign_session(app.state.serializer, user_id=int(u["id"]))
    """
    token = request.cookies.get(COOKIE_NAME)
    if not token:
        return None

    serializer = getattr(request.app.state, "serializer", None)
    if serializer is None:
        return None

    try:
        data = serializer.loads(token)
    except Exception:
        return None

    # Admitimos varios formatos (por robustez):
    # - {"user_id": 1}
    # - {"uid": 1}
    # - 1 (int)
    try:
        if isinstance(data, dict):
            if "user_id" in data:
                return int(data["user_id"])
            if "uid" in data:
                return int(data["uid"])
            if "id" in data:
                return int(data["id"])
            return None
        return int(data)
    except Exception:
        return None


def get_current_user(request: Request) -> Dict[str, Any]:
    """
    Dependency principal. Si no hay sesión válida -> redirige a /login.
    Si hay sesión -> devuelve dict con id/username/role.
    """
    user_id = _decode_session(request)
    if user_id is None:
        _redirect_to_login()

    conn = get_conn()
    cur = conn.cursor()

    # Ajusta columnas si tu tabla users es distinta. En tu app normalmente son:
    # id, username, role, password_hash ...
    cur.execute(
        "SELECT id, username, role, cliente_id FROM users WHERE id = ?",
        (user_id,),
    )
    row = cur.fetchone()
    if not row:
        _redirect_to_login()

    # sqlite3 puede devolver tuple o Row según cómo abras conexión.
    if isinstance(row, dict):
        return row

    # tuple -> dict
    return {"id": row[0], "username": row[1], "role": row[2], "cliente_id": row[3]}


def require_roles(*roles: str) -> Callable:
    """
    Dependency para proteger endpoints por rol.
    Uso en app.py:
        @app.get("/algo")
        def algo(user=Depends(require_roles("admin"))):
            ...
    """
    allowed = {r.strip().lower() for r in roles if r and r.strip()}

    def _dep(user: Dict[str, Any] = Depends(get_current_user)) -> Dict[str, Any]:
        role = str(user.get("role", "")).strip().lower()

        # Si no se especifican roles, se permite
        if not allowed:
            return user

        if role not in allowed:
            # Para web: podrías redirigir a /login o devolver 403.
            # Mantengo 403 porque ya estás logueado pero sin permisos.
            raise HTTPException(status_code=403, detail="Sin permisos")

        return user

    return _dep
# security.py
from fastapi import Request, HTTPException, Depends
from auth_utils import make_serializer, read_session

def get_current_user(request: Request):
    token = request.cookies.get("xurgical_session")
    if not token:
        raise HTTPException(status_code=401, detail="Not authenticated")

    serializer = make_serializer(request.app.state.secret_key)
    user_id = read_session(serializer, token)
    if not user_id:
        raise HTTPException(status_code=401, detail="Invalid session")

    from db import get_conn  # import local para evitar ciclos
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT id, username, role, is_active, cliente_id FROM users WHERE id=?", (user_id,))
    u = cur.fetchone()
    conn.close()

    if not u or int(u["is_active"] or 0) != 1:
        raise HTTPException(status_code=401, detail="User inactive")

    return dict(u)

def require_roles(*roles: str):
    def _guard(user=Depends(get_current_user)):
        if user["role"] not in roles:
            raise HTTPException(status_code=403, detail="Forbidden")
        return user
    return _guard
