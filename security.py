# security.py
from __future__ import annotations
import os
from typing import Callable, Optional, Dict, Any
from fastapi import Request, HTTPException, Depends
from db import get_conn
from auth_utils import (
    hash_password, 
    verify_password, 
    make_serializer, 
    sign_session, 
    read_session as decode_session
)

COOKIE_NAME = "xurgical_session"

def _redirect_to_login() -> None:
    raise HTTPException(status_code=303, headers={"Location": "/login"})

def _get_user_id_from_session(request: Request) -> Optional[int]:
    token = request.cookies.get(COOKIE_NAME)
    if not token:
        return None

    serializer = getattr(request.app.state, "serializer", None)
    if serializer is None:
        # Fallback si no está en state
        serializer = make_serializer(os.environ.get("XURGICAL_SECRET_KEY", "dev"))
    
    return decode_session(serializer, token)

def get_current_user(request: Request) -> Dict[str, Any]:
    user_id = _get_user_id_from_session(request)
    if user_id is None:
        _redirect_to_login()

    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "SELECT id, username, role, is_active, cliente_id FROM users WHERE id = ?",
        (user_id,),
    )
    u = cur.fetchone()
    conn.close()

    if not u or int(u.get("is_active", 0) or 0) != 1:
        _redirect_to_login()

    return dict(u)

def require_roles(*roles: str) -> Callable:
    allowed = {r.strip().lower() for r in roles if r and r.strip()}

    def _dep(user: Dict[str, Any] = Depends(get_current_user)) -> Dict[str, Any]:
        role = str(user.get("role", "")).strip().lower()
        if not allowed:
            return user
        if role not in allowed:
            raise HTTPException(status_code=403, detail="Sin permisos")
        return user

    return _dep
