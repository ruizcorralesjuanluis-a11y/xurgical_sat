import os
from datetime import datetime
from typing import Optional, List, Dict, Any

from fastapi import APIRouter, Request, Form, Depends, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse

from db import get_conn, get_table_columns
from security import get_current_user, require_roles, hash_password
from shared import (
    _user_role, _user_id, can_action, ACTIONS, _default_allowed_by_role,
    _users_schema, _select_users_sql, _list_clientes
)

router = APIRouter()

# Helper para fechas
def _now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

@router.get("/usuarios", response_class=HTMLResponse)
def usuarios_list(request: Request, user=Depends(require_roles("admin"))):
    templates = request.app.state.templates
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT id, username, role, is_active, created_at FROM users ORDER BY id DESC;")
    users = [dict(r) for r in cur.fetchall()]
    conn.close()
    return templates.TemplateResponse("users.html", {"request": request, "user": user, "users": users})

@router.get("/usuarios/nuevo", response_class=HTMLResponse)
def usuarios_new_form(request: Request, user=Depends(require_roles("admin"))):
    templates = request.app.state.templates
    conn = get_conn()
    cur = conn.cursor()
    clientes = _list_clientes(cur)
    conn.close()
    return templates.TemplateResponse(
        "user_form.html",
        {"request": request, "user": user, "mode": "new", "u": {"username": "", "role": "tecnico", "is_active": 1, "cliente_id": None}, "clientes": clientes, "error": None},
    )

@router.post("/usuarios/nuevo")
def usuarios_new(
    request: Request,
    username: str = Form(...),
    password: str = Form(...),
    role: str = Form(...),
    is_active: int = Form(1),
    cliente_id: Optional[str] = Form(None),
    user=Depends(require_roles("admin")),
):
    templates = request.app.state.templates
    username = (username or "").strip()
    if role not in ("admin", "recepcion", "tecnico", "grabado", "cliente", "socio"):
        return templates.TemplateResponse("user_form.html", {"request": request, "user": user, "mode": "new", "u": {"username": username, "role": role}, "error": "Rol inválido"}, status_code=400)
    if len(username) < 3:
        return templates.TemplateResponse("user_form.html", {"request": request, "user": user, "mode": "new", "u": {"username": username, "role": role}, "error": "Al menos 3 caracteres"}, status_code=400)
    
    conn = get_conn()
    cur = conn.cursor()
    schema = _users_schema(cur)
    pw_col = schema.get("password_col") or "password_hash"
    
    try:
        cli_id_val = int(cliente_id) if cliente_id else None
        cols = ["username", pw_col, "role", "cliente_id"]
        vals = [username, hash_password(password), role, cli_id_val]
        if schema.get("has_is_active"):
            cols.append("is_active"); vals.append(int(is_active))
        
        sql = f"INSERT INTO users ({', '.join(cols)}) VALUES ({', '.join(['?']*len(cols))})"
        is_pg = bool(os.environ.get("DATABASE_URL"))
        if is_pg:
            cur.execute(sql + " RETURNING id", tuple(vals))
            row = cur.fetchone()
            new_id = int(row["id"]) if row else None
        else:
            cur.execute(sql, tuple(vals))
            new_id = cur.lastrowid
            
        for action, _ in ACTIONS:
            allowed = _default_allowed_by_role(role, action)
            cur.execute("INSERT INTO user_permissions (user_id, action, allowed) VALUES (?,?,?) ON CONFLICT (user_id, action) DO UPDATE SET allowed=excluded.allowed", (new_id, action, int(allowed)))
        conn.commit()
    except Exception as e:
        conn.close()
        return templates.TemplateResponse("user_form.html", {"request": request, "user": user, "mode": "new", "u": {"username": username, "role": role}, "error": str(e)}, status_code=400)
    conn.close()
    return RedirectResponse(url="/usuarios", status_code=303)

@router.post("/dash_users/nuevo")
def dash_users_nuevo(username: str = Form(...), password: str = Form(...), role: str = Form(...), cliente_id: Optional[str] = Form(None), user=Depends(require_roles("admin"))):
    username = (username or "").strip()
    conn = get_conn(); cur = conn.cursor()
    schema = _users_schema(cur); pw_col = schema.get("password_col") or "password_hash"
    try:
        cli_id_val = int(cliente_id) if cliente_id else None
        cols = ["username", pw_col, "role", "cliente_id"]
        vals = [username, hash_password(password), role, cli_id_val]
        sql = f"INSERT INTO users ({', '.join(cols)}) VALUES ({', '.join(['?']*len(cols))})"
        is_pg = bool(os.environ.get("DATABASE_URL"))
        if is_pg:
            cur.execute(sql + " RETURNING id", tuple(vals))
            new_id = int(cur.fetchone()["id"])
        else:
            cur.execute(sql, tuple(vals)); new_id = cur.lastrowid
        for action, _ in ACTIONS:
            allowed = _default_allowed_by_role(role, action)
            cur.execute("INSERT INTO user_permissions (user_id, action, allowed) VALUES (?,?,?) ON CONFLICT (user_id, action) DO UPDATE SET allowed=excluded.allowed", (new_id, action, int(allowed)))
        conn.commit()
    except Exception as e:
        conn.close(); return RedirectResponse(url="/?users=1&uerr=db", status_code=303)
    conn.close()
    return RedirectResponse(url="/?users=1&uok=created", status_code=303)

@router.post("/dash_users/{user_id}/role")
def dash_users_set_role(user_id: int, role: str = Form(...), user=Depends(require_roles("admin"))):
    conn = get_conn(); cur = conn.cursor()
    cur.execute("UPDATE users SET role=? WHERE id=?", (role, user_id))
    for action, _ in ACTIONS:
        allowed = _default_allowed_by_role(role, action)
        cur.execute("INSERT INTO user_permissions (user_id, action, allowed) VALUES (?,?,?) ON CONFLICT (user_id, action) DO UPDATE SET allowed=excluded.allowed", (user_id, action, int(allowed)))
    conn.commit(); conn.close()
    return RedirectResponse(url="/?users=1&uok=role", status_code=303)

@router.post("/dash_users/{user_id}/toggle")
def dash_users_toggle_active(user_id: int, user=Depends(require_roles("admin"))):
    conn = get_conn(); cur = conn.cursor()
    cur.execute("UPDATE users SET is_active = 1 - COALESCE(is_active, 1) WHERE id=?", (user_id,))
    conn.commit(); conn.close()
    return RedirectResponse(url="/?users=1&uok=active", status_code=303)

@router.post("/dash_users/{user_id}/perms")
async def dash_users_set_perms(request: Request, user_id: int, user=Depends(require_roles("admin"))):
    form = await request.form()
    conn = get_conn(); cur = conn.cursor()
    for action, _ in ACTIONS:
        allowed = 1 if form.get(f"perm_{action}") == "on" else 0
        cur.execute("INSERT INTO user_permissions (user_id, action, allowed) VALUES (?,?,?) ON CONFLICT (user_id, action) DO UPDATE SET allowed=excluded.allowed", (user_id, action, allowed))
    conn.commit(); conn.close()
    return RedirectResponse(url="/?users=1&uok=perms", status_code=303)

@router.post("/dash_users/{user_id}/delete")
def dash_users_delete(user_id: int, user=Depends(require_roles("admin"))):
    if _user_id(user) == int(user_id): return RedirectResponse(url="/?users=1&uerr=selfdelete", status_code=303)
    conn = get_conn(); cur = conn.cursor()
    cur.execute("DELETE FROM user_permissions WHERE user_id=?", (user_id,))
    cur.execute("DELETE FROM users WHERE id=?", (user_id,))
    conn.commit(); conn.close()
    return RedirectResponse(url="/?users=1&uok=deleted", status_code=303)

@router.get("/checklist_admin", response_class=HTMLResponse)
def checklist_admin(request: Request, tipo: str = "REPARACION", user=Depends(require_roles("admin", "socio"))):
    templates = request.app.state.templates
    tipo = (tipo or "REPARACION").strip().upper()
    conn = get_conn(); cur = conn.cursor()
    cur.execute("SELECT id, nombre, orden, COALESCE(activo,1) AS activo, COALESCE(tipo_trabajo,'REPARACION') AS tipo_trabajo FROM checklist_items WHERE COALESCE(tipo_trabajo,'REPARACION') = ? ORDER BY LOWER(nombre) ASC", (tipo,))
    items = [dict(r) for r in cur.fetchall()]
    conn.close()
    return templates.TemplateResponse("checklist_admin.html", {"request": request, "user": user, "tipo": tipo, "items": items, "tipos": ["REPARACION", "OPTICA_RIGIDA", "TRAZABILIDAD"]})

@router.get("/repuestos_catalogo", response_class=HTMLResponse)
def view_repuestos_catalogo(request: Request, user=Depends(require_roles("admin", "socio"))):
    templates = request.app.state.templates
    conn = get_conn(); cur = conn.cursor()
    cur.execute("SELECT * FROM repuestos_catalogo ORDER BY nombre")
    items = [dict(r) for r in cur.fetchall()]
    conn.close()
    return templates.TemplateResponse("repuestos_catalogo.html", {"request": request, "user": user, "items": items})

@router.get("/estadisticas_tecnicos", response_class=HTMLResponse)
def estadisticas_tecnicos(request: Request, fecha_inicio: Optional[str] = None, fecha_fin: Optional[str] = None, user=Depends(require_roles("admin", "socio"))):
    templates = request.app.state.templates
    conn = get_conn(); cur = conn.cursor()
    hoy = datetime.now()
    if not fecha_inicio: fecha_inicio = hoy.replace(day=1).strftime("%Y-%m-%d")
    if not fecha_fin: fecha_fin = hoy.strftime("%Y-%m-%d")
    sql = """
        SELECT tecnico_reparacion as tecnico,
            COUNT(CASE WHEN estado = 'Reparado' THEN 1 END) as reparados,
            COUNT(CASE WHEN estado = 'Baja' THEN 1 END) as bajas,
            COUNT(CASE WHEN estado NOT IN ('Reparado', 'Baja') THEN 1 END) as curso,
            COUNT(*) as total
        FROM instrumentos
        WHERE tecnico_reparacion IS NOT NULL AND tecnico_reparacion != ''
          AND tecnico_reparacion_en >= ? AND tecnico_reparacion_en <= ?
        GROUP BY tecnico_reparacion ORDER BY reparados DESC
    """
    cur.execute(sql, (f"{fecha_inicio} 00:00:00", f"{fecha_fin} 23:59:59"))
    stats = [dict(r) for r in cur.fetchall()]
    conn.close()
    return templates.TemplateResponse("estadisticas_tecnicos.html", {"request": request, "user": user, "stats": stats, "fecha_inicio": fecha_inicio, "fecha_fin": fecha_fin})
