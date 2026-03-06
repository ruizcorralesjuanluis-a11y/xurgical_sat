import os
import re
import time
import uuid
from datetime import datetime
from typing import Optional, List, Dict, Any

from fastapi import APIRouter, Request, Form, File, UploadFile, Depends, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse, JSONResponse

from db import get_conn
from security import get_current_user, require_roles
from shared import (
    _user_role, _user_id, _list_clientes, _get_cliente
)

router = APIRouter()

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FOTOS_DIR = os.path.join(BASE_DIR, "static", "fotos")

# -----------------------------
# CLIENTES (admin/recepcion)
# -----------------------------
@router.get("/clientes", response_class=HTMLResponse)
def clientes_list(request: Request, user=Depends(require_roles("admin", "recepcion", "socio"))):
    templates = request.app.state.templates
    conn = get_conn()
    cur = conn.cursor()
    clientes = _list_clientes(cur)
    conn.close()
    return templates.TemplateResponse(
        "clientes_list.html",
        {"request": request, "user": user, "clientes": clientes},
    )

@router.get("/clientes/nuevo", response_class=HTMLResponse)
def clientes_nuevo_form(request: Request, user=Depends(require_roles("admin", "recepcion", "socio"))):
    templates = request.app.state.templates
    return templates.TemplateResponse(
        "clientes_form.html",
        {"request": request, "user": user, "mode": "new", "cliente": None},
    )

@router.post("/clientes/nuevo")
def clientes_nuevo_crear(
    nombre: str = Form(""),
    numero_cliente: int = Form(None),
    prefijo: str = Form(""),
    email: str = Form(""),
    prefijo_nombre: str = Form(""),
    ultimo_numero: int = Form(0),
    user=Depends(require_roles("admin", "recepcion")),
):
    nombre = (nombre or "").strip()
    if not nombre:
        return RedirectResponse(url="/clientes?err=nombre", status_code=303)
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT id FROM clientes WHERE nombre = ?", (nombre,))
    if cur.fetchone():
        conn.close()
        return RedirectResponse(url="/clientes?err=exists", status_code=303)
    if not numero_cliente:
        cur.execute("SELECT MAX(numero_cliente) as vmax FROM clientes")
        vmax = cur.fetchone()["vmax"]
        numero_cliente = 1 if vmax is None else int(vmax) + 1
    cur.execute(
        "INSERT INTO clientes (numero_cliente, nombre, prefijo, email, prefijo_nombre, ultimo_numero) VALUES (?, ?, ?, ?, ?, ?)",
        (numero_cliente, nombre, (prefijo or "").strip(), (email or "").strip(), (prefijo_nombre or "").strip(), int(ultimo_numero or 0)),
    )
    conn.commit()
    conn.close()
    return RedirectResponse(url="/clientes", status_code=303)

@router.get("/clientes/{cliente_id}/editar", response_class=HTMLResponse)
def clientes_editar_form(request: Request, cliente_id: int, user=Depends(require_roles("admin", "recepcion", "socio"))):
    templates = request.app.state.templates
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

@router.post("/clientes/{cliente_id}/editar")
def clientes_editar_guardar(
    cliente_id: int,
    nombre: str = Form(""),
    numero_cliente: int = Form(None),
    prefijo: str = Form(""),
    email: str = Form(""),
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
        "UPDATE clientes SET numero_cliente=?, nombre=?, prefijo=?, email=?, prefijo_nombre=?, ultimo_numero=? WHERE id=?",
        (numero_cliente, nombre, (prefijo or "").strip(), (email or "").strip(), (prefijo_nombre or "").strip(), int(ultimo_numero or 0), int(cliente_id)),
    )
    conn.commit()
    conn.close()
    return RedirectResponse(url="/clientes", status_code=303)

# -----------------------------
# RECOGIDAS
# -----------------------------
@router.post("/peticion_recogida")
async def peticion_recogida(
    n_instrumentos: int = Form(0),
    contacto: str = Form(""),
    telefono: str = Form(""),
    observaciones: str = Form(""),
    user=Depends(require_roles("cliente", "admin", "recepcion"))
):
    cli_id = (user.get("cliente_id") if isinstance(user, dict) else getattr(user, "cliente_id", None))
    u_id = (user.get("id") if isinstance(user, dict) else getattr(user, "id", None))
    if not cli_id:
        return RedirectResponse(url="/?err=nocli", status_code=303)
    conn = get_conn()
    cur = conn.cursor()
    now = datetime.now()
    prefix = f"REC-{now.strftime('%y%m')}-"
    cur.execute("SELECT num_peticion FROM peticiones_recogida WHERE num_peticion LIKE ? ORDER BY id DESC LIMIT 1", (f"{prefix}%",))
    row = cur.fetchone()
    new_num = 1
    if row and row["num_peticion"]:
        try:
            last_num = int(row["num_peticion"].split("-")[-1])
            new_num = last_num + 1
        except: pass
    num_peticion = f"{prefix}{str(new_num).zfill(3)}"
    cur.execute("""
        INSERT INTO peticiones_recogida (num_peticion, cliente_id, usuario_id, n_instrumentos, contacto, telefono, observaciones, estado, creado_en)
        VALUES (?, ?, ?, ?, ?, ?, ?, 'Pendiente', CURRENT_TIMESTAMP)
    """, (num_peticion, int(cli_id or 0), u_id, n_instrumentos, contacto, telefono, observaciones))
    conn.commit()
    conn.close()
    return RedirectResponse(url=f"/?msg=recogida_ok&num={num_peticion}", status_code=303)

@router.get("/recogidas", response_class=HTMLResponse)
def recogidas_list(request: Request, user=Depends(get_current_user)):
    templates = request.app.state.templates
    role = _user_role(user)
    if role not in ["admin", "recepcion", "cliente", "socio"]:
        return RedirectResponse(url="/?err=perm", status_code=303)
    conn = get_conn()
    cur = conn.cursor()
    where_sql, params = "", []
    if role == "cliente":
        cli_id = (user.get("cliente_id") if isinstance(user, dict) else getattr(user, "cliente_id", None))
        where_sql, params = "WHERE pr.cliente_id = ?", [int(cli_id or 0)]
    cur.execute(f"SELECT pr.*, c.nombre as cliente_nombre FROM peticiones_recogida pr JOIN clientes c ON pr.cliente_id = c.id {where_sql} ORDER BY pr.creado_en DESC", tuple(params))
    items = [dict(r) for r in cur.fetchall()]
    conn.close()
    return templates.TemplateResponse("recogidas_list.html", {"request": request, "user": user, "items": items})

@router.post("/peticion_recogida/{peticion_id}/completar")
def completar_peticion_recogida(peticion_id: int, user=Depends(require_roles("admin", "recepcion"))):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("UPDATE peticiones_recogida SET estado = 'Completada' WHERE id = ?", (peticion_id,))
    conn.commit()
    conn.close()
    return RedirectResponse(url="/", status_code=303)

@router.get("/api/peticiones_recogida/count")
def count_peticiones_recogida(user=Depends(require_roles("admin", "recepcion"))):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) as n FROM peticiones_recogida WHERE estado = 'Pendiente'")
    row = cur.fetchone()
    conn.close()
    return {"count": int(row["n"] or 0) if row else 0}

# -----------------------------
# CONSULTAS TÉCNICAS
# -----------------------------
@router.post("/consultas/nueva")
async def consulta_nueva(
    titulo: str = Form(...),
    descripcion: str = Form(...),
    foto1: UploadFile = File(None),
    foto2: UploadFile = File(None),
    foto3: UploadFile = File(None),
    user=Depends(get_current_user)
):
    u_id = (user.get("id") if isinstance(user, dict) else getattr(user, "id", None))
    cli_id = (user.get("cliente_id") if isinstance(user, dict) else getattr(user, "cliente_id", None))
    if not cli_id and _user_role(user) == "cliente":
        return RedirectResponse(url="/?err=nocli", status_code=303)
    fotos_urls = []
    for f in [foto1, foto2, foto3]:
        if f and f.filename:
            safe_name = re.sub(r"[^a-zA-Z0-9.-]", "_", f.filename)
            fname = f"consulta_{int(time.time())}_{uuid.uuid4().hex[:6]}_{safe_name}"
            path = os.path.join(FOTOS_DIR, fname)
            with open(path, "wb") as out: out.write(await f.read())
            fotos_urls.append(f"/static/fotos/{fname}")
        else: fotos_urls.append(None)
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("INSERT INTO consultas (cliente_id, usuario_id, titulo, descripcion, foto_1, foto_2, foto_3) VALUES (?, ?, ?, ?, ?, ?, ?)",
                (cli_id or 0, u_id, titulo, descripcion, fotos_urls[0], fotos_urls[1], fotos_urls[2]))
    conn.commit()
    conn.close()
    return RedirectResponse(url="/?msg=consulta_ok", status_code=303)

@router.get("/consultas/{consulta_id}", response_class=HTMLResponse)
async def consulta_detalle(request: Request, consulta_id: int, user=Depends(get_current_user)):
    templates = request.app.state.templates
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT c.*, cl.nombre as cliente_nombre FROM consultas c LEFT JOIN clientes cl ON c.cliente_id = cl.id WHERE c.id = ?", (consulta_id,))
    consulta = cur.fetchone()
    if not consulta:
        conn.close(); return HTMLResponse("No encontrada", status_code=404)
    if _user_role(user) == "cliente":
        u_cli_id = (user.get("cliente_id") if isinstance(user, dict) else getattr(user, "cliente_id", None))
        if int(consulta.get("cliente_id", 0)) != int(u_cli_id or -1):
            conn.close(); return HTMLResponse("Denegado", status_code=403)
    cur.execute("SELECT m.*, u.username, u.role FROM consultas_mensajes m JOIN users u ON m.usuario_id = u.id WHERE m.consulta_id = ? ORDER BY m.creado_en ASC", (consulta_id,))
    mensajes = [dict(r) for r in cur.fetchall()]
    conn.close()
    return templates.TemplateResponse("consulta_chat.html", {"request": request, "user": user, "consulta": dict(consulta), "mensajes": mensajes})

@router.post("/consultas/{consulta_id}/mensaje")
async def consulta_mensaje_enviar(consulta_id: int, mensaje: str = Form(...), user=Depends(get_current_user)):
    u_id = (user.get("id") if isinstance(user, dict) else getattr(user, "id", None))
    conn = get_conn(); cur = conn.cursor()
    nuevo_estado = "Respondida" if _user_role(user) in ["admin", "recepcion", "tecnico"] else "Abierta"
    cur.execute("INSERT INTO consultas_mensajes (consulta_id, usuario_id, mensaje) VALUES (?, ?, ?)", (consulta_id, u_id, mensaje))
    cur.execute("UPDATE consultas SET estado = ?, actualizado_en = CURRENT_TIMESTAMP WHERE id = ?", (nuevo_estado, consulta_id))
    conn.commit(); conn.close()
    return RedirectResponse(url=f"/consultas/{consulta_id}", status_code=303)

@router.get("/api/consultas/unread_count")
def count_consultas_unread(user=Depends(get_current_user)):
    conn = get_conn(); cur = conn.cursor()
    role = _user_role(user)
    count = 0
    if role in ["admin", "recepcion", "tecnico"]:
        cur.execute("SELECT COUNT(*) as n FROM consultas WHERE estado='Abierta'")
        count = cur.fetchone()["n"]
    elif role == "cliente" and user.get("cliente_id"):
        cur.execute("SELECT COUNT(*) as n FROM consultas WHERE cliente_id=? AND estado='Respondida'", (int(user.get("cliente_id")),))
        count = cur.fetchone()["n"]
    conn.close()
    return {"count": int(count or 0)}
