import os
import io
import re
import time
from datetime import datetime
from fastapi import APIRouter, Request, Depends, Form, File, UploadFile, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse, JSONResponse, FileResponse
from db import get_conn, get_table_columns
from security import get_current_user, require_roles, hash_password
from shared import (
    _user_role, _user_id, can_action, _get_cliente, _reserve_numeros_cliente, _envios_has_column, ACTIONS, _default_allowed_by_role
)
from utils import _build_nombre_trazabilidad, _clean_trz
import base64
import uuid

router = APIRouter()

@router.get("/envios/{envio_id}/revision", response_class=HTMLResponse)
def envio_revision(request: Request, envio_id: int, user=Depends(require_roles("admin", "recepcion", "socio"))):
    templates = request.app.state.templates
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT * FROM envios WHERE id=?", (envio_id,))
    envio = cur.fetchone()
    if not envio: return HTMLResponse("Envío no encontrado", status_code=404)
    cur.execute("SELECT i.* FROM instrumentos i WHERE i.envio_id=? ORDER BY i.id ASC", (envio_id,))
    instrumentos = []
    for r in cur.fetchall():
        d = dict(r)
        for k, v in d.items():
            if hasattr(v, "isoformat"): d[k] = v.isoformat()
        instrumentos.append(d)
    n_revisados = sum(1 for i in instrumentos if i["revisado"])
    total = len(instrumentos)
    conn.close()
    return templates.TemplateResponse("envio_revision.html", {
        "request": request, "user": user, "envio": dict(envio),
        "instrumentos": instrumentos, "total": total, "n_revisados": n_revisados,
    })

@router.post("/instrumentos/{instrumento_id}/revisar")
def revisar_instrumento(instrumento_id: int, user=Depends(require_roles("admin", "recepcion"))):
    conn = get_conn()
    cur = conn.cursor()
    username = user.get("username") if isinstance(user, dict) else getattr(user, "username", "S/N")
    cur.execute("UPDATE instrumentos SET revisado=1, revisado_por=?, revisado_en=CURRENT_TIMESTAMP WHERE id=?", 
                (username, instrumento_id))
    conn.commit()
    conn.close()
    return {"ok": True}

@router.post("/instrumentos/{instrumento_id}/desrevisar")
def desrevisar_instrumento(instrumento_id: int, user=Depends(require_roles("admin", "recepcion"))):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("UPDATE instrumentos SET revisado=0, revisado_por=NULL, revisado_en=NULL WHERE id=?", (instrumento_id,))
    conn.commit()
    conn.close()
    return {"ok": True}

@router.get("/envios/{envio_id}/instrumentos/nuevo", response_class=HTMLResponse)
def instrumento_nuevo_form(request: Request, envio_id: int, user=Depends(require_roles("admin", "recepcion", "socio"))):
    templates = request.app.state.templates
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT * FROM envios WHERE id=?", (envio_id,))
    envio = cur.fetchone()
    conn.close()
    if not envio: return HTMLResponse("Envío no encontrado", status_code=404)
    return templates.TemplateResponse("instrumento_nuevo.html", {
        "request": request, "user": user, "mode": "new", "envio": dict(envio), "inst": None,
    })

@router.post("/envios/{envio_id}/instrumentos/nuevo")
async def instrumento_nuevo_crear(
    envio_id: int,
    codigo_producto: str = Form(""),
    fabricante: str = Form(""),
    num_serie: str = Form(""),
    denominacion: str = Form(""),
    observaciones: str = Form(""),
    codigo_datamatrix: str = Form(""),
    codigo_cliente: str = Form(""),
    unidades: int = Form(1),
    user=Depends(require_roles("admin", "recepcion")),
):
    conn = get_conn()
    cur = conn.cursor()
    if not can_action(user, "instrumento_crear", cur):
        conn.close()
        return RedirectResponse(url="/?err=perm", status_code=303)
    cur.execute("SELECT id, cliente_id, tipo_trabajo FROM envios WHERE id=?", (envio_id,))
    e = cur.fetchone()
    if not e:
        conn.close()
        return HTMLResponse("Envío no encontrado", status_code=404)
    unidades = max(1, unidades)
    is_trazabilidad = (e["tipo_trabajo"] or "").upper() == "TRAZABILIDAD"
    prefijo_dm, prefijo_nombre, nums = "", "", []
    if is_trazabilidad:
        if not e["cliente_id"]:
            conn.close()
            return HTMLResponse("OT de trazabilidad sin cliente registrado", status_code=400)
        prefijo_dm, prefijo_nombre, nums = _reserve_numeros_cliente(cur, int(e["cliente_id"]), unidades)
    is_pg = bool(os.environ.get("DATABASE_URL"))
    inst_ids = []
    for i in range(unidades):
        dm_auto = f"{prefijo_dm}{str(nums[i]).zfill(5)}" if is_trazabilidad else ""
        nombre_trz_auto = _build_nombre_trazabilidad(prefijo_nombre, dm_auto) if is_trazabilidad else ""
        vals = [envio_id, (codigo_producto or "").strip(), (fabricante or "").strip(),
                (num_serie or "").strip(), (denominacion or "").strip(), (observaciones or "").strip(),
                (dm_auto or (codigo_datamatrix or "").strip()), (nombre_trz_auto or ""), (codigo_cliente or "").strip()]
        sql = """INSERT INTO instrumentos (envio_id, codigo_producto, fabricante, num_serie, denominacion, observaciones, codigo_datamatrix, nombre_trazabilidad, codigo_cliente, estado, creado_en)
                 VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'Pendiente', CURRENT_TIMESTAMP)"""
        if is_pg:
            cur.execute(sql + " RETURNING id", tuple(vals))
            row = cur.fetchone()
            inst_id = int(row["id"]) if row else None
        else:
            cur.execute(sql, tuple(vals))
            inst_id = cur.lastrowid
        inst_ids.append(inst_id)
    conn.commit()
    conn.close()
    if unidades > 1: return RedirectResponse(url=f"/envios/{envio_id}", status_code=303)
    target_id = inst_ids[0]
    if is_trazabilidad: return RedirectResponse(url=f"/instrumentos/{target_id}", status_code=303)
    return RedirectResponse(url=f"/instrumentos/{target_id}/editar", status_code=303)

@router.get("/instrumentos/{instrumento_id}/editar", response_class=HTMLResponse)
def instrumento_editar_form(request: Request, instrumento_id: int, user=Depends(require_roles("admin", "recepcion", "tecnico", "socio"))):
    templates = request.app.state.templates
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
    return templates.TemplateResponse("instrumento_nuevo.html", {
        "request": request, "user": user, "mode": "edit", "envio": dict(envio) if envio else None, "inst": dict(inst),
    })

@router.post("/instrumentos/{instrumento_id}/editar")
def instrumento_editar_guardar(
    instrumento_id: int,
    codigo_producto: str = Form(""),
    fabricante: str = Form(""),
    num_serie: str = Form(""),
    denominacion: str = Form(""),
    observaciones: str = Form(""),
    codigo_datamatrix: str = Form(""),
    codigo_cliente: str = Form(""),
    user=Depends(require_roles("admin", "recepcion", "tecnico")),
):
    conn = get_conn()
    cur = conn.cursor()
    if not (can_action(user, "instrumento_editar", cur) or can_action(user, "fotos_gestionar", cur)):
        conn.close()
        return RedirectResponse(url="/?err=perm", status_code=303)
    dm = (codigo_datamatrix or "").strip()
    cur.execute("""SELECT e.id AS envio_id, COALESCE(e.tipo_trabajo,'REPARACION') AS tipo_trabajo, e.cliente_id, c.prefijo_nombre
                   FROM instrumentos i JOIN envios e ON e.id = i.envio_id LEFT JOIN clientes c ON c.id = e.cliente_id
                   WHERE i.id=?""", (instrumento_id,))
    meta = cur.fetchone()
    if not meta:
        conn.close()
        return HTMLResponse("Instrumento no encontrado", status_code=404)
    nombre_trz = _build_nombre_trazabilidad((meta.get("prefijo_nombre") or ""), dm) if (meta["tipo_trabajo"] or "").upper() == "TRAZABILIDAD" else ""
    cur.execute("""UPDATE instrumentos SET codigo_producto=?, fabricante=?, num_serie=?, denominacion=?, observaciones=?, codigo_datamatrix=?, nombre_trazabilidad=?, codigo_cliente=?, actualizado_en=CURRENT_TIMESTAMP WHERE id=?""",
                ((codigo_producto or "").strip(), (fabricante or "").strip(), (num_serie or "").strip(), (denominacion or "").strip(), (observaciones or "").strip(), dm, nombre_trz, (codigo_cliente or "").strip(), instrumento_id))
    conn.commit()
    conn.close()
    return RedirectResponse(url=f"/instrumentos/{instrumento_id}/editar", status_code=303)

@router.get("/instrumentos/{instrumento_id}", response_class=HTMLResponse)
def instrumento_detalle(request: Request, instrumento_id: int, user=Depends(get_current_user)):
    templates = request.app.state.templates
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT i.*, (SELECT 1 FROM instrumento_checklist ic WHERE ic.instrumento_id=i.id LIMIT 1) as has_checklist FROM instrumentos i WHERE i.id=?", (instrumento_id,))
    inst = cur.fetchone()
    if not inst:
        conn.close()
        return HTMLResponse("Instrumento no encontrado", status_code=404)
    cur.execute("SELECT * FROM envios WHERE id=?", (inst["envio_id"],))
    envio = cur.fetchone()
    if _user_role(user) == "cliente":
        u_cli_id = (user.get("cliente_id") if isinstance(user, dict) else getattr(user, "cliente_id", None))
        if not envio or not u_cli_id or int(envio.get("cliente_id") or 0) != int(u_cli_id):
            conn.close()
            return HTMLResponse("Acceso denegado", status_code=403)
    cliente = _get_cliente(cur, int(envio["cliente_id"])) if envio and envio["cliente_id"] else None
    tipo_ot = str(envio.get("tipo_trabajo") or "REPARACION").strip().upper() if envio else "REPARACION"
    if tipo_ot == "TRAZABILIDAD": tipo_ot = "REPARACION"
    cur.execute("""SELECT ci.id AS item_id, ci.nombre, ci.orden, COALESCE(ic.hecho,0) AS hecho, ic.hecho_por, ic.hecho_en FROM checklist_items ci LEFT JOIN instrumento_checklist ic ON ic.item_id = ci.id AND ic.instrumento_id = ? WHERE COALESCE(ci.activo,1)=1 AND COALESCE(ci.tipo_trabajo,'REPARACION') = ? ORDER BY LOWER(ci.nombre) ASC""", (instrumento_id, tipo_ot))
    checklist = [dict(r) for r in cur.fetchall()]
    if not checklist and tipo_ot != "REPARACION":
        cur.execute("""SELECT ci.id AS item_id, ci.nombre, ci.orden, COALESCE(ic.hecho,0) AS hecho, ic.hecho_por, ic.hecho_en FROM checklist_items ci LEFT JOIN instrumento_checklist ic ON ic.item_id = ci.id AND ic.instrumento_id = ? WHERE COALESCE(ci.activo,1)=1 AND COALESCE(ci.tipo_trabajo,'REPARACION') = 'REPARACION' ORDER BY LOWER(ci.nombre) ASC""", (instrumento_id,))
        checklist = [dict(r) for r in cur.fetchall()]
    cur.execute("SELECT * FROM instrumento_informes WHERE instrumento_id=? ORDER BY id DESC LIMIT 1", (instrumento_id,))
    informe = cur.fetchone()
    dm, sn, historial = (inst["codigo_datamatrix"] or "").strip(), (inst["num_serie"] or "").strip(), []
    if dm or sn:
        clauses, p_hist = [], []
        if dm: clauses.append("i.codigo_datamatrix = ?"); p_hist.append(dm)
        if sn: clauses.append("i.num_serie = ?"); p_hist.append(sn)
        sql_hist = f"SELECT i.id, i.envio_id, e.ot_num, e.fecha, i.estado, i.creado_en FROM instrumentos i JOIN envios e ON e.id = i.envio_id WHERE ({ ' OR '.join(clauses) }) AND i.id != ?"
        p_hist.append(instrumento_id)
        if _user_role(user) == "cliente":
            u_cli_id = (user.get("cliente_id") if isinstance(user, dict) else getattr(user, "cliente_id", None))
            if u_cli_id: sql_hist += " AND e.cliente_id = ?"; p_hist.append(int(u_cli_id))
        sql_hist += " ORDER BY e.id DESC LIMIT 10"
        cur.execute(sql_hist, tuple(p_hist))
        historial = [dict(r) for r in cur.fetchall()]
    cur.execute("SELECT rc.*, COALESCE((SELECT ir.cantidad FROM instrumento_repuestos ir WHERE ir.instrumento_id=? AND ir.repuesto_id=rc.id), 0) as cantidad FROM repuestos_catalogo rc WHERE COALESCE(rc.activo,1)=1 ORDER BY rc.nombre", (instrumento_id,))
    repuestos_frecuentes = [dict(r) for r in cur.fetchall()]
    conn.close()
    return templates.TemplateResponse("instrumento_detalle.html", {
        "request": request, "user": user, "inst": dict(inst), "checklist": checklist,
        "repuestos_frecuentes": repuestos_frecuentes, "envio": dict(envio) if envio else None,
        "cliente": dict(cliente) if cliente else None, "informe": dict(informe) if informe else None, "historial": historial
    })

@router.get("/api/instrumentos/{instrumento_id}/checklist")
def api_get_checklist(instrumento_id: int, user=Depends(get_current_user)):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT e.tipo_trabajo FROM instrumentos i JOIN envios e ON e.id = i.envio_id WHERE i.id=?", (instrumento_id,))
    row = cur.fetchone()
    tipo_ot = (row["tipo_trabajo"] if row else "REPARACION") or "REPARACION"
    if tipo_ot == "TRAZABILIDAD": tipo_ot = "REPARACION"
    cur.execute("""SELECT ci.id AS item_id, ci.nombre, COALESCE(ic.hecho,0) AS hecho FROM checklist_items ci LEFT JOIN instrumento_checklist ic ON ic.item_id = ci.id AND ic.instrumento_id = ? WHERE ci.activo = 1 AND ci.tipo_trabajo = ? ORDER BY LOWER(ci.nombre) ASC""", (instrumento_id, tipo_ot))
    checklist = [dict(r) for r in cur.fetchall()]
    conn.close()
    return {"checklist": checklist}

@router.post("/api/instrumentos/{instrumento_id}/checklist")
async def api_save_checklist(instrumento_id: int, request: Request, user=Depends(require_roles("admin", "tecnico"))):
    data = await request.json()
    items_hechos = data.get("items", [])
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM instrumento_checklist WHERE instrumento_id=?", (instrumento_id,))
    for item_id in items_hechos:
        cur.execute("INSERT INTO instrumento_checklist (instrumento_id, item_id, hecho, hecho_por, hecho_en) VALUES (?, ?, 1, ?, ?)", (instrumento_id, item_id, user.get("username"), datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit()
    conn.close()
    return {"ok": True}

@router.post("/instrumentos/{instrumento_id}/check/{item_id}")
async def checklist_toggle(instrumento_id: int, item_id: int, request: Request, user=Depends(require_roles("admin", "tecnico"))):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT hecho FROM instrumento_checklist WHERE instrumento_id=? AND item_id=?", (instrumento_id, item_id))
    row = cur.fetchone()
    actual = int(row["hecho"]) if row else 0
    form = await request.form()
    raw = form.get("hecho")
    nuevo = 1 if (raw in ("1", "true", "on") if raw is not None else 1 - actual) else 0
    if row:
        cur.execute("UPDATE instrumento_checklist SET hecho=?, hecho_por=?, hecho_en=CURRENT_TIMESTAMP WHERE instrumento_id=? AND item_id=?", (nuevo, user.get("username"), instrumento_id, item_id))
    else:
        cur.execute("INSERT INTO instrumento_checklist (instrumento_id, item_id, hecho, hecho_por, hecho_en) VALUES (?, ?, ?, ?, CURRENT_TIMESTAMP)", (instrumento_id, item_id, nuevo, user.get("username")))
    conn.commit()
    conn.close()
    if "application/json" in (request.headers.get("accept") or "").lower(): return {"ok": True, "hecho": nuevo}
    return RedirectResponse(url=f"/instrumentos/{instrumento_id}", status_code=303)

@router.post("/instrumentos/{instrumento_id}/estado")
async def cambiar_estado(instrumento_id: int, request: Request, user=Depends(require_roles("admin", "tecnico"))):
    form = await request.form()
    estado = (form.get("estado") or "").strip()
    if estado not in ("Pendiente", "En proceso", "Reparado", "Baja"): return HTMLResponse("Estado inválido", status_code=400)
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT envio_id FROM instrumentos WHERE id=?", (instrumento_id,))
    envio_id = cur.fetchone()["envio_id"]
    cur.execute("UPDATE instrumentos SET estado=?, tecnico_reparacion=?, tecnico_reparacion_en=CURRENT_TIMESTAMP, repuesto_info=?, repuesto_precio=?, recomendada_sustitucion=? WHERE id=?",
                (estado, user.get("username"), form.get("repuesto_info"), float(form.get("repuesto_precio") or 0), 1 if form.get("recomendada_sustitucion") else 0, instrumento_id))
    conn.commit()
    conn.close()
    return RedirectResponse(url=f"/envios/{envio_id}" if envio_id else "/", status_code=303)

@router.post("/instrumentos/{instrumento_id}/foto_webcam/{slot}")
async def foto_webcam(instrumento_id: int, slot: int, request: Request, user=Depends(require_roles("admin", "recepcion", "tecnico"))):
    if slot not in range(1, 7): return JSONResponse({"ok": False, "error": "slot inválido"}, status_code=400)
    data = await request.json()
    image = (data.get("image") or "").split(",", 1)[-1]
    raw = base64.b64decode(image)
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(f"SELECT codigo_datamatrix, nombre_trazabilidad, foto_entrada_{slot} AS old FROM instrumentos WHERE id=?", (instrumento_id,))
    row = dict(cur.fetchone())
    tag = re.sub(r"[^A-Za-z0-9_-]+", "_", (row.get("nombre_trazabilidad") or row.get("codigo_datamatrix") or "SIN_CODIGO"))[:40]
    filename = f"inst_{instrumento_id}_{tag}_f{slot}_{uuid.uuid4().hex[:8]}.jpg"
    path_fs = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "static", "fotos", filename)
    with open(path_fs, "wb") as f: f.write(raw)
    public_path = f"/static/fotos/{filename}"
    if row["old"]:
        try: os.remove(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), row["old"].lstrip("/")))
        except: pass
    cur.execute(f"UPDATE instrumentos SET foto_entrada_{slot}=? WHERE id=?", (public_path, instrumento_id))
    conn.commit()
    conn.close()
    return {"ok": True, "path": public_path}

@router.post("/instrumentos/{instrumento_id}/repuesto/{repuesto_id}/adjust")
def adjust_instrumento_repuesto(instrumento_id: int, repuesto_id: int, action: str = "add", user=Depends(require_roles("admin", "tecnico"))):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT cantidad FROM instrumento_repuestos WHERE instrumento_id=? AND repuesto_id=?", (instrumento_id, repuesto_id))
    row = cur.fetchone()
    if action == "add":
        if row: cur.execute("UPDATE instrumento_repuestos SET cantidad = cantidad + 1 WHERE instrumento_id=? AND repuesto_id=?", (instrumento_id, repuesto_id))
        else:
            cur.execute("SELECT precio FROM repuestos_catalogo WHERE id=?", (repuesto_id,))
            rep = cur.fetchone()
            cur.execute("INSERT INTO instrumento_repuestos (instrumento_id, repuesto_id, precio_aplicado, cantidad) VALUES (?, ?, ?, 1)", (instrumento_id, repuesto_id, rep["precio"] if rep else 0))
    elif action == "sub" and row:
        if row["cantidad"] > 1: cur.execute("UPDATE instrumento_repuestos SET cantidad = cantidad - 1 WHERE instrumento_id=? AND repuesto_id=?", (instrumento_id, repuesto_id))
        else: cur.execute("DELETE FROM instrumento_repuestos WHERE instrumento_id=? AND repuesto_id=?", (instrumento_id, repuesto_id))
    is_pg = os.environ.get("DATABASE_URL") is not None
    agg = "string_agg(CASE WHEN cantidad > 1 THEN nombre || ' x' || cantidad ELSE nombre END, ', ')" if is_pg else "GROUP_CONCAT(CASE WHEN cantidad > 1 THEN nombre || ' x' || cantidad ELSE nombre END, ', ')"
    cur.execute(f"SELECT SUM(precio_aplicado * cantidad) as total, {agg} as nombres FROM instrumento_repuestos ir JOIN repuestos_catalogo rc ON ir.repuesto_id = rc.id WHERE ir.instrumento_id = ?", (instrumento_id,))
    res = cur.fetchone()
    cur.execute("UPDATE instrumentos SET repuesto_precio=?, repuesto_info=? WHERE id=?", (res["total"] or 0, res["nombres"] or "", instrumento_id))
    conn.commit()
    conn.close()
    return {"ok": True, "total_precio": res["total"] or 0, "total_info": res["nombres"] or ""}

@router.post("/instrumentos/{instrumento_id}/borrar")
def borrar_instrumento(instrumento_id: int, user=Depends(require_roles("admin", "recepcion"))):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT envio_id FROM instrumentos WHERE id=?", (instrumento_id,))
    row = cur.fetchone()
    envio_id = row["envio_id"] if row else None
    cur.execute("DELETE FROM instrumentos WHERE id=?", (instrumento_id,))
    conn.commit()
    conn.close()
    return RedirectResponse(url=f"/envios/{envio_id}" if envio_id else "/", status_code=303)
