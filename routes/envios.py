import os
import io
import re
import time
from datetime import datetime
from typing import Optional, List, Dict, Any

from fastapi import APIRouter, Request, Form, File, UploadFile, Depends, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse
from fastapi.templating import Jinja2Templates

from db import get_conn, get_table_columns
from security import get_current_user, require_roles
from shared import (
    _user_role, _envios_has_column, can_action, 
    _get_cliente, _user_id
)
from utils import _clean_trz, _build_nombre_trazabilidad
from excel_import import leer_excel_envio

# Para etiquetas PDF
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.graphics.barcode import qr

router = APIRouter()

# Debería configurarse en app.py y pasarse aquí o usar app.state
templates = Jinja2Templates(directory="templates")

# Configuración de rutas de archivos (esto debería venir de una config central)
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
FOTOS_DIR = os.path.join(BASE_DIR, "static", "fotos")

@router.post("/envios/{envio_id}/aviso_finalizacion")
def envio_aviso_finalizacion(envio_id: int, user=Depends(require_roles("admin", "recepcion"))):
    """Envía un aviso por email al cliente si el parte está terminado."""
    conn = get_conn()
    cur = conn.cursor()
    
    cur.execute("""
        SELECT e.*, c.email as cliente_email, c.nombre as cliente_nombre
        FROM envios e
        LEFT JOIN clientes c ON e.cliente_id = c.id
        WHERE e.id=?
    """, (envio_id,))
    envio = cur.fetchone()
    
    if not envio:
        conn.close()
        return {"ok": False, "error": "Envío no encontrado"}
        
    if not envio["cliente_email"]:
        conn.close()
        return {"ok": False, "error": "El cliente no tiene un email configurado para notificaciones"}

    cur.execute("SELECT COUNT(*) as total FROM instrumentos WHERE envio_id=?", (envio_id,))
    total = cur.fetchone()["total"]
    
    cur.execute("SELECT COUNT(*) as done FROM instrumentos WHERE envio_id=? AND estado IN ('Reparado', 'Baja')", (envio_id,))
    done = cur.fetchone()["done"]
    
    if total == 0 or done < total:
        conn.close()
        return {"ok": False, "error": "El parte aún no está totalmente terminado (faltan piezas por revisar)"}

    from mail_utils import send_finish_notification
    success, msg = send_finish_notification(
        envio["cliente_email"], 
        envio["cliente_nombre"] or envio["cliente"], 
        envio["ot_num"], 
        total
    )
    
    if success:
        try:
            cur.execute("UPDATE envios SET aviso_enviado=1 WHERE id=?", (envio_id,))
            conn.commit()
        except:
            pass
            
    conn.close()
    return {"ok": success, "message": msg}

@router.get("/envios/{envio_id}", response_class=HTMLResponse)
def ver_envio(request: Request, envio_id: int, user=Depends(get_current_user)):
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("SELECT * FROM envios WHERE id=?", (envio_id,))
    envio = cur.fetchone()
    if not envio:
        conn.close()
        return HTMLResponse("Envío no encontrado", status_code=404)

    if _user_role(user) == "cliente":
        u_cli_id = (user.get("cliente_id") if isinstance(user, dict) else getattr(user, "cliente_id", None))
        e_cli_id = envio.get("cliente_id")
        if not u_cli_id or int(e_cli_id or 0) != int(u_cli_id):
            conn.close()
            return HTMLResponse("Acceso denegado: este envío no pertenece a su centro", status_code=403)
        
        if not bool(envio["aceptado"]):
            conn.close()
            return HTMLResponse("<h1>Acceso restringido</h1><p>Este parte aún no ha sido validado/aceptado. Por favor, contacte con recepción para más información.</p>", status_code=403)

    cur.execute("""
        SELECT i.*, 
               (SELECT 1 FROM instrumento_informes ii WHERE ii.instrumento_id = i.id LIMIT 1) as has_archived,
               (SELECT 1 FROM instrumento_qc_optica qc WHERE qc.instrumento_id = i.id LIMIT 1) as has_qc_data,
               (SELECT 1 FROM instrumento_checklist ic WHERE ic.instrumento_id = i.id LIMIT 1) as has_checklist
        FROM instrumentos i 
        WHERE i.envio_id=? 
        ORDER BY i.id DESC
    """, (envio_id,))
    instrumentos = [dict(r) for r in cur.fetchall()]
    
    total_inst = len(instrumentos)
    done_inst = sum(1 for i in instrumentos if i["estado"] in ["Reparado", "Baja"])
    is_finished = (total_inst > 0 and done_inst == total_inst)

    for r in instrumentos:
        if r.get("nombre_trazabilidad"):
            r["nombre_trazabilidad"] = _clean_trz(r["nombre_trazabilidad"])
        r["has_informe"] = bool(r.get("has_archived") or r.get("has_qc_data"))

    envio_dict = dict(envio)
    prefijo = ""
    for k in ("prefijo_nombre", "prefijo_trazabilidad", "prefijo_traz", "prefijo_etiqueta", "prefijo"):
        v = envio_dict.get(k)
        if v:
            prefijo = str(v).strip()
            break

    for r in instrumentos:
        nt = (r.get("nombre_trazabilidad") or "").strip()
        if nt: nt = nt.strip().strip("'").strip('"')
        if nt and prefijo and re.fullmatch(r"\d{5}", nt):
            nt = f"{prefijo}{nt}"
        if (not nt) and prefijo:
            dm = (r.get("codigo_datamatrix") or "").strip()
            if dm:
                try:
                    nt = _build_nombre_trazabilidad(prefijo, dm)
                except Exception:
                    nt = ""
        r["nombre_trazabilidad"] = _clean_trz(nt)

    conn.close()

    if _user_role(user) == "tecnico":
        return templates.TemplateResponse(
            "tecnico_parte.html",
            {"request": request, "user": user, "envio": dict(envio), "instrumentos": instrumentos, "is_finished": is_finished},
        )

    return templates.TemplateResponse(
        "envio_detalle.html",
        {"request": request, "user": user, "envio": dict(envio), "instrumentos": instrumentos, "is_finished": is_finished},
    )

def _build_etiqueta_pdf(ot_num: str, cliente: str, fecha: str, n_instrumentos: int, referencia: str = "",
                        fabricante: str = "", modelo: str = "", serie: str = "") -> bytes:
    w, h = 62 * mm, 29 * mm
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(w, h))
    x0 = 7 * mm
    y_top = h - 2 * mm
    c.setFont("Helvetica-Bold", 10)
    c.drawString(x0, y_top - 3.5 * mm, f"OT: {ot_num}")
    c.setFont("Helvetica", 7.5)
    
    def truncate(s, limit=22):
        s = (s or "").strip()
        return s[:limit-3] + "…" if len(s) > limit else s

    cli = truncate(cliente, 22)
    ref = truncate(referencia, 22)
    fab = truncate(fabricante, 22)
    mod = truncate(modelo, 22)
    sn = truncate(serie, 22)

    curr_y = y_top - 7.5 * mm
    step = 3.2 * mm
    c.drawString(x0, curr_y, f"Cli: {cli}")
    curr_y -= step
    if fab or mod:
        c.drawString(x0, curr_y, f"Art: {fab} {mod}")
        curr_y -= step
    else:
        c.drawString(x0, curr_y, f"Ref: {ref}")
        curr_y -= step
    if sn:
        c.drawString(x0, curr_y, f"S/N: {sn}")
        curr_y -= step
    else:
        c.drawString(x0, curr_y, f"Fecha: {fecha}")
        curr_y -= step
    c.drawString(x0, curr_y, f"Inst: {n_instrumentos} | {fecha if sn else ''}")
    payload = str(ot_num)
    from reportlab.graphics.shapes import Drawing
    from reportlab.graphics import renderPDF
    qr_code = qr.QrCodeWidget(payload)
    bounds = qr_code.getBounds()
    qr_w = bounds[2] - bounds[0]
    qr_h = bounds[3] - bounds[1]
    size = 17 * mm
    d = Drawing(size, size, transform=[size/qr_w, 0, 0, size/qr_h, 0, 0])
    d.add(qr_code)
    renderPDF.draw(d, c, w - size - 3 * mm, 3.5 * mm)
    c.showPage()
    c.save()
    return buf.getvalue()

@router.get("/envios/{envio_id}/etiqueta.pdf")
def etiqueta_envio(envio_id: int, user=Depends(get_current_user)):
    conn = get_conn()
    cur = conn.cursor()
    from db import get_table_columns
    cols = get_table_columns(cur, "envios")
    if "tipo_trabajo" in cols:
        cur.execute("SELECT id, ot_num, cliente, fecha, cliente_id, nombre_archivo, tipo_trabajo FROM envios WHERE id=?", (envio_id,))
    else:
        cur.execute("SELECT id, ot_num, cliente, fecha, cliente_id, nombre_archivo FROM envios WHERE id=?", (envio_id,))
    e = cur.fetchone()
    if not e:
        conn.close()
        return HTMLResponse("OT no encontrada", status_code=404)
    if _user_role(user) == "cliente":
        u_cli_id = (user.get("cliente_id") if isinstance(user, dict) else getattr(user, "cliente_id", None))
        if not u_cli_id or int(e["cliente_id"] or 0) != int(u_cli_id):
            conn.close()
            return HTMLResponse("No tienes permiso para ver esta etiqueta", status_code=403)
    cur.execute("SELECT COUNT(*) AS n FROM instrumentos WHERE envio_id=?", (envio_id,))
    n_inst = int(cur.fetchone()["n"] or 0)
    fabricante, modelo, serie = "", "", ""
    tipo = str(e["tipo_trabajo"]).upper() if "tipo_trabajo" in e.keys() else "REPARACION"
    if tipo == "OPTICA_RIGIDA" and n_inst > 0:
        cur.execute("SELECT fabricante, codigo_producto, num_serie FROM instrumentos WHERE envio_id=? LIMIT 1", (envio_id,))
        inst_row = cur.fetchone()
        if inst_row:
            fabricante = inst_row["fabricante"] or ""
            modelo = inst_row["codigo_producto"] or ""
            serie = inst_row["num_serie"] or ""
    conn.close()
    fecha_raw = (e["fecha"] or "").strip()
    if fecha_raw:
        try:
            dt = datetime.fromisoformat(fecha_raw)
            fecha = dt.strftime("%d/%m/%Y")
        except ValueError:
            fecha = fecha_raw
    else:
        fecha = datetime.now().strftime("%d/%m/%Y")
    pdf = _build_etiqueta_pdf(str(e["ot_num"]), str(e["cliente"] or ""), fecha, n_inst,
                            referencia=str(e["nombre_archivo"] or ""), fabricante=fabricante,
                            modelo=modelo, serie=serie)
    filename = f"OT_{e['ot_num']}_etiqueta.pdf"
    return StreamingResponse(io.BytesIO(pdf), media_type="application/pdf",
                            headers={"Content-Disposition": f"inline; filename={filename}"})

@router.post("/envios/{envio_id}/importar")
async def envio_importar_excel(envio_id: int, file: UploadFile = File(...), user=Depends(require_roles("admin", "recepcion"))):
    path = os.path.join(UPLOAD_DIR, file.filename)
    with open(path, "wb") as f:
        f.write(await file.read())
    try:
        _, _, df = leer_excel_envio(path)
    except Exception as e:
        return RedirectResponse(url=f"/envios/{envio_id}?err=excel&msg={str(e)}", status_code=303)
    conn = get_conn()
    cur = conn.cursor()
    rows = []
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for _, r in df.iterrows():
        rows.append((envio_id, r.get("codigo_producto"), r.get("fabricante"), r.get("num_serie"),
                    r.get("denominacion"), r.get("observaciones"), str(r.get("codigo_datamatrix") or "").strip(),
                    str(r.get("codigo_cliente") or "").strip(), "", "Pendiente", now_str))
    cur.executemany("""
        INSERT INTO instrumentos
        (envio_id, codigo_producto, fabricante, num_serie, denominacion, observaciones, codigo_datamatrix, codigo_cliente, nombre_trazabilidad, estado, creado_en)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, rows)
    return RedirectResponse(url=f"/envios/{envio_id}", status_code=303)

@router.get("/importar", response_class=HTMLResponse)
def importar_form(request: Request, user=Depends(require_roles("admin", "recepcion"))):
    conn = get_conn()
    cur = conn.cursor()
    from shared import _list_clientes
    clientes = _list_clientes(cur)
    conn.close()
    return templates.TemplateResponse("importar.html", {"request": request, "user": user, "clientes": clientes})

@router.post("/importar")
async def importar_excel(
    tipo_trabajo: str = Form("REPARACION"),
    referencia: str = Form(""),
    cliente_id: str = Form(""),
    cliente: str = Form("", alias="cliente"),
    fecha: str = Form("", alias="fecha"),
    observaciones: str = Form(""),
    file: UploadFile = File(...),
    user=Depends(require_roles("admin", "recepcion")),
):
    path = os.path.join(UPLOAD_DIR, file.filename)
    with open(path, "wb") as f:
        f.write(await file.read())
    try:
        cliente_auto, fecha_auto, df = leer_excel_envio(path)
    except Exception as e:
        return RedirectResponse(url=f"/importar?err=excel&msg={str(e)}", status_code=303)
    conn = get_conn()
    cur = conn.cursor()
    final_cliente, target_cliente_id = "", None
    if cliente_id:
        cur.execute("SELECT nombre FROM clientes WHERE id=?", (cliente_id,))
        row = cur.fetchone()
        if row:
            final_cliente = row["nombre"]
            target_cliente_id = int(cliente_id)
    elif cliente:
        final_cliente = cliente.strip()
    else:
        final_cliente = str(cliente_auto).strip()
    if not final_cliente:
        conn.close()
        return RedirectResponse(url="/importar?err=cliente", status_code=303)
    final_fecha = fecha if fecha else fecha_auto
    final_ref = referencia.strip() if referencia.strip() else file.filename
    from dashboard import _next_ot_num
    ot_num = _next_ot_num(cur)
    tipo_trabajo = (tipo_trabajo or "REPARACION").strip().upper()
    if tipo_trabajo not in ("REPARACION", "TRAZABILIDAD", "OPTICA_RIGIDA"):
        tipo_trabajo = "REPARACION"
    cols, vals = ["ot_num", "nombre_archivo", "cliente", "fecha"], [ot_num, final_ref, final_cliente, final_fecha]
    if _envios_has_column(cur, "tipo_trabajo"):
        cols.append("tipo_trabajo"); vals.append(tipo_trabajo)
    if _envios_has_column(cur, "cliente_id") and target_cliente_id:
        cols.append("cliente_id"); vals.append(target_cliente_id)
    if _envios_has_column(cur, "observaciones"):
        cols.append("observaciones"); vals.append(observaciones)
    sql = f"INSERT INTO envios ({', '.join(cols)}) VALUES ({', '.join(['?']*len(vals))})"
    is_pg = os.environ.get("DATABASE_URL") is not None
    if is_pg:
        cur.execute(sql + " RETURNING id", tuple(vals))
        row = cur.fetchone()
        envio_id = int(row["id"]) if row else 0
    else:
        cur.execute(sql, tuple(vals))
        envio_id = cur.lastrowid
    rows, now_str = [], datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for _, r in df.iterrows():
        rows.append((envio_id, r.get("codigo_producto"), r.get("fabricante"), r.get("num_serie"),
                    r.get("denominacion"), r.get("observaciones"), str(r.get("codigo_datamatrix") or "").strip(),
                    str(r.get("codigo_cliente") or "").strip(), "", "Pendiente", now_str))
    cur.executemany("""
        INSERT INTO instrumentos (envio_id, codigo_producto, fabricante, num_serie, denominacion, observaciones, codigo_datamatrix, codigo_cliente, nombre_trazabilidad, estado, creado_en)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, rows)
    return RedirectResponse(url=f"/envios/{envio_id}", status_code=303)

@router.get("/envios/nuevo", response_class=HTMLResponse)
def nuevo_envio_form(request: Request, user=Depends(require_roles("admin", "recepcion", "socio"))):
    conn = get_conn(); cur = conn.cursor()
    from shared import _list_clientes
    clientes = _list_clientes(cur)
    conn.close()
    return templates.TemplateResponse("envio_nuevo.html", {"request": request, "user": user, "clientes": clientes})

@router.post("/envios/nuevo")
def nuevo_envio_crear(referencia: str = Form(""), cliente_id: str = Form(""), cliente: str = Form(""), tipo_trabajo: str = Form("REPARACION"), fecha: str = Form(""), observaciones: str = Form(""), user=Depends(require_roles("admin", "recepcion"))):
    fecha = (fecha or "").strip()
    if not fecha: return RedirectResponse(url="/envios/nuevo?err=fecha", status_code=303)
    conn = get_conn(); cur = conn.cursor()
    if not can_action(user, "envio_crear", cur):
        conn.close(); return RedirectResponse(url="/?err=perm", status_code=303)
    from dashboard import _next_ot_num
    ot_num = _next_ot_num(cur)
    if cliente_id:
        cur.execute("SELECT nombre FROM clientes WHERE id=?", (cliente_id,))
        row = cur.fetchone()
        if row: cliente = row["nombre"]
    cols, vals = ["ot_num", "nombre_archivo", "cliente", "fecha", "observaciones"], [ot_num, referencia, cliente, fecha, observaciones]
    if _envios_has_column(cur, "tipo_trabajo"): cols.append("tipo_trabajo"); vals.append(tipo_trabajo.strip().upper())
    if _envios_has_column(cur, "cliente_id") and cliente_id: cols.append("cliente_id"); vals.append(int(cliente_id))
    cur.execute(f"INSERT INTO envios ({', '.join(cols)}) VALUES ({', '.join(['?']*len(vals))})", tuple(vals))
    conn.commit(); conn.close()
    return RedirectResponse(url="/", status_code=303)

@router.post("/envios/{envio_id}/borrar")
def borrar_envio(request: Request, envio_id: int, confirm_ot: str = Form(""), user=Depends(require_roles("admin", "recepcion"))):
    conn_perm = get_conn()
    if not can_action(user, "envio_borrar", conn_perm.cursor()):
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
    if (confirm_ot or "").strip() != ot_num:
        conn.close()
        return RedirectResponse(url="/?err=confirm_ot", status_code=303)
    cur.execute("SELECT COUNT(*) AS n FROM instrumentos WHERE envio_id=?", (envio_id,))
    n_inst = int(cur.fetchone()["n"] or 0)
    cur.execute("SELECT COUNT(*) AS n FROM instrumentos WHERE envio_id=? AND estado IN ('Pendiente','En proceso')", (envio_id,))
    n_pend = int(cur.fetchone()["n"] or 0)
    if n_inst > 0 and n_pend == 0:
        conn.close()
        return RedirectResponse(url="/?err=cerrado", status_code=303)
    cur.execute("DELETE FROM instrumentos WHERE envio_id=?", (envio_id,))
    cur.execute("DELETE FROM envios WHERE id=?", (envio_id,))
    conn.commit()
    conn.close()
    return RedirectResponse(url="/?ok=borrado", status_code=303)
