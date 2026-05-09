import os
from fastapi import APIRouter, Request, Depends
from fastapi.responses import HTMLResponse
from db import get_conn
from security import get_current_user
from shared import (
    _user_role, _envios_has_column, _list_clientes, 
    _users_schema, _select_users_sql, _get_user_permissions_map, ACTIONS
)

router = APIRouter()

@router.get("/", response_class=HTMLResponse)
def dashboard(request: Request, user=Depends(get_current_user)):
    templates = request.app.state.templates
    
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

    # --- 0. Notificaciones de Recogida ---
    n_peticiones_pendientes = 0
    peticiones_recogida = []
    if _user_role(user) in ["admin", "recepcion"]:
        cur.execute("""
            SELECT pr.*, c.nombre as cliente_nombre 
            FROM peticiones_recogida pr
            JOIN clientes c ON pr.cliente_id = c.id
            WHERE pr.estado = 'Pendiente'
            ORDER BY pr.creado_en DESC
        """)
        peticiones_recogida = [dict(r) for r in cur.fetchall()]
        n_peticiones_pendientes = len(peticiones_recogida)

    # --- 0b. Consultas Técnicas (NUEVO Chat) ---
    consultas_list = []
    n_consultas_pendientes = 0
    n_consultas_activas = 0
    if _user_role(user) in ["admin", "recepcion", "tecnico"]:
        cur.execute("""
            SELECT c.*, cl.nombre as cliente_nombre 
            FROM consultas c
            JOIN clientes cl ON c.cliente_id = cl.id
            WHERE c.estado != 'Cerrada'
            ORDER BY c.actualizado_en DESC
        """)
        consultas_list = [dict(r) for r in cur.fetchall()]
        n_consultas_pendientes = sum(1 for c in consultas_list if c["estado"] == "Abierta")
        n_consultas_activas = len(consultas_list)
    elif _user_role(user) == "cliente" and user.get("cliente_id"):
        cur.execute("""
            SELECT * FROM consultas 
            WHERE cliente_id = ? AND estado != 'Cerrada'
            ORDER BY actualizado_en DESC
        """, (int(user.get("cliente_id") or 0),))
        consultas_list = [dict(r) for r in cur.fetchall()]
        n_consultas_pendientes = sum(1 for c in consultas_list if c["estado"] == "Respondida")
        n_consultas_activas = len(consultas_list)

    if _user_role(user) in ["admin", "recepcion"]:
        cur.execute("DELETE FROM instrumentos WHERE envio_id IS NULL OR envio_id = 0 OR envio_id NOT IN (SELECT id FROM envios)")
        conn.commit()

    # --- 1. KPIs Globales (Instrumentos) ---
    kpi_where = ""
    kpi_params = []
    if _user_role(user) == "cliente" and user.get("cliente_id"):
        kpi_where = "WHERE e.cliente_id = ?"
        kpi_params = [int(user["cliente_id"])]

    cur.execute(f"""
        SELECT 
            COUNT(i.id) as total,
            SUM(CASE WHEN i.estado='Pendiente' THEN 1 ELSE 0 END) as pendientes,
            SUM(CASE WHEN i.estado='En proceso' THEN 1 ELSE 0 END) as en_proceso,
            SUM(CASE WHEN i.estado='Reparado' THEN 1 ELSE 0 END) as reparados,
            SUM(CASE WHEN i.estado='Baja' THEN 1 ELSE 0 END) as baja
        FROM instrumentos i
        JOIN envios e ON e.id = i.envio_id
        {kpi_where}
    """, tuple(kpi_params))
    
    row_kpi = cur.fetchone()
    kpis = {"total": 0, "pendientes": 0, "en_proceso": 0, "reparado": 0, "baja": 0}
    if row_kpi:
        kpis = {
            "total": row_kpi["total"] or 0,
            "pendientes": row_kpi["pendientes"] or 0,
            "en_proceso": row_kpi["en_proceso"] or 0,
            "reparado": row_kpi["reparados"] or 0,
            "baja": row_kpi["baja"] or 0
        }

    # --- 2. KPIs de Partes (OTs) ---
    total_partes = 0
    abiertos = 0
    cerrados = 0
    opticas_total = 0
    opticas_cerrados = 0

    sql_parts = f"""
        SELECT 
            e.id, 
            COALESCE(e.tipo_trabajo, 'REPARACION') as tipo_trabajo,
            COUNT(i.id) as total_inst,
            SUM(CASE WHEN i.estado IN ('Reparado','Baja') THEN 1 ELSE 0 END) as n_terminados,
            SUM(CASE WHEN COALESCE(i.grabado,0)=1 THEN 1 ELSE 0 END) as n_grabados
        FROM envios e
        LEFT JOIN instrumentos i ON i.envio_id = e.id
        {kpi_where}
        GROUP BY e.id
    """
    cur.execute(sql_parts, tuple(kpi_params))
    parts_rows = cur.fetchall()
    total_partes = len(parts_rows)
    for r in parts_rows:
        t_inst = r["total_inst"] or 0
        tipo = r["tipo_trabajo"]
        is_cerrado = False
        if t_inst == 0: pass
        elif tipo == "TRAZABILIDAD": is_cerrado = (r["n_grabados"] == t_inst)
        else: is_cerrado = (r["n_terminados"] == t_inst)
        if is_cerrado: cerrados += 1
        else: abiertos += 1
        if tipo == "OPTICA_RIGIDA":
            opticas_total += 1
            if is_cerrado: opticas_cerrados += 1

    kpis_partes = {
        "total": total_partes, "abiertos": abiertos, "cerrados": cerrados,
        "opticas_total": opticas_total, "opticas_cerrados": opticas_cerrados
    }

    # --- Buscador ---
    where_clauses = []
    params_q: list = []
    if _user_role(user) == "cliente" and user.get("cliente_id"):
        where_clauses.append("e.cliente_id = ?")
        params_q.append(int(user["cliente_id"]))
    elif _user_role(user) == "cliente":
        where_clauses.append("1=0")

    if q:
        is_pg = os.environ.get("DATABASE_URL") is not None
        if is_pg:
            where_clauses.append("(e.ot_num ILIKE ? OR e.cliente ILIKE ? OR i.codigo_datamatrix ILIKE ? OR i.num_serie ILIKE ? OR i.codigo_producto ILIKE ? OR EXISTS (SELECT 1 FROM peticiones_recogida pr WHERE pr.num_peticion ILIKE ? AND pr.cliente_id=e.cliente_id) OR EXISTS (SELECT 1 FROM clientes c WHERE c.id=e.cliente_id AND CAST(c.numero_cliente AS TEXT) ILIKE ?))")
        else:
            where_clauses.append("(e.ot_num LIKE ? OR e.cliente LIKE ? OR i.codigo_datamatrix LIKE ? OR i.num_serie LIKE ? OR i.codigo_producto LIKE ? OR EXISTS (SELECT 1 FROM peticiones_recogida pr WHERE pr.num_peticion LIKE ? AND pr.cliente_id=e.cliente_id) OR EXISTS (SELECT 1 FROM clientes c WHERE c.id=e.cliente_id AND CAST(c.numero_cliente AS TEXT) LIKE ?))")
        params_q.extend([f"%{q}%"] * 7)

    where_q = "WHERE " + " AND ".join(where_clauses) if where_clauses else ""
    cur.execute(f"""
        SELECT 
            e.*,
            COUNT(i.id) AS n_instrumentos,
            SUM(CASE WHEN 
                (i.foto_entrada_1 IS NOT NULL AND i.foto_entrada_1 != '') OR 
                (i.foto_entrada_2 IS NOT NULL AND i.foto_entrada_2 != '') OR 
                (i.foto_entrada_3 IS NOT NULL AND i.foto_entrada_3 != '') OR 
                (i.foto_entrada_4 IS NOT NULL AND i.foto_entrada_4 != '') OR 
                (i.foto_entrada_5 IS NOT NULL AND i.foto_entrada_5 != '') OR 
                (i.foto_entrada_6 IS NOT NULL AND i.foto_entrada_6 != '') THEN 1 ELSE 0 END) AS n_con_alguna_foto,
            SUM(CASE WHEN COALESCE(i.grabado,0)=1 THEN 1 ELSE 0 END) AS n_grabados,
            SUM(CASE WHEN i.estado = 'En proceso' THEN 1 ELSE 0 END) AS n_en_proceso,
            SUM(CASE WHEN COALESCE(i.repuesto_precio, 0) > 0 THEN 1 ELSE 0 END) AS n_con_repuesto,
            SUM(CASE 
                WHEN COALESCE(e.tipo_trabajo,'REPARACION') = 'TRAZABILIDAD' THEN (CASE WHEN COALESCE(i.grabado,0)=1 THEN 1 ELSE 0 END)
                ELSE (CASE WHEN COALESCE(i.estado,'') IN ('Reparado','Baja') THEN 1 ELSE 0 END)
            END) AS n_done
        FROM envios e
        LEFT JOIN instrumentos i ON i.envio_id = e.id
        {where_q}
        GROUP BY e.id
        ORDER BY e.id DESC
        LIMIT 200
    """, tuple(params_q))

    envios = []
    for r in cur.fetchall():
        d = dict(r)
        total = int(d.get("n_instrumentos") or 0)
        done = int(d.get("n_done") or 0)
        d["is_closed"] = (total > 0 and done == total)
        d["n_pendientes"] = max(total - done, 0)
        d["color"] = "green" if d["is_closed"] else "red"
        n_con_alguna_foto = int(d.get("n_con_alguna_foto") or 0)
        if total == 0 or n_con_alguna_foto == 0: d["foto_dot"] = "red"
        elif n_con_alguna_foto == total: d["foto_dot"] = "green"
        else: d["foto_dot"] = "yellow"
        envios.append(d)

    envios.sort(key=lambda x: (1 if x.get('is_closed') else 0, -int(x.get('id') or 0)))

    open_users_modal = (request.query_params.get("users") == "1")
    users_list = []
    perms_by_user = {}
    clientes_list_global = []

    if _user_role(user) == "admin":
        schema = _users_schema(cur)
        cur.execute(_select_users_sql(schema))
        users_list = [dict(r) for r in cur.fetchall()]
        for u in users_list:
            perms_by_user[int(u["id"])] = _get_user_permissions_map(cur, int(u["id"]))
        clientes_list_global = _list_clientes(cur)

    # RECOGIDAS SEARCH
    found_recogidas = []
    if q:
        rec_where = ["(num_peticion LIKE ? OR observaciones LIKE ? OR contacto LIKE ?)"]
        rec_params = [f"%{q}%"] * 3
        if _user_role(user) == "cliente":
            rec_where.append("cliente_id = ?")
            rec_params.append(int(user.get("cliente_id") or 0))
        cur.execute(f"""
            SELECT pr.*, c.nombre as cliente_nombre 
            FROM peticiones_recogida pr
            JOIN clientes c ON pr.cliente_id = c.id
            WHERE {" AND ".join(rec_where)}
            ORDER BY pr.creado_en DESC
        """, tuple(rec_params))
        found_recogidas = [dict(r) for r in cur.fetchall()]

    conn.close()
    return templates.TemplateResponse("dashboard.html", {
        "request": request, "user": user, "kpis": kpis, "kpis_partes": kpis_partes,
        "envios": envios, "q": q, "open_users_modal": open_users_modal,
        "users_list": users_list, "actions": ACTIONS, "perms_by_user": perms_by_user,
        "db_type": db_type, "clientes_list_global": clientes_list_global,
        "n_peticiones_pendientes": n_peticiones_pendientes, "peticiones_recogida": peticiones_recogida,
        "found_recogidas": found_recogidas, "consultas_list": consultas_list,
        "n_consultas_pendientes": n_consultas_pendientes, "n_consultas_activas": n_consultas_activas or 0,
    })
