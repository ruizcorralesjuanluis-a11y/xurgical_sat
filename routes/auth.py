from fastapi import APIRouter, Request, Depends, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse
from db import get_conn
from auth_utils import verify_password, sign_session
from security import get_current_user
from mail_utils import send_credentials_request
from utils import format_fecha # If needed elsewhere, but mostly for templates

router = APIRouter()

@router.get("/login", response_class=HTMLResponse)
def login_form(request: Request):
    templates = request.app.state.templates
    return templates.TemplateResponse("login.html", {"request": request, "error": None})

@router.post("/solicitar_acceso")
async def handle_solicitud_acceso(request: Request):
    form = await request.form()
    nombre = (form.get("nombre") or "").strip()
    centro = (form.get("centro") or "").strip()
    telefono = (form.get("telefono") or "").strip()
    email = (form.get("email") or "").strip()

    if not all([nombre, centro, email]):
        return RedirectResponse(url="/login?err=missing_fields", status_code=303)

    success, msg = send_credentials_request(nombre, centro, telefono, email)
    if success:
        return RedirectResponse(url="/login?msg=solicitud_enviada", status_code=303)
    else:
        return RedirectResponse(url=f"/login?err={msg}", status_code=303)

@router.post("/login")
async def login(request: Request):
    form = await request.form()
    username = (form.get("username") or "").strip()
    password = form.get("password") or ""

    conn = get_conn()
    cur = conn.cursor()

    # Need schema check or direct query? 
    # For now, keeping it consistent with app.py's inline check
    from db import get_table_columns
    cols = get_table_columns(cur, "users")
    colset = set(cols)
    pw_col = "password_hash" if "password_hash" in colset else ("password" if "password" in colset else None)
    has_is_active = "is_active" in colset

    if not pw_col:
        conn.close()
        templates = request.app.state.templates
        return templates.TemplateResponse("login.html", {"request": request, "error": "Credenciales inválidas"})

    if has_is_active:
        cur.execute(f"SELECT id, {pw_col} AS pw, is_active FROM users WHERE username=?", (username,))
    else:
        cur.execute(f"SELECT id, {pw_col} AS pw, 1 AS is_active FROM users WHERE username=?", (username,))

    u = cur.fetchone()
    conn.close()

    if not u:
        templates = request.app.state.templates
        return templates.TemplateResponse("login.html", {"request": request, "error": "Credenciales inválidas"})

    if int(u["is_active"] or 0) == 0:
        templates = request.app.state.templates
        return templates.TemplateResponse("login.html", {"request": request, "error": "Usuario desactivado (contacta con el administrador)"})

    if not verify_password(password, u["pw"]):
        templates = request.app.state.templates
        return templates.TemplateResponse("login.html", {"request": request, "error": "Credenciales inválidas"})

    response = RedirectResponse(url="/", status_code=303)
    token = sign_session(request.app.state.serializer, user_id=int(u["id"]))
    response.set_cookie("xurgical_session", token, httponly=True, samesite="lax")
    return response

@router.post("/logout")
def logout():
    resp = RedirectResponse(url="/login", status_code=303)
    resp.delete_cookie("xurgical_session")
    return resp
