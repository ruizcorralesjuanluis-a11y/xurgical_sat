import os
import re
import time
import io
import csv
import json
import base64
import sys
from datetime import datetime
from typing import List, Optional
from pathlib import Path

from dotenv import load_dotenv
load_dotenv()

from fastapi import FastAPI, Request, UploadFile, File, Form, Depends, HTTPException
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel

# Rutas compatibles con PyInstaller (sys._MEIPASS) y ejecución normal
BASE_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))

from db import get_conn, get_table_columns, init_db
from security import (
    get_current_user, require_roles, verify_password, hash_password,
    make_serializer, sign_session, decode_session
)
from shared import (
    ACTIONS, _user_role, _user_id, _default_allowed_by_role, 
    _get_user_permissions_map, can_action, _users_schema, 
    _select_users_sql, _envios_has_column, _list_clientes,
    _get_cliente, _reserve_numeros_cliente, FOTOS_DIR, BASE_DIR
)
from utils import format_fecha

app = FastAPI(title="Xurgical SAT")

# Configuración básica
app.state.secret_key = os.environ.get("XURGICAL_SECRET_KEY", "dev-secret-change-me")
app.state.serializer = make_serializer(app.state.secret_key)

# Montaje de estáticos
app.mount("/static/fotos", StaticFiles(directory=str(FOTOS_DIR)), name="fotos_externas")
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")
# Migración de fotos antigua (de static/fotos a uploads/fotos)
try:
    OLD_FOTOS = BASE_DIR / "static" / "fotos"
    if OLD_FOTOS.exists() and OLD_FOTOS.is_dir() and OLD_FOTOS != FOTOS_DIR:
        import shutil
        for f in OLD_FOTOS.iterdir():
            if f.is_file() and not (FOTOS_DIR / f.name).exists():
                shutil.copy2(f, FOTOS_DIR)
except Exception as e:
    print(f"Error migrando fotos: {e}")


templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))
templates.env.filters["fecha"] = format_fecha
app.state.templates = templates

@app.on_event("startup")
def on_startup():
    init_db()
    # Asegurar tabla de permisos
    conn = get_conn(); cur = conn.cursor()
    cur.execute("""
    CREATE TABLE IF NOT EXISTS user_permissions (
      user_id INTEGER NOT NULL,
      action TEXT NOT NULL,
      allowed INTEGER NOT NULL DEFAULT 0,
      PRIMARY KEY (user_id, action),
      FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
    );
    """)
    conn.commit(); conn.close()

# Exception Handlers
@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    import traceback
    return HTMLResponse(f"<h1>Error Global del Servidor</h1><pre>{traceback.format_exc()}</pre>", status_code=500)

@app.exception_handler(HTTPException)
async def http_exception_handler(request: Request, exc: HTTPException):
    if exc.status_code == 303:
        target = exc.headers.get("Location", "/login")
        return RedirectResponse(url=target, status_code=303)
    if exc.status_code == 401:
        accept = request.headers.get("accept", "")
        if "text/html" in accept:
            return RedirectResponse(url="/login", status_code=303)
    return JSONResponse({"detail": exc.detail}, status_code=exc.status_code)

# -----------------------------
# RUTAS MODULARES
# -----------------------------
from routes import auth, dashboard, instruments, envios, admin, clientes_recogidas, extras

app.include_router(auth.router)
app.include_router(dashboard.router)
app.include_router(instruments.router)
app.include_router(envios.router)
app.include_router(admin.router)
app.include_router(clientes_recogidas.router)
app.include_router(extras.router)

@app.get("/health")
@app.head("/")
def health_check():
    return {"status": "ok"}
