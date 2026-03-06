import re
import os
from datetime import datetime
from pathlib import Path

def _clean_trz(s):
    if not s:
        return ""
    s = str(s).strip()
    # quita backslashes que vienen escapando comillas
    s = s.replace("\\'", "").replace('\\"', "").replace("\\", "")
    # quita comillas envolventes repetidas
    while len(s) >= 2 and ((s[0] == "'" and s[-1] == "'") or (s[0] == '"' and s[-1] == '"')):
        s = s[1:-1].strip()
    # quita comillas sueltas en extremos
    s = s.strip().strip("'").strip('"').strip()
    return s

def _last5_digits_from_dm(dm: str) -> str:
    """Extrae los últimos 5 dígitos del DataMatrix. Si hay menos de 5 dígitos, usa los últimos 5 caracteres."""
    s = (dm or "").strip()
    digits = re.findall(r"\d", s)
    if len(digits) >= 5:
        return "".join(digits[-5:])
    s2 = re.sub(r"\s+", "", s)
    return (s2[-5:] if len(s2) >= 5 else s2)

def _build_nombre_trazabilidad(prefijo_nombre: str, codigo_datamatrix: str) -> str:
    pref = (prefijo_nombre or "").strip()
    suf = _last5_digits_from_dm(codigo_datamatrix)
    if not pref and not suf:
        return ""
    return f"{pref}{suf}"

def format_fecha(value):
    if not value or value == "-":
        return "-"
    try:
        # Detectar si es un string y tratar de parsearlo
        s = str(value).strip()
        if not s: return "-"
        
        # Intentar varios formatos comunes
        dt = None
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
            try:
                dt = datetime.strptime(s, fmt)
                break
            except:
                continue
        
        if not dt:
            try:
                dt = datetime.fromisoformat(s.replace("Z", "+00:00"))
            except:
                return s # Si falla todo, devuelve el original
        
        return dt.strftime("%d-%m-%Y")
    except:
        return value

def _norm_codigo(x: str) -> str:
    """Normaliza códigos de artículo."""
    s = (x or '').strip().upper()
    for pfx in ("RP", "MT"):
        if s.startswith(pfx):
            s = s[len(pfx):]
            break
    s = s.lstrip(" -_\t")
    return s

def _codigo_variants(codigo_norm: str) -> list:
    """Genera variantes de búsqueda para un código normalizado."""
    c = (codigo_norm or '').strip().upper()
    if not c:
        return []

    out = []
    def add(x: str):
        x = (x or '').strip().upper()
        if x and x not in out:
            out.append(x)

    add(c)
    add(c.replace(' ', ''))
    if c.endswith('R') and len(c) > 1:
        add(c[:-1])
        add(c[:-1].replace(' ', ''))
    else:
        add(c + 'R')
        add((c + 'R').replace(' ', ''))
    return out
