import re
import pandas as pd
from datetime import datetime


def _norm(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip().upper()
    s = s.replace("\n", " ").replace("\r", " ")
    s = re.sub(r"\s+", " ", s)
    s = (s.replace("Á", "A").replace("É", "E").replace("Í", "I")
           .replace("Ó", "O").replace("Ú", "U").replace("Ü", "U"))
    s = s.replace(".", "")  # UDS. -> UDS
    return s


def _extract_cliente_y_fecha(raw: pd.DataFrame):
    cliente = None
    fecha = None

    top = raw.head(20).fillna("").astype(str)
    date_re = re.compile(r"(\d{1,2}/\d{1,2}/\d{4})")

    # Fecha
    for _, row in top.iterrows():
        for cell in row.tolist():
            m = date_re.search(cell)
            if m:
                try:
                    fecha = datetime.strptime(m.group(1), "%d/%m/%Y")
                    break
                except Exception:
                    pass
        if fecha:
            break

    # Cliente: buscamos una fila “sola” tipo SORUMEDIC
    candidate = None
    for i in range(min(20, len(top))):
        row = top.iloc[i].tolist()
        non_empty = [c.strip() for c in row if c.strip() != ""]
        if len(non_empty) == 1:
            txt = non_empty[0]
            ntxt = _norm(txt)
            if "FECHA" not in ntxt and "REPARACIONES" not in ntxt:
                candidate = txt
    cliente = candidate

    return cliente, fecha


def leer_excel_envio(path_excel: str):
    """
    Devuelve:
      cliente, fecha, df_final con columnas que app.py usa:
        codigo_producto, fabricante, num_serie, denominacion, observaciones, codigo_datamatrix
    """

    # 1) leer en bruto sin cabecera
    raw = pd.read_excel(path_excel, header=None, dtype=str)
    if raw.empty:
        raise ValueError("El Excel está vacío.")

    cliente, fecha = _extract_cliente_y_fecha(raw)

    # 2) localizar fila de cabecera (en tu excel está en la fila 5)
    header_row = None
    for i in range(min(80, len(raw))):
        row = raw.iloc[i].tolist()
        row_norm = {_norm(v) for v in row if v is not None and str(v).strip() != ""}
        # Buscamos columnas clave
        has_code = any(k in row_norm for k in ["CODIGO PRODUCTO", "CODIGO", "COD", "REF", "REFERENCIA"])
        has_desc = any(k in row_norm for k in ["DENOMINACION", "DESCRIPCION", "NOMBRE", "PRODUCTO"])
        
        if has_code and has_desc:
            header_row = i
            break

    if header_row is None:
        # Fallback: Check if row 0 has the columns directly (simple table)
        row0 = raw.iloc[0].tolist()
        row0_norm = {_norm(v) for v in row0 if v is not None}
        has_code_0 = any(k in row0_norm for k in ["CODIGO PRODUCTO", "CODIGO", "COD", "REF", "REFERENCIA"])
        has_desc_0 = any(k in row0_norm for k in ["DENOMINACION", "DESCRIPCION", "NOMBRE", "PRODUCTO"])
        
        if has_code_0 and has_desc_0:
            header_row = 0
        else:
            # Fallback EXTREMO: Si no encuentro cabeceras, asumo que es una tabla plana donde:
            # Columna 0 = Código
            # Columna 1 = Descripción
            # (Y si hay más de 2 columnas, la 2 sea num serie, etc. pero vamos a lo mínimo)
            if len(raw.columns) >= 2:
                # Forzamos cabecera manual
                raw.columns = ["CODIGO", "DESCRIPCION"] + [f"COL_{i}" for i in range(2, len(raw.columns))]
                # Devolvemos esto procesado directamtente (saltándonos el paso 3 y 4 de lectura con header)
                df = raw.copy()
                # Limpieza básica
                df["CODIGO"] = df["CODIGO"].astype(str).str.strip().replace("nan", "")
                df = df[df["CODIGO"] != ""]
                
                df_final = pd.DataFrame()
                df_final["codigo_producto"] = df["CODIGO"]
                df_final["denominacion"] = df["DESCRIPCION"].astype(str).str.strip()
                df_final["fabricante"] = ""
                df_final["num_serie"] = ""
                df_final["observaciones"] = ""
                df_final["codigo_datamatrix"] = ""
                return cliente, fecha, df_final
            
            preview = raw.head(15).fillna("").astype(str).values.tolist()
            raise ValueError(
                "No encuentro la fila de cabecera ni estructura válida.\n"
                f"Vista previa (15 primeras filas): {preview}"
            )

    # 3) leer con cabecera correcta
    df = pd.read_excel(path_excel, header=header_row, dtype=str)
    df = df.dropna(how="all")
    if df.empty:
        raise ValueError("He encontrado la cabecera, pero no hay filas de datos debajo.")

    # 4) mapear columnas
    colmap = {_norm(c): c for c in df.columns}

    def pick(*names):
        for n in names:
            k = _norm(n)
            if k in colmap:
                return colmap[k]
        return None

    c_codigo = pick("CODIGO PRODUCTO", "CODIGO", "CÓDIGO", "REF", "REFERENCIA")
    c_fab   = pick("FABRICANTE", "MARCA")
    c_ns    = pick("N/S", "NS", "N SERIE", "Nº SERIE", "NUMERO SERIE", "SERIE")
    c_deno  = pick("DENOMINACION", "DESCRIPCION", "DESCRIPCIÓN", "NOMBRE")
    c_obs   = pick("OBSERVACIONES", "OBSERVACION", "OBS")

    if not c_codigo or not c_deno:
        raise ValueError(f"Columnas detectadas: {list(df.columns)}")

    # 5) limpiar
    def clean_series(s):
        return s.where(s.notna(), "").astype(str).str.strip()

    codigo = clean_series(df[c_codigo])
    df = df[codigo != ""].copy()

    # 6) DF FINAL con los nombres EXACTOS que app.py usa
    df_final = pd.DataFrame()
    df_final["codigo_producto"] = clean_series(df[c_codigo])
    df_final["fabricante"] = clean_series(df[c_fab]) if c_fab else ""
    df_final["num_serie"] = clean_series(df[c_ns]) if c_ns else ""
    df_final["denominacion"] = clean_series(df[c_deno])
    df_final["observaciones"] = clean_series(df[c_obs]) if c_obs else ""

    # En tu Excel no hay DataMatrix → vacío
    df_final["codigo_datamatrix"] = ""

    return cliente, fecha, df_final
