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

    # 2) localizar fila de cabecera (escanear hasta 100 filas buscando palabras clave)
    header_row = None
    kw_c = ["CODIGO PRODUCTO", "CODIGO", "COD", "REF", "REFERENCIA", "MODELO", "PN", "PART NUMBER", "ARTICULO"]
    kw_d = ["DENOMINACION", "DESCRIPCION", "NOMBRE", "PRODUCTO", "INSTRUMENTO"]

    for i in range(min(100, len(raw))):
        row = raw.iloc[i].tolist()
        # Normalizamos cada celda de la fila para comparar
        row_norm = [_norm(str(v)) for v in row if v is not None and str(v).strip() != ""]
        
        has_code = any(k in row_norm for k in kw_c)
        has_desc = any(k in row_norm for k in kw_d)
        
        if has_code and has_desc:
            header_row = i
            break

    if header_row is None:
        # Fallback: Si no encuentro cabeceras, intento ver si es una tabla plana básica (col 0 y col 1)
        if len(raw.columns) >= 2:
            # Forzamos cabecera manual para que el resto del script no explote
            raw.columns = ["CODIGO", "DESCRIPCION"] + [f"COL_{i}" for i in range(2, len(raw.columns))]
            df = raw.copy()
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
            "No se ha podido localizar la tabla de datos.\n"
            "Asegúrese de que el Excel tenga columnas llamadas 'Modelo/Código' y 'Denominación/Descripción'."
        )

    # 3) leer con cabecera detectada
    df = pd.read_excel(path_excel, header=header_row, dtype=str)
    df = df.dropna(how="all")
    if df.empty:
        raise ValueError("Se detectó la cabecera, pero no hay filas con datos debajo.")

    # 4) mapear columnas de forma flexible (Búsqueda por coincidencia o contenido)
    colmap = {_norm(c): c for c in df.columns}
    all_cols_norm = list(colmap.keys())

    def pick(*keywords):
        # 1. Intento: Coincidencia exacta (normalizada)
        for kw in keywords:
            kn = _norm(kw)
            if kn in colmap:
                return colmap[kn]
        # 2. Intento: ¿Alguna columna CONTIENE la palabra clave?
        for kw in keywords:
            kn = _norm(kw)
            if not kn: continue
            for cn in all_cols_norm:
                if kn in cn:
                    return colmap[cn]
        return None

    c_codigo = pick("CODIGO PRODUCTO", "CODIGO", "MODELO", "REF", "REFERENCIA", "PN", "PART NUMBER", "ARTICULO")
    c_fab    = pick("FABRICANTE", "MARCA", "MANUFACTURER", "BRAND", "FAB")
    c_ns     = pick("N/S", "NS", "SN", "S/N", "N SERIE", "NUMERO SERIE", "SERIE", "SERIAL")
    c_deno   = pick("DENOMINACION", "DESCRIPCION", "NOMBRE", "PRODUCTO", "INSTRUMENTO")
    c_obs    = pick("OBSERVACIONES", "OBSERVACION", "OBS", "COMENTARIOS", "NOTAS")
    c_dm     = pick("CODIGO DATAMATRIX", "DATAMATRIX", "DATA MATRIX", "QR", "CODIGO MATRIZ", "DM")

    if not c_codigo or not c_deno:
        raise ValueError(f"No se encontraron las columnas Modelo/Denominación. Columnas detectadas: {list(df.columns)}")

    # 5) Limpieza de datos
    def clean_series(series):
        return series.where(series.notna(), "").astype(str).str.strip().replace("nan", "")

    codigo_s = clean_series(df[c_codigo])
    df = df[codigo_s != ""].copy()

    # 6) Creación del DF FINAL con los nombres que app.py espera
    df_result = pd.DataFrame()
    df_result["codigo_producto"] = clean_series(df[c_codigo])
    df_result["fabricante"] = clean_series(df[c_fab]) if c_fab else ""
    df_result["num_serie"] = clean_series(df[c_ns]) if c_ns else ""
    df_result["denominacion"] = clean_series(df[c_deno])
    df_result["observaciones"] = clean_series(df[c_obs]) if c_obs else ""
    df_result["codigo_datamatrix"] = clean_series(df[c_dm]) if c_dm else ""

    return cliente, fecha, df_result
