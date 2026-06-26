"""
SGAM - ingestion.py  v3.5 (Blindado contra FutureWarnings + Soporte de Festivos)
Plantilla única + escáner dual + Parser Optimizado + Ampliación Personal.
"""

import re
from datetime import date, timedelta, datetime, time
from pathlib import Path
from typing import Optional
import pandas as pd

# ── [CORRECCIÓN] Blindaje para Pandas 2.1.0+ y 3.0 ──────────────
try:
    pd.set_option('future.no_silent_downcasting', True)
except Exception:
    pass
# ────────────────────────────────────────────────────────────────

# ──────────────────────────────────────────────
# Constantes de dominio
# ──────────────────────────────────────────────
TIPOS_PERSONAL_NORM = {
    "INTERNO":   "Interno",
    "RESIDENTE": "Residente",
}
PERIODOS_VALIDOS = {"A", "B"}

INCIDENCIA_TIPO_MAP = {
    "vacaciones":         "VACACIONES",
    "vacacion":           "VACACIONES",
    "incapacidad":        "INCAPACIDAD",
    "permisos":           "PERMISO",
    "permiso":            "PERMISO",
    "comisiones":         "COMISION",
    "comision":           "COMISION",
    "comisión":           "COMISION",
    "rotacion de campo":  "ROTACION",
    "rotación de campo":  "ROTACION",
    "rotacion externa":   "ROTACION",
    "rotación externa":   "ROTACION",
    "rotacion":           "ROTACION",
    "rotación":           "ROTACION",
    "no aplica":          "OTRO",
}

GUARDIAS_ORDEN = ["A", "B", "C", "D"]

HORARIOS_GUARDIA_DEFAULT = {
    "A": {"hora_inicio": time(7, 0), "hora_fin": time(15, 0), "label": "Guardia A"},
    "B": {"hora_inicio": time(7, 0), "hora_fin": time(15, 0), "label": "Guardia B"},
    "C": {"hora_inicio": time(7, 0), "hora_fin": time(15, 0), "label": "Guardia C"},
    "D": {"hora_inicio": time(7, 0), "hora_fin": time(15, 0), "label": "Guardia D"},
}

GRADOS_RESIDENTE_VALIDOS = re.compile(r"^[Rr][1-5]$")

# ──────────────────────────────────────────────
# Helpers y Limpiadores
# ──────────────────────────────────────────────
def _limpiar_unnamed(df: pd.DataFrame) -> pd.DataFrame:
    return df[[c for c in df.columns if not str(c).startswith("Unnamed")]]

def _norm_id(val) -> str:
    try:
        if val is None or pd.isna(val): return ""
        s = str(val).strip()
        if not s or s.lower() in ("nan", "nat", "none"): return ""
        if s.endswith(".0"): s = s[:-2]
        if s.isdigit(): return str(int(s))
        return s
    except Exception:
        return str(val)

def _to_date(val) -> Optional[date]:
    if val is None or (isinstance(val, float) and pd.isna(val)): return None
    if isinstance(val, (datetime, pd.Timestamp)):
        try: return val.date()
        except Exception: return None
    if isinstance(val, date): return val
    s = str(val).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try: return datetime.strptime(s, fmt).date()
        except ValueError: pass
    return None

def _series_to_date(series: pd.Series) -> pd.Series:
    return series.apply(_to_date)

def _normalizar_tipo(val: str) -> str:
    upper = str(val).upper().strip()
    for key, norm in TIPOS_PERSONAL_NORM.items():
        if key in upper: return norm
    return str(val).strip()

def _normalizar_periodo(val: str) -> str:
    v = str(val).upper().strip()
    return v if v in PERIODOS_VALIDOS else ""

def _estatus_info(tipo: str, estatus_raw: str) -> tuple[bool, str]:
    s = str(estatus_raw).lower().strip()
    if tipo == "Residente":
        if GRADOS_RESIDENTE_VALIDOS.match(s): return True, s.upper()
        if s in ("activo", "active", "1", "si", "sí"): return True, ""
        return False, ""
    else:
        return s in ("activo", "active", "1", "si", "sí"), ""

def _validar_columnas(df: pd.DataFrame, obligatorias: list, nombre: str):
    faltantes = [c for c in obligatorias if c not in df.columns]
    if faltantes:
        raise ValueError(f"[{nombre}] Faltan columnas: {faltantes}\nEncontradas: {list(df.columns)}")

def _parse_hora_segura(val) -> Optional[time]:
    try:
        if val is None or pd.isna(val): return None
        if isinstance(val, time): return val
        if isinstance(val, datetime): return val.time()
        s = str(val).strip()
        if not s or s.lower() in ("nan", "nat", "none"): return None
        m = re.search(r"(\d{1,2}):(\d{2})", s)
        if m:
            h = int(m.group(1)) % 24
            minuto = int(m.group(2))
            return time(h, minuto)
        frac = float(s)
        frac = frac % 1
        total_sec = int(frac * 86400)
        h = (total_sec // 3600) % 24
        minuto = (total_sec % 3600) // 60
        return time(h, minuto)
    except Exception:
        pass
    return None

def _limpiar_na(val):
    """Helper para normalizar vacíos en texto hacia 'N/A' visual."""
    s = str(val).strip() if pd.notna(val) else ""
    return s if s and s.lower() not in ("nan", "nat", "none", "no aplica", "n/a", "") else "N/A"

# ──────────────────────────────────────────────
# Carga de Plantilla ÚNICA Institucional
# ──────────────────────────────────────────────
def cargar_plantilla(ruta: str) -> dict:
    ruta = Path(ruta)
    if not ruta.exists(): raise FileNotFoundError(f"Plantilla no encontrada: {ruta}")

    try:
        xls = pd.ExcelFile(str(ruta))
    except Exception as e:
        raise IOError(f"No se pudo abrir: {e}")

    hojas_req = ["1_Catalogo_Personal", "2_Rol_Guardias", "3_Registro_Incidencias", "4_Vacaciones", "5_Rotaciones"]
    for h in hojas_req:
        if h not in xls.sheet_names:
            raise ValueError(f"Hoja '{h}' no encontrada en {ruta.name}.\nHojas presentes: {xls.sheet_names}")

    alertas: list[str] = []

    # ── 1. Catálogo ────────────────────────────────────────────────────
    df_cat = _limpiar_unnamed(pd.read_excel(xls, "1_Catalogo_Personal"))
    df_cat.columns = [str(c).strip() for c in df_cat.columns]

    _validar_columnas(df_cat, ["ID", "Nombre completo", "Estatus", "Tipo de personal"], "1_Catalogo_Personal")

    # [CORRECCIÓN] Los nombres aquí NO coincidían con los encabezados reales de
    # la plantilla institucional ("Sub especialidad" vs "Subespecialidad" sin
    # espacio, "Alta especialidad" vs "Alta Especialidad" con mayúscula, y
    # "Periodo de ingreso" vs "Periodo ingreso" sin "de"). Esto generaba
    # columnas fantasma con "N/A" que nadie usaba, mientras las columnas
    # REALES de la plantilla quedaban sin pasar por limpieza de NaN: una
    # celda vacía llegaba como NaN crudo hasta los reportes Excel (incluido
    # el Reporte Maestro) y se imprimía literalmente como "nan" en vez de
    # "N/A". Se corrigen los nombres para que coincidan con la plantilla real
    # y se conservan alias de compatibilidad por si una plantilla antigua
    # todavía usa la grafía anterior.
    ALIAS_COLUMNAS = {
        "Sub especialidad":   "Subespecialidad",
        "Alta especialidad":  "Alta Especialidad",
        "Periodo de ingreso": "Periodo ingreso",
    }
    for alias, real in ALIAS_COLUMNAS.items():
        if alias in df_cat.columns and real not in df_cat.columns:
            df_cat.rename(columns={alias: real}, inplace=True)

    nuevas_columnas = ["Universidad", "Especialidad", "Subespecialidad", "Alta Especialidad", "Periodo ingreso"]
    for col in nuevas_columnas:
        if col not in df_cat.columns:
            df_cat[col] = "N/A"

    foto_col = next((c for c in df_cat.columns if "foto" in c.lower()), None)
    df_cat["Foto_Ruta"] = df_cat[foto_col].apply(lambda x: str(x) if pd.notna(x) else "") if foto_col else ""

    df_cat.dropna(subset=["ID"], inplace=True)
    df_cat["ID"] = df_cat["ID"].apply(_norm_id)
    df_cat = df_cat[df_cat["ID"].ne("")]

    dupes_cat = df_cat[df_cat.duplicated(subset=["ID"], keep=False)]["ID"].unique().tolist()
    if dupes_cat:
        alertas.append(f"Catálogo: IDs duplicados encontrados y purgados: {dupes_cat[:5]}")
        df_cat.drop_duplicates(subset=["ID"], keep="first", inplace=True)

    df_cat["Tipo de personal"]   = df_cat["Tipo de personal"].apply(lambda x: str(x) if pd.notna(x) else "").apply(_normalizar_tipo)
    df_cat["Nombre completo"]    = df_cat["Nombre completo"].apply(lambda x: str(x) if pd.notna(x) else "").str.strip()
    
    # [MEJORA] Limpiamos y aseguramos "N/A" visual (ahora sobre las columnas REALES)
    df_cat["Universidad"]        = df_cat["Universidad"].apply(_limpiar_na)
    df_cat["Especialidad"]       = df_cat["Especialidad"].apply(_limpiar_na)
    df_cat["Subespecialidad"]    = df_cat["Subespecialidad"].apply(_limpiar_na)
    df_cat["Alta Especialidad"]  = df_cat["Alta Especialidad"].apply(_limpiar_na)

    # [CORRECCIÓN] "Periodo ingreso" no siempre es A/B: para Residentes la
    # plantilla real usa "R". La normalización anterior forzaba cualquier
    # valor fuera de {A,B} a "N/A", borrando ese dato válido sin avisar.
    # Ahora sólo se limpia el vacío/NaN a "N/A" y se conserva el valor capturado;
    # se valida A/B únicamente para personal de tipo "Interno" (que es donde
    # aplica esa regla) y se alerta sin descartar el dato.
    df_cat["Periodo ingreso"] = df_cat["Periodo ingreso"].apply(_limpiar_na).apply(
        lambda x: x.upper() if isinstance(x, str) and x not in ("N/A",) else x
    )

    df_cat["_activo"] = False
    df_cat["Grado"]   = ""
    for idx, row in df_cat.iterrows():
        activo, grado = _estatus_info(row["Tipo de personal"], str(row["Estatus"]))
        df_cat.at[idx, "_activo"] = activo
        df_cat.at[idx, "Grado"]   = grado

    res_mal_grado = df_cat[(df_cat["Tipo de personal"] == "Residente") & df_cat["_activo"] & ~df_cat["Grado"].str.match(r"^R[1-5]$", na=False)]["ID"].tolist()
    if res_mal_grado: alertas.append(f"Catálogo: Residentes con grado no estándar (IDs): {res_mal_grado[:10]}")

    periodo_invalido = df_cat[
        (df_cat["Tipo de personal"] == "Interno") & df_cat["_activo"] &
        ~df_cat["Periodo ingreso"].isin(PERIODOS_VALIDOS | {"N/A"})
    ]["ID"].tolist()
    if periodo_invalido:
        alertas.append(f"Catálogo: Internos con 'Periodo ingreso' distinto de A/B (revisar captura): {periodo_invalido[:10]}")

    # ── 2. Rol de Guardias ─────────────────────────────────────────────
    df_rol = _limpiar_unnamed(pd.read_excel(xls, "2_Rol_Guardias"))
    df_rol.columns = [str(c).strip() for c in df_rol.columns]
    _validar_columnas(df_rol, ["ID", "Rotación"], "2_Rol_Guardias")

    df_rol.dropna(subset=["ID"], inplace=True)
    df_rol["ID"]       = df_rol["ID"].apply(_norm_id)
    df_rol             = df_rol[df_rol["ID"].ne("")]

    dupes_rol = df_rol[df_rol.duplicated(subset=["ID"], keep=False)]["ID"].unique().tolist()
    if dupes_rol:
        alertas.append(f"Rol Guardias: IDs duplicados encontrados y purgados: {dupes_rol[:5]}")
        df_rol.drop_duplicates(subset=["ID"], keep="first", inplace=True)

    df_rol["Rotación"] = df_rol["Rotación"].apply(lambda x: str(x) if pd.notna(x) else "").str.strip().str.upper()

    area_col = next((c for c in df_rol.columns if c.lower() in ("area asignada", "especialidad")), None)
    df_rol["Area asignada"] = df_rol[area_col].apply(lambda x: str(x) if pd.notna(x) else "").str.strip() if area_col else ""

    irp_col = next((c for c in df_rol.columns if "irp" in c.lower()), None)
    df_rol["IRP"] = df_rol[irp_col].apply(lambda x: str(x) if pd.notna(x) else "").str.strip() if irp_col else ""

    tipos_inv = df_rol.loc[df_rol["Rotación"].ne("") & ~df_rol["Rotación"].isin(GUARDIAS_ORDEN), "Rotación"].unique().tolist()
    if tipos_inv: alertas.append(f"Rol Guardias: Rotaciones inválidas ignoradas (esperado A-D): {tipos_inv}")

    activos_ids = set(df_cat[df_cat["_activo"]]["ID"])
    rol_ids = set(df_rol["ID"])
    sin_rol = activos_ids - rol_ids
    if sin_rol:
        alertas.append(f"Regla rota: Empleados activos sin rol asignado (se asume 'A'): {list(sin_rol)[:5]}")

    # ── 3+4+5. Incidencias unificadas ──────────────────────────────────
    dfs_inc = []
    for hoja in ["3_Registro_Incidencias", "4_Vacaciones", "5_Rotaciones"]:
        df_h = _limpiar_unnamed(pd.read_excel(xls, hoja))
        df_h.columns = [str(c).strip() for c in df_h.columns]
        if df_h.empty or len(df_h.columns) == 0: continue

        _validar_columnas(df_h, ["ID", "Ausencia Justificada", "Fecha Inicio", "Fecha Termino"], hoja)

        df_h.dropna(subset=["ID"], inplace=True)
        df_h["ID"] = df_h["ID"].apply(_norm_id)
        df_h       = df_h[df_h["ID"].ne("")]
        df_h["Fecha Inicio"]         = _series_to_date(df_h["Fecha Inicio"])
        df_h["Fecha Termino"]        = _series_to_date(df_h["Fecha Termino"])

        df_h = df_h[df_h["Fecha Inicio"].notna() & df_h["Fecha Termino"].notna()].copy()
        mask_inv = df_h["Fecha Inicio"] > df_h["Fecha Termino"]
        if mask_inv.any():
            inv_ids = df_h[mask_inv]["ID"].unique().tolist()
            alertas.append(f"Incidencias ({hoja}): Fechas invertidas ignoradas para IDs: {inv_ids[:5]}")
            df_h = df_h[~mask_inv]

        df_h["Ausencia Justificada"] = df_h["Ausencia Justificada"].apply(lambda x: str(x) if pd.notna(x) else "").str.lower().str.strip()
        df_h = df_h[~df_h["Ausencia Justificada"].isin(["", "nan"])]

        notas_col    = next((c for c in df_h.columns if "observac" in c.lower()), None)
        cantidad_col = next((c for c in df_h.columns if "cantidad" in c.lower()), None)
        df_h["Notas_Motivo"] = df_h[notas_col].apply(lambda x: str(x) if pd.notna(x) else "").str.strip() if notas_col else ""

        if cantidad_col:
            def _merge_nota(r):
                cant = str(r[cantidad_col]).strip() if pd.notna(r[cantidad_col]) else ""
                if cant in ("", "nan"): return r["Notas_Motivo"]
                base = r["Notas_Motivo"].strip()
                return f"{base} · {cant}" if base else cant
            df_h["Notas_Motivo"] = df_h.apply(_merge_nota, axis=1)

        df_h["Destino"] = df_h["Destino"].apply(lambda x: str(x) if pd.notna(x) else "").str.strip() if "Destino" in df_h.columns else ""
        dfs_inc.append(df_h[["ID", "Ausencia Justificada", "Fecha Inicio", "Fecha Termino", "Notas_Motivo", "Destino"]])

    df_inc = pd.concat(dfs_inc, ignore_index=True) if dfs_inc else pd.DataFrame(columns=["ID", "Ausencia Justificada", "Fecha Inicio", "Fecha Termino", "Notas_Motivo", "Destino"])

    # ── 6. Días Festivos [NUEVA MEJORA] ────────────────────────────────
    dic_festivos = {}
    if "6_Dias_Festivos" in xls.sheet_names:
        try:
            # Esperamos estrictamente las columnas A (Motivo) y B (Fecha)
            df_festivos = pd.read_excel(xls, sheet_name="6_Dias_Festivos", usecols="A:B")
            df_festivos.columns = ['Motivo', 'Fecha']
            df_festivos = df_festivos.dropna(subset=['Fecha'])
            # Conversión segura a formato Date para cruce exacto en core.py
            df_festivos['Fecha'] = pd.to_datetime(df_festivos['Fecha'], errors='coerce').dt.date
            df_festivos = df_festivos.dropna(subset=['Fecha'])
            dic_festivos = df_festivos.set_index('Fecha')['Motivo'].to_dict()
        except Exception as e:
            alertas.append(f"Días Festivos: Hubo un problema procesando la pestaña - {e}")
    else:
        alertas.append("No se detectó la pestaña '6_Dias_Festivos'. Se omitirán descansos festivos automáticos.")


    return {
        "catalogo":     df_cat,
        "rol_guardias": df_rol,
        "incidencias":  df_inc,
        "festivos":     dic_festivos, # Exponemos el diccionario hacia el exterior
        "alertas":      alertas,
    }

def filtrar_scanner_por_catalogo(df_scanner: pd.DataFrame, df_catalogo: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    ids_validos = set(df_catalogo.loc[df_catalogo["_activo"], "ID"].astype(str).str.strip())
    df = df_scanner.copy()
    df["ID_Biometrico"] = df["ID_Biometrico"].astype(str).str.strip()

    huerfanos = df[~df["ID_Biometrico"].isin(ids_validos)].copy()
    validos = df[df["ID_Biometrico"].isin(ids_validos)].copy()

    return validos.reset_index(drop=True), huerfanos.reset_index(drop=True)

# ──────────────────────────────────────────────
# Parser Biométrico (El Magico)
# ──────────────────────────────────────────────
def _procesar_matriz_zkteco(xl, sheet_name) -> pd.DataFrame:
    df_raw = pd.read_excel(xl, sheet_name=sheet_name, header=None)
    periodo_str = datetime.now().strftime("%Y-%m")

    # [CORRECCIÓN FUTURE WARNING] Reemplazo de fillna por list comprehension pura de Python
    for i in range(min(15, len(df_raw))):
        fila_str = " ".join([str(x) if pd.notna(x) else "" for x in df_raw.iloc[i]]).lower()
        m = re.search(r"(\d{4}-\d{2})-\d{2}", fila_str)
        if m:
            periodo_str = m.group(1)
            break

    dias_cols = {}
    start_row = 0
    for i in range(min(15, len(df_raw))):
        fila = [str(x) if pd.notna(x) else "" for x in df_raw.iloc[i]]
        for j, val in enumerate(fila):
            v = val.replace(".0", "").strip()
            if v == "1":
                if j+1 < len(fila) and fila[j+1].replace(".0", "").strip() == "2":
                    for k in range(j, len(fila)):
                        dk = fila[k].replace(".0", "").strip()
                        if dk.isdigit(): dias_cols[k] = int(dk)
                    start_row = i + 1
                    break
        if dias_cols: break

    if not dias_cols: raise ValueError("No se detectó la matriz de días (1..31) en el Excel.")

    registros = []
    current_id = None

    for i in range(start_row, len(df_raw)):
        fila = [str(x) if pd.notna(x) else "" for x in df_raw.iloc[i]]
        fila_lower = [x.lower().strip() for x in fila]

        if any("id:" in x for x in fila_lower):
            for j, celda in enumerate(fila_lower):
                if "id:" in celda:
                    nums = re.findall(r"\d+", celda)
                    if nums: current_id = nums[-1]
                    else:
                        for k in range(1, 6):
                            if j+k < len(fila) and fila[j+k].strip().isdigit():
                                current_id = fila[j+k].strip()
                                break
                    break
        elif current_id:
            tiene_horas = any(re.search(r"\d{2}:\d{2}", fila[col]) for col in dias_cols if col < len(fila))
            if tiene_horas:
                for col_idx, dia_num in dias_cols.items():
                    if col_idx < len(fila):
                        celda = fila[col_idx]
                        horas = re.findall(r"\d{2}:\d{2}", celda)
                        if horas:
                            registros.append({
                                "ID": current_id,
                                "Fecha": f"{periodo_str}-{dia_num:02d}",
                                "Hora_CheckIn": horas[0],
                                "Hora_CheckOut": horas[-1] if len(horas) > 1 else None
                            })
                current_id = None

    df_final = pd.DataFrame(registros)
    if df_final.empty: return pd.DataFrame(columns=["ID_Biometrico", "Fecha", "Hora_CheckIn", "Hora_CheckOut"])

    df_final["ID"] = df_final["ID"].apply(_norm_id)
    df_final["Fecha"] = pd.to_datetime(df_final["Fecha"], errors="coerce").dt.date
    df_final["Hora_CheckIn"] = df_final["Hora_CheckIn"].apply(_parse_hora_segura)
    df_final["Hora_CheckOut"] = df_final["Hora_CheckOut"].apply(_parse_hora_segura)

    df_final.drop_duplicates(inplace=True)
    return df_final.rename(columns={"ID": "ID_Biometrico"}).reset_index(drop=True)

def cargar_reporte_scanner(ruta: str) -> pd.DataFrame:
    try:
        df = pd.read_csv(ruta) if str(ruta).lower().endswith('.csv') else pd.read_excel(ruta)
        for col in ["ID", "Fecha", "Entrada", "Salida"]:
            if col not in df.columns: df[col] = None
        df = df[df["ID"].notna()]
        df["ID"] = df["ID"].apply(_norm_id)
        df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce").dt.date
        df["Hora_CheckIn"] = df["Entrada"].apply(_parse_hora_segura)
        df["Hora_CheckOut"] = df["Salida"].apply(_parse_hora_segura)
        df = df[df["Hora_CheckIn"].notna()]
        df.drop_duplicates(inplace=True)
        return df[["ID", "Fecha", "Hora_CheckIn", "Hora_CheckOut"]].rename(columns={"ID": "ID_Biometrico"}).reset_index(drop=True)
    except Exception as e:
        raise IOError(f"Error procesando el escáner genérico: {str(e)}")

def cargar_reporte_scanner_hospital(ruta: str) -> pd.DataFrame:
    ruta = Path(ruta)
    if not ruta.exists(): raise FileNotFoundError(f"Archivo no encontrado: {ruta}")

    ext = ruta.suffix.lower()
    if ext == ".csv": return cargar_reporte_scanner(str(ruta))

    engine = "xlrd" if ext == ".xls" else "openpyxl"
    try: xl = pd.ExcelFile(str(ruta), engine=engine)
    except Exception as e: raise IOError(f"No se pudo abrir el Excel: {e}")

    hoja_matriz = None
    for sheet in xl.sheet_names:
        if "asistencia" in sheet.lower() or "eventos" in sheet.lower():
            hoja_matriz = sheet
            break

    if hoja_matriz:
        try:
            df_matriz = _procesar_matriz_zkteco(xl, hoja_matriz)
            if not df_matriz.empty: return df_matriz
        except Exception: pass

    hoja_encontrada = xl.sheet_names[0]
    for sheet in xl.sheet_names:
        if "excepciones" in sheet.lower():
            hoja_encontrada = sheet
            break

    df_raw = pd.read_excel(xl, sheet_name=hoja_encontrada, header=None)
    header_idx = -1

    # [CORRECCIÓN FUTURE WARNING]
    for i in range(min(20, len(df_raw))):
        fila_str = " ".join([str(x) if pd.notna(x) else "" for x in df_raw.iloc[i]]).lower()
        if "fecha" in fila_str and ("emp" in fila_str or "nombre" in fila_str or "id" in fila_str):
            header_idx = i
            break

    if header_idx == -1: return cargar_reporte_scanner(str(ruta))

    fila_principal = pd.Series([str(x) if pd.notna(x) else "" for x in df_raw.iloc[header_idx]]).str.lower().str.strip()
    if header_idx + 1 < len(df_raw):
        fila_sub = pd.Series([str(x) if pd.notna(x) else "" for x in df_raw.iloc[header_idx + 1]]).str.lower().str.strip()
        headers = (fila_principal + " " + fila_sub).tolist()
        df_data = df_raw.iloc[header_idx + 2:].copy()
    else:
        headers = fila_principal.tolist()
        df_data = df_raw.iloc[header_idx + 1:].copy()

    col_id, col_fecha, col_ent1, col_sal1, col_falta = None, None, None, None, None
    for idx, col in enumerate(headers):
        if not col.strip(): continue
        if col_id is None and (col == "id" or col.startswith("id ") or "emp" in col or "núm" in col): col_id = idx
        elif col_fecha is None and "fecha" in col: col_fecha = idx
        elif col_ent1 is None and ("entrada" in col or "check in" in col or "check-in" in col): col_ent1 = idx
        elif col_sal1 is None and ("salida" in col or "check out" in col or "check-out" in col) and "temprano" not in col: col_sal1 = idx
        elif col_falta is None and ("falta" in col or "ausent" in col): col_falta = idx

    if col_id is None: col_id = 0
    if col_fecha is None: col_fecha = 3
    if col_ent1 is None: col_ent1 = 4
    if col_sal1 is None: col_sal1 = 5

    df_clean = pd.DataFrame()
    try:
        df_clean["ID"] = df_data.iloc[:, col_id]
        df_clean["Fecha"] = df_data.iloc[:, col_fecha]
        df_clean["Entrada1"] = df_data.iloc[:, col_ent1]
        df_clean["Salida1"] = df_data.iloc[:, col_sal1]
    except IndexError:
        return cargar_reporte_scanner(str(ruta))

    try: df_clean["Falta_Min"] = df_data.iloc[:, col_falta] if col_falta is not None else 0
    except Exception: df_clean["Falta_Min"] = 0

    df_clean = df_clean[df_clean["ID"].notna()]
    df_clean["ID"] = df_clean["ID"].apply(_norm_id)
    df_clean = df_clean[df_clean["ID"].ne("")]
    df_clean["Fecha"] = pd.to_datetime(df_clean["Fecha"], errors="coerce").dt.date
    df_clean = df_clean[df_clean["Fecha"].notna()]
    df_clean["Falta_Min"] = pd.to_numeric(df_clean["Falta_Min"], errors="coerce").fillna(0)

    df_clean["Hora_CheckIn"] = df_clean["Entrada1"].apply(_parse_hora_segura)
    df_clean["Hora_CheckOut"] = df_clean["Salida1"].apply(_parse_hora_segura)

    mask = df_clean["Hora_CheckIn"].notna() | (df_clean["Falta_Min"] >= 480)
    df_final = df_clean[mask]
    df_final.drop_duplicates(inplace=True)
    return df_final[["ID", "Fecha", "Hora_CheckIn", "Hora_CheckOut"]].rename(columns={"ID": "ID_Biometrico"}).reset_index(drop=True)

def cargar_dos_scanners(ruta_int: str, ruta_res: str) -> pd.DataFrame:
    df1 = cargar_reporte_scanner_hospital(ruta_int)
    df2 = cargar_reporte_scanner_hospital(ruta_res)
    df = pd.concat([df1, df2], ignore_index=True)
    if df.empty: return df

    def first_valid(x):
        vals = [v for v in x if pd.notna(v)]
        return vals[0] if vals else None

    def last_valid(x):
        vals = [v for v in x if pd.notna(v)]
        return vals[-1] if vals else None

    return df.groupby(["ID_Biometrico", "Fecha"], as_index=False).agg({
        "Hora_CheckIn": first_valid,
        "Hora_CheckOut": last_valid
    })

# ──────────────────────────────────────────────
# Extractor de Reglas
# ──────────────────────────────────────────────
def extraer_reglas(df_reglas: pd.DataFrame) -> dict:
    reglas = {
        "tolerancia_retardo_min": 10,
        "tolerancia_falta_min":   15,
        "tolerancia_salida_min":  10,
        "horarios_guardia": {k: dict(v) for k, v in HORARIOS_GUARDIA_DEFAULT.items()},
    }
    if df_reglas is None or df_reglas.empty: return reglas

    df = df_reglas.copy()
    df.columns = [str(c).strip() for c in df.columns]

    def _int(val):
        nums = re.findall(r"\d+", str(val))
        return int(nums[0]) if nums else None

    def _time(val):
        m = re.search(r"(\d{1,2}):(\d{2})", str(val))
        if m:
            try: return time(int(m.group(1)), int(m.group(2)))
            except ValueError: return None
        return None

    for _, row in df.iterrows():
        clave = str(row.get("Regla de Negocio", "")).strip().upper()
        valor = str(row.get("Descripción", "")).strip()

        if clave == "TOLERANCIA_RETARDO_MIN":
            v = _int(valor)
            if v: reglas["tolerancia_retardo_min"] = v
        elif clave == "TOLERANCIA_FALTA_MIN":
            v = _int(valor)
            if v: reglas["tolerancia_falta_min"] = v
        elif clave == "TOLERANCIA_SALIDA_MIN":
            v = _int(valor)
            if v: reglas["tolerancia_salida_min"] = v
        else:
            m = re.match(r"^HORARIO_([ABCD])_(INICIO|FIN)$", clave)
            if m:
                tg, campo = m.group(1), "hora_inicio" if m.group(2) == "INICIO" else "hora_fin"
                t = _time(valor)
                if t and tg in reglas["horarios_guardia"]:
                    reglas["horarios_guardia"][tg][campo] = t
    return reglas