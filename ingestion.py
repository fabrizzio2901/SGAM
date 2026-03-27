"""
SGAM - Sistema de Gestión de Asistencias Médicas
Módulo: ingestion.py
Responsabilidad: Lectura, validación y normalización de archivos fuente.

Reglas de negocio v1.2:
  Catálogo Personal:
    - Internos: campo Especialidad_Base puede quedar vacío (no se valida como obligatorio).
    - Residentes: Grado indica el año (R1, R2, …).
    - Periodo_Ingreso: valores A (Primavera) o B (Otoño).
  Registro de Incidencias:
    - Tipo 'Rotación' requiere Destino_o_Servicio obligatorio.
    - Notas_Motivo es campo de justificación libre; se preserva y propaga.
  Rol de Guardias:
    - Tipos A, B y C con horarios configurables (hora_inicio, hora_fin por tipo).
    - Los horarios se leen de Reglas_ID si están definidos, o usan defaults.
"""

import re
from datetime import time
from pathlib import Path

import pandas as pd


# ──────────────────────────────────────────────
# Columnas obligatorias y opcionales por hoja
# ──────────────────────────────────────────────
COLS_REGLAS = ["Regla de Negocio", "Descripción"]

COLS_CATALOGO = [
    "ID_Biometrico_SIRA", "Nombre_Completo", "Tipo",
    "Grado", "Universidad", "Estatus", "Vigencia",
]
# Columnas opcionales del catálogo (se crean vacías si no existen):
COLS_CATALOGO_OPCIONALES = ["Especialidad_Base", "Periodo_Ingreso", "Foto_Ruta"]

COLS_INCIDENCIAS = [
    "ID_Institucional", "Tipo_Incidencia",
    "Fecha_Inicio", "Fecha_Fin", "Destino_o_Servicio", "Notas_Motivo",
]

COLS_GUARDIAS = [
    "Fecha_Guardia", "ID_Institucional",
    "Servicio_Cubierto", "TIPO",
]

COLS_SCANNER = [
    "ID_Biometrico", "Fecha", "Hora_CheckIn", "Hora_CheckOut",
]

# ── Catálogos de dominio ─────────────────────────────────────────────────
TIPOS_PERSONAL_NORM = {
    "INTERNO":         "Interno",
    "RESIDENTE":       "Residente",
    "MÉDICO ADSCRITO": "Médico Adscrito",
    "MEDICO ADSCRITO": "Médico Adscrito",
}
PERIODOS_VALIDOS = {"A", "B"}
INCIDENCIAS_CON_DESTINO_OBLIGATORIO = {"rotacion", "rotación"}

# ── Horarios de guardia por defecto (sobreescritos desde Reglas_ID) ───────
HORARIOS_GUARDIA_DEFAULT: dict[str, dict] = {
    "A": {"hora_inicio": time(8, 0),  "hora_fin": time(15, 0),
          "label": "Turno A – Matutino  (08:00–15:00)"},
    "B": {"hora_inicio": time(15, 0), "hora_fin": time(21, 0),
          "label": "Turno B – Vespertino (15:00–21:00)"},
    "C": {"hora_inicio": time(21, 0), "hora_fin": time(8, 0),
          "label": "Turno C – Nocturno   (21:00–08:00)"},
}


# ──────────────────────────────────────────────
# Helpers internos
# ──────────────────────────────────────────────
def _validar_columnas(df: pd.DataFrame, obligatorias: list[str],
                       nombre_hoja: str, opcionales: list[str] | None = None) -> None:
    faltantes = [c for c in obligatorias if c not in df.columns]
    if faltantes:
        raise ValueError(
            f"[{nombre_hoja}] Faltan columnas obligatorias: {faltantes}\n"
            f"Columnas encontradas: {list(df.columns)}"
        )
    if opcionales:
        ausentes = [c for c in opcionales if c not in df.columns]
        if ausentes:
            print(f"  [INFO] {nombre_hoja}: columnas opcionales no presentes (se crearán vacías): {ausentes}")


def _limpiar_unnamed(df: pd.DataFrame) -> pd.DataFrame:
    return df[[c for c in df.columns if not str(c).startswith("Unnamed")]]


def _to_date(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, dayfirst=False, errors="coerce").dt.date


def _normalizar_tipo(tipo_raw: str) -> str:
    upper = str(tipo_raw).upper().strip()
    for clave, valor in TIPOS_PERSONAL_NORM.items():
        if clave in upper:
            return valor
    return str(tipo_raw).strip()


def _normalizar_periodo(val: str) -> str:
    v = str(val).upper().strip()
    return v if v in PERIODOS_VALIDOS else ""


# ──────────────────────────────────────────────
# Carga del Archivo Maestro
# ──────────────────────────────────────────────
def cargar_archivo_maestro(ruta: str) -> dict:
    """
    Lee Estructura_Maestra_Hospital2.xlsx (4 hojas).

    Retorna dict con claves:
        'reglas', 'catalogo', 'incidencias', 'guardias', 'alertas'

    'alertas': lista de advertencias no-bloqueantes (ej. rotaciones sin destino).
    """
    ruta = Path(ruta)
    if not ruta.exists():
        raise FileNotFoundError(f"Archivo maestro no encontrado: {ruta}")

    try:
        xls = pd.ExcelFile(ruta)
    except Exception as e:
        raise IOError(f"No se pudo abrir el archivo maestro: {e}")

    hojas_req = {
        "1_Reglas_ID":            None,
        "2_Catalogo_Personal":    None,
        "3_Registro_Incidencias": None,
        "4_Rol_Guardias":         None,
    }
    for hoja in hojas_req:
        if hoja not in xls.sheet_names:
            raise ValueError(
                f"Hoja '{hoja}' no encontrada.\nHojas disponibles: {xls.sheet_names}"
            )
        hojas_req[hoja] = pd.read_excel(xls, sheet_name=hoja, header=0)

    alertas: list[str] = []

    # ── 1. Reglas ────────────────────────────────────────────────────────
    df_reglas = _limpiar_unnamed(hojas_req["1_Reglas_ID"])
    _validar_columnas(df_reglas, COLS_REGLAS, "1_Reglas_ID")
    df_reglas.dropna(how="all", inplace=True)

    # ── 2. Catálogo Personal ─────────────────────────────────────────────
    df_cat = _limpiar_unnamed(hojas_req["2_Catalogo_Personal"])
    _validar_columnas(df_cat, COLS_CATALOGO, "2_Catalogo_Personal",
                       opcionales=COLS_CATALOGO_OPCIONALES)
    df_cat.dropna(subset=["ID_Biometrico_SIRA"], inplace=True)
    df_cat["ID_Biometrico_SIRA"] = df_cat["ID_Biometrico_SIRA"].astype(str).str.strip()

    # Crear columnas opcionales vacías si no existen
    for col in COLS_CATALOGO_OPCIONALES:
        if col not in df_cat.columns:
            df_cat[col] = ""

    # Normalizar Tipo
    df_cat["Tipo"] = df_cat["Tipo"].fillna("").astype(str).apply(_normalizar_tipo)

    # Regla: Internos → Especialidad_Base puede quedar vacía (no es error)
    # Regla: Residentes → Grado debe coincidir con R\d+
    mask_res = df_cat["Tipo"].str.lower().str.contains("residente", na=False)
    grados_res = df_cat.loc[mask_res, "Grado"].fillna("").astype(str)
    sin_grado_std = grados_res[~grados_res.str.upper().str.match(r"^R\d+$")].tolist()
    if sin_grado_std:
        alertas.append(
            f"Residentes con grado no estándar (se esperan R1, R2…): {sin_grado_std}"
        )

    # Normalizar Periodo_Ingreso
    df_cat["Periodo_Ingreso"] = (
        df_cat["Periodo_Ingreso"].fillna("").astype(str).apply(_normalizar_periodo)
    )
    # Foto_Ruta: ruta relativa o absoluta a imagen del médico
    df_cat["Foto_Ruta"] = df_cat["Foto_Ruta"].fillna("").astype(str).str.strip()

    # Limpiar Especialidad_Base
    df_cat["Especialidad_Base"] = df_cat["Especialidad_Base"].fillna("").astype(str).str.strip()

    # ── 3. Registro de Incidencias ───────────────────────────────────────
    df_inc = _limpiar_unnamed(hojas_req["3_Registro_Incidencias"])
    _validar_columnas(df_inc, COLS_INCIDENCIAS, "3_Registro_Incidencias")
    df_inc.dropna(subset=["ID_Institucional", "Fecha_Inicio", "Fecha_Fin"], inplace=True)

    df_inc["Fecha_Inicio"]       = _to_date(df_inc["Fecha_Inicio"])
    df_inc["Fecha_Fin"]          = _to_date(df_inc["Fecha_Fin"])
    df_inc["ID_Institucional"]   = df_inc["ID_Institucional"].astype(str).str.strip()
    df_inc["Notas_Motivo"]       = df_inc["Notas_Motivo"].fillna("").astype(str).str.strip()
    df_inc["Destino_o_Servicio"] = df_inc["Destino_o_Servicio"].fillna("").astype(str).str.strip()

    # Validar: Rotaciones requieren destino
    for _, row in df_inc.iterrows():
        tipo_inc = str(row.get("Tipo_Incidencia", "")).lower().strip()
        if tipo_inc in INCIDENCIAS_CON_DESTINO_OBLIGATORIO:
            destino = str(row.get("Destino_o_Servicio", "")).strip()
            if not destino or destino.lower() in ("nan", "none", ""):
                alertas.append(
                    f"⚠ Rotación sin destino: ID {row.get('ID_Institucional')} "
                    f"({row.get('Fecha_Inicio')} → {row.get('Fecha_Fin')})"
                )

    # ── 4. Rol de Guardias ───────────────────────────────────────────────
    df_guard = _limpiar_unnamed(hojas_req["4_Rol_Guardias"])
    _validar_columnas(df_guard, COLS_GUARDIAS, "4_Rol_Guardias")
    df_guard.dropna(subset=["Fecha_Guardia", "ID_Institucional"], inplace=True)

    df_guard["Fecha_Guardia"]    = _to_date(df_guard["Fecha_Guardia"])
    df_guard["ID_Institucional"] = df_guard["ID_Institucional"].astype(str).str.strip()
    df_guard["TIPO"]             = df_guard["TIPO"].fillna("").astype(str).str.strip().str.upper()

    tipos_inv = df_guard.loc[
        df_guard["TIPO"].str.strip().ne("") & ~df_guard["TIPO"].isin(["A", "B", "C"]),
        "TIPO"
    ].unique().tolist()
    if tipos_inv:
        alertas.append(
            f"⚠ Guardias con TIPO no reconocido (esperado A, B o C): {tipos_inv}. "
            f"Se usarán horarios default."
        )

    return {
        "reglas":      df_reglas,
        "catalogo":    df_cat,
        "incidencias": df_inc,
        "guardias":    df_guard,
        "alertas":     alertas,
    }


# ──────────────────────────────────────────────
# Carga del Reporte del Escáner Biométrico
# ──────────────────────────────────────────────
def cargar_reporte_scanner(ruta: str) -> pd.DataFrame:
    """Lee el reporte del lector biométrico (Excel o CSV)."""
    ruta = Path(ruta)
    if not ruta.exists():
        raise FileNotFoundError(f"Archivo del escáner no encontrado: {ruta}")

    try:
        if ruta.suffix.lower() in [".xlsx", ".xls"]:
            df = pd.read_excel(ruta, header=0)
        elif ruta.suffix.lower() == ".csv":
            df = pd.read_csv(ruta, encoding="utf-8-sig")
        else:
            raise ValueError(f"Formato no soportado: {ruta.suffix}. Use .xlsx o .csv")
    except Exception as e:
        raise IOError(f"No se pudo leer el reporte del escáner: {e}")

    df = _limpiar_unnamed(df)
    _validar_columnas(df, COLS_SCANNER, "Reporte_Scanner")
    df.dropna(subset=["ID_Biometrico", "Fecha"], inplace=True)
    df["ID_Biometrico"] = df["ID_Biometrico"].astype(str).str.strip()
    df["Fecha"]         = _to_date(df["Fecha"])

    for col_hora in ["Hora_CheckIn", "Hora_CheckOut"]:
        if df[col_hora].dtype == object:
            df[col_hora] = pd.to_datetime(
                df[col_hora], format="%H:%M:%S", errors="coerce"
            ).dt.time
        else:
            df[col_hora] = pd.to_datetime(df[col_hora], errors="coerce").dt.time

    return df


# ──────────────────────────────────────────────
# Extractor de Reglas de Negocio
# ──────────────────────────────────────────────
def extraer_reglas(df_reglas: pd.DataFrame) -> dict:
    """
    Parsea la hoja de reglas y retorna:
        tolerancia_retardo_min  → int (minutos)
        tolerancia_salida_min   → int (minutos)
        horarios_guardia        → dict {
            'A': {'hora_inicio': time, 'hora_fin': time, 'label': str},
            'B': {...}, 'C': {...}
        }
    """
    reglas: dict = {
        "tolerancia_retardo_min": 10,
        "tolerancia_salida_min":  10,
        "horarios_guardia":       {k: dict(v) for k, v in HORARIOS_GUARDIA_DEFAULT.items()},
    }

    df = df_reglas.copy()
    df.columns = [str(c).strip() for c in df.columns]

    for _, row in df.iterrows():
        clave_raw   = str(row.get("Regla de Negocio", "")).lower().strip()
        descripcion = str(row.get("Descripción", "")).strip()

        if "retardo" in clave_raw:
            nums = re.findall(r"\d+", descripcion)
            if nums:
                reglas["tolerancia_retardo_min"] = int(nums[0])

        elif "salida anticipada" in clave_raw or "tolerancia salida" in clave_raw:
            nums = re.findall(r"\d+", descripcion)
            if nums:
                reglas["tolerancia_salida_min"] = int(nums[0])

        # Horarios de guardias A, B, C — formato "HH:MM - HH:MM" en Descripción
        for tipo_g in ["A", "B", "C"]:
            if f"guardia {tipo_g.lower()}" in clave_raw or \
               f"turno {tipo_g.lower()}" in clave_raw:
                horas = re.findall(r"\d{1,2}:\d{2}", descripcion)
                if len(horas) >= 2:
                    try:
                        h_ini = time(*map(int, horas[0].split(":")))
                        h_fin = time(*map(int, horas[1].split(":")))
                        reglas["horarios_guardia"][tipo_g]["hora_inicio"] = h_ini
                        reglas["horarios_guardia"][tipo_g]["hora_fin"]    = h_fin
                        reglas["horarios_guardia"][tipo_g]["label"] = (
                            f"Turno {tipo_g} ({horas[0]}–{horas[1]})"
                        )
                    except ValueError:
                        pass

    return reglas
