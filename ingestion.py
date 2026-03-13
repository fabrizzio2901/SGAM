"""
SGAM - Sistema de Gestión de Asistencias Médicas
Módulo: ingestion.py
Responsabilidad: Lectura, validación y normalización de archivos fuente.
"""

import pandas as pd
from pathlib import Path
from typing import Optional


# ──────────────────────────────────────────────
# Constantes de columnas esperadas
# ──────────────────────────────────────────────
COLS_REGLAS = ["Regla de Negocio", "Descripción"]

COLS_CATALOGO = [
    "ID_Biometrico_SIRA", "Nombre_Completo", "Tipo",
    "Especialidad_Base", "Grado", "Universidad", "Estatus", "Vigencia"
]

COLS_INCIDENCIAS = [
    "ID_Institucional", "Tipo_Incidencia",
    "Fecha_Inicio", "Fecha_Fin", "Destino_o_Servicio", "Notas_Motivo"
]

COLS_GUARDIAS = [
    "Fecha_Guardia", "ID_Institucional",
    "Servicio_Cubierto", "TIPO"
]

COLS_SCANNER = [
    "ID_Biometrico", "Fecha", "Hora_CheckIn", "Hora_CheckOut"
]


# ──────────────────────────────────────────────
# Helpers internos
# ──────────────────────────────────────────────
def _validar_columnas(df: pd.DataFrame, esperadas: list[str], nombre_hoja: str) -> None:
    """Valida que las columnas obligatorias existan en el DataFrame."""
    faltantes = [c for c in esperadas if c not in df.columns]
    if faltantes:
        raise ValueError(
            f"[{nombre_hoja}] Faltan columnas obligatorias: {faltantes}\n"
            f"Columnas encontradas: {list(df.columns)}"
        )


def _limpiar_unnamed(df: pd.DataFrame) -> pd.DataFrame:
    """Elimina columnas 'Unnamed' generadas por Excel."""
    cols_limpias = [c for c in df.columns if not str(c).startswith("Unnamed")]
    return df[cols_limpias]


def _to_date(series: pd.Series) -> pd.Series:
    """Convierte una serie a tipo fecha, manejando formatos mixtos."""
    return pd.to_datetime(series, dayfirst=False, errors="coerce").dt.date


def _to_time(series: pd.Series) -> pd.Series:
    """Convierte una serie a tipo hora."""
    return pd.to_datetime(series, format="%H:%M:%S", errors="coerce").dt.time


# ──────────────────────────────────────────────
# Lectura del Archivo Maestro
# ──────────────────────────────────────────────
def cargar_archivo_maestro(ruta: str) -> dict:
    """
    Lee el archivo Estructura_Maestra_Hospital2.xlsx con sus 4 hojas.

    Retorna un dict con claves:
        'reglas', 'catalogo', 'incidencias', 'guardias'

    Lanza ValueError si falta alguna hoja o columna obligatoria.
    """
    ruta = Path(ruta)
    if not ruta.exists():
        raise FileNotFoundError(f"Archivo maestro no encontrado: {ruta}")

    try:
        xls = pd.ExcelFile(ruta)
    except Exception as e:
        raise IOError(f"No se pudo abrir el archivo maestro: {e}")

    hojas_requeridas = {
        "1_Reglas_ID": None,
        "2_Catalogo_Personal": None,
        "3_Registro_Incidencias": None,
        "4_Rol_Guardias": None,
    }

    for hoja in hojas_requeridas:
        if hoja not in xls.sheet_names:
            raise ValueError(
                f"Hoja '{hoja}' no encontrada en el archivo maestro.\n"
                f"Hojas disponibles: {xls.sheet_names}"
            )
        hojas_requeridas[hoja] = pd.read_excel(xls, sheet_name=hoja, header=0)

    # ── Hoja 1: Reglas ──────────────────────────────
    df_reglas = _limpiar_unnamed(hojas_requeridas["1_Reglas_ID"])
    _validar_columnas(df_reglas, COLS_REGLAS, "1_Reglas_ID")
    df_reglas.dropna(how="all", inplace=True)

    # ── Hoja 2: Catálogo Personal ───────────────────
    df_catalogo = _limpiar_unnamed(hojas_requeridas["2_Catalogo_Personal"])
    _validar_columnas(df_catalogo, COLS_CATALOGO, "2_Catalogo_Personal")
    df_catalogo.dropna(subset=["ID_Biometrico_SIRA"], inplace=True)
    df_catalogo["ID_Biometrico_SIRA"] = df_catalogo["ID_Biometrico_SIRA"].astype(str).str.strip()

    # ── Hoja 3: Incidencias ─────────────────────────
    df_incidencias = _limpiar_unnamed(hojas_requeridas["3_Registro_Incidencias"])
    _validar_columnas(df_incidencias, COLS_INCIDENCIAS, "3_Registro_Incidencias")
    df_incidencias.dropna(subset=["ID_Institucional", "Fecha_Inicio", "Fecha_Fin"], inplace=True)
    df_incidencias["Fecha_Inicio"] = _to_date(df_incidencias["Fecha_Inicio"])
    df_incidencias["Fecha_Fin"] = _to_date(df_incidencias["Fecha_Fin"])
    df_incidencias["ID_Institucional"] = df_incidencias["ID_Institucional"].astype(str).str.strip()

    # ── Hoja 4: Rol de Guardias ─────────────────────
    df_guardias = _limpiar_unnamed(hojas_requeridas["4_Rol_Guardias"])
    _validar_columnas(df_guardias, COLS_GUARDIAS, "4_Rol_Guardias")
    df_guardias.dropna(subset=["Fecha_Guardia", "ID_Institucional"], inplace=True)
    df_guardias["Fecha_Guardia"] = _to_date(df_guardias["Fecha_Guardia"])
    df_guardias["ID_Institucional"] = df_guardias["ID_Institucional"].astype(str).str.strip()

    return {
        "reglas": df_reglas,
        "catalogo": df_catalogo,
        "incidencias": df_incidencias,
        "guardias": df_guardias,
    }


# ──────────────────────────────────────────────
# Lectura del Reporte del Escáner Biométrico
# ──────────────────────────────────────────────
def cargar_reporte_scanner(ruta: str) -> pd.DataFrame:
    """
    Lee el archivo del lector biométrico (Excel o CSV).

    Normaliza columnas y tipos de dato.
    Retorna DataFrame con columnas: ID_Biometrico, Fecha, Hora_CheckIn, Hora_CheckOut
    """
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
    df["Fecha"] = _to_date(df["Fecha"])

    # Hora_CheckIn y Hora_CheckOut pueden ser string "HH:MM:SS" o datetime
    for col_hora in ["Hora_CheckIn", "Hora_CheckOut"]:
        if df[col_hora].dtype == object:
            df[col_hora] = pd.to_datetime(df[col_hora], format="%H:%M:%S", errors="coerce").dt.time
        else:
            df[col_hora] = pd.to_datetime(df[col_hora], errors="coerce").dt.time

    return df


# ──────────────────────────────────────────────
# Extractor de Reglas de Negocio
# ──────────────────────────────────────────────
def extraer_reglas(df_reglas: pd.DataFrame) -> dict:
    """
    Parsea la hoja de reglas y retorna un dict con parámetros de tolerancia.

    Claves esperadas (si existen):
        tolerancia_retardo_min  → int, minutos permitidos de retraso
        tolerancia_salida_min   → int, minutos antes de salida permitidos
    """
    reglas = {
        "tolerancia_retardo_min": 10,   # Valor default
        "tolerancia_salida_min": 10,
    }

    df = df_reglas.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # Mapeo de nombres posibles en la hoja hacia las claves internas
    mapeo = {
        "tolerancia retardo": "tolerancia_retardo_min",
        "retardo": "tolerancia_retardo_min",
        "tolerancia salida": "tolerancia_salida_min",
        "salida anticipada": "tolerancia_salida_min",
    }

    for _, row in df.iterrows():
        regla_raw = str(row.get("Regla de Negocio", "")).lower().strip()
        descripcion = str(row.get("Descripción", "")).strip()

        for clave_busq, clave_dict in mapeo.items():
            if clave_busq in regla_raw:
                # Intentar extraer número de la descripción
                import re
                numeros = re.findall(r"\d+", descripcion)
                if numeros:
                    reglas[clave_dict] = int(numeros[0])
                break

    return reglas
