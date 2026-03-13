"""
SGAM - Sistema de Gestión de Asistencias Médicas
Módulo: core.py
Responsabilidad: Motor de cruce de datos y asignación de estatus por día/empleado.
"""

import pandas as pd
from datetime import date, timedelta, datetime, time
from typing import Optional


# ──────────────────────────────────────────────
# Paleta de colores semáforo (HEX sin #)
# ──────────────────────────────────────────────
COLOR_MAP = {
    "ASISTENCIA":    {"hex": "FFD966", "label": "Asistencia",        "emoji": "🟡"},
    "RETARDO":       {"hex": "FF4B4B", "label": "Retardo",           "emoji": "🔴"},
    "FALTA":         {"hex": "FF8C00", "label": "Falta",             "emoji": "🟠"},
    "NO_LABORABLE":  {"hex": "D9D9D9", "label": "No Laborable",      "emoji": "⚪"},
    "VACACIONES":    {"hex": "70AD47", "label": "Vacaciones",        "emoji": "🟢"},
    "INCAPACIDAD":   {"hex": "9DC3E6", "label": "Incapacidad",       "emoji": "🔵"},
    "PERMISO":       {"hex": "FFE699", "label": "Permiso",           "emoji": "🟡"},
    "COMISION":      {"hex": "BDD7EE", "label": "Comisión",          "emoji": "🔵"},
    "OTRO":          {"hex": "E2EFDA", "label": "Otro",              "emoji": "⚫"},
    "SIN_DATOS":     {"hex": "FFFFFF", "label": "Sin datos",         "emoji": "◻️"},
}

# Mapeo de tipo de incidencia (texto libre del Excel) → clave interna
INCIDENCIA_TIPO_MAP = {
    "vacacion": "VACACIONES",
    "vacaciones": "VACACIONES",
    "incapacidad": "INCAPACIDAD",
    "permiso": "PERMISO",
    "comision": "COMISION",
    "comisión": "COMISION",
}


# ──────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────
def _expandir_incidencias(df_incidencias: pd.DataFrame) -> dict[tuple, str]:
    """
    Expande los rangos Fecha_Inicio → Fecha_Fin de cada incidencia
    en un dict {(id_institucional, fecha): tipo_incidencia}.
    """
    mapa = {}
    for _, row in df_incidencias.iterrows():
        id_inst = str(row["ID_Institucional"]).strip()
        tipo_raw = str(row.get("Tipo_Incidencia", "OTRO")).lower().strip()
        tipo = INCIDENCIA_TIPO_MAP.get(tipo_raw, "OTRO")
        f_ini = row["Fecha_Inicio"]
        f_fin = row["Fecha_Fin"]

        if not f_ini or not f_fin:
            continue

        # Asegurar que sean objetos date
        if isinstance(f_ini, datetime):
            f_ini = f_ini.date()
        if isinstance(f_fin, datetime):
            f_fin = f_fin.date()

        delta = (f_fin - f_ini).days + 1
        for i in range(delta):
            dia = f_ini + timedelta(days=i)
            mapa[(id_inst, dia)] = tipo

    return mapa


def _construir_set_guardias(df_guardias: pd.DataFrame) -> set[tuple]:
    """
    Construye un set de (id_institucional, fecha) con días laborables programados.
    """
    guardias = set()
    for _, row in df_guardias.iterrows():
        id_inst = str(row["ID_Institucional"]).strip()
        fecha = row["Fecha_Guardia"]
        if isinstance(fecha, datetime):
            fecha = fecha.date()
        if fecha:
            guardias.add((id_inst, fecha))
    return guardias


def _construir_indice_scanner(df_scanner: pd.DataFrame,
                               catalogo: pd.DataFrame) -> dict[tuple, dict]:
    """
    Construye dict {(id_institucional, fecha): {'checkin': time, 'checkout': time}}
    usando el vínculo ID_Biometrico_SIRA ↔ ID_Biometrico del escáner.
    """
    # Crear mapa biometrico → id_institucional desde el catálogo
    # Se usa ID_Biometrico_SIRA como clave de cruce
    bio_to_id = {}
    for _, row in catalogo.iterrows():
        bio = str(row["ID_Biometrico_SIRA"]).strip()
        # ID Institucional puede no existir en catálogo directamente,
        # en ese caso usamos el mismo biométrico como fallback
        id_inst = str(row.get("ID_Institucional", row["ID_Biometrico_SIRA"])).strip()
        bio_to_id[bio] = id_inst

    indice = {}
    for _, row in df_scanner.iterrows():
        bio = str(row["ID_Biometrico"]).strip()
        fecha = row["Fecha"]
        if isinstance(fecha, datetime):
            fecha = fecha.date()

        id_inst = bio_to_id.get(bio, bio)   # Fallback: usar el biométrico como ID

        checkin = row.get("Hora_CheckIn")
        checkout = row.get("Hora_CheckOut")

        indice[(id_inst, fecha)] = {
            "checkin": checkin if pd.notna(checkin) else None,
            "checkout": checkout if pd.notna(checkout) else None,
        }

    return indice


def _minutos_desde_medianoche(t: Optional[time]) -> Optional[int]:
    """Convierte un objeto time a minutos desde medianoche."""
    if t is None:
        return None
    return t.hour * 60 + t.minute


def _evaluar_asistencia(checkin: Optional[time],
                         hora_entrada_esperada: time,
                         tolerancia_min: int) -> str:
    """
    Dado el checkin real y la hora esperada, retorna el estatus:
    ASISTENCIA | RETARDO | FALTA
    """
    if checkin is None:
        return "FALTA"

    minutos_real = _minutos_desde_medianoche(checkin)
    minutos_esperado = _minutos_desde_medianoche(hora_entrada_esperada)

    diferencia = minutos_real - minutos_esperado

    if diferencia <= tolerancia_min:
        return "ASISTENCIA"
    else:
        return "RETARDO"


# ──────────────────────────────────────────────
# Motor Principal
# ──────────────────────────────────────────────
def procesar_asistencias(datos: dict, reglas: dict) -> pd.DataFrame:
    """
    Motor principal del SGAM.

    Parámetros:
        datos: dict con claves 'catalogo', 'incidencias', 'guardias', 'scanner'
        reglas: dict con tolerancias extraídas de extraer_reglas()

    Retorna:
        DataFrame con columnas:
            ID_Institucional, Nombre_Completo, Tipo, Especialidad_Base,
            Fecha, Estatus, Color_Hex, Label, Notas
    """
    catalogo     = datos["catalogo"]
    df_incidencias = datos["incidencias"]
    df_guardias   = datos["guardias"]
    df_scanner    = datos["scanner"]

    tolerancia_min = reglas.get("tolerancia_retardo_min", 10)
    hora_entrada_default = time(8, 0, 0)   # 08:00 AM por defecto

    # Pre-calcular estructuras de búsqueda rápida
    mapa_incidencias = _expandir_incidencias(df_incidencias)
    set_guardias     = _construir_set_guardias(df_guardias)
    indice_scanner   = _construir_indice_scanner(df_scanner, catalogo)

    # Determinar rango de fechas del mes a evaluar
    todas_fechas_guardias = [f for (_, f) in set_guardias]
    todas_fechas_scanner  = [f for (_, f) in indice_scanner.keys()]
    todas_fechas = todas_fechas_guardias + todas_fechas_scanner

    if not todas_fechas:
        raise ValueError("No se encontraron fechas en guardias ni en el escáner. Verifique los archivos.")

    fecha_min = min(todas_fechas)
    fecha_max = max(todas_fechas)

    # Construir el rango completo del mes
    fecha_inicio_mes = fecha_min.replace(day=1)
    # Último día del mes
    if fecha_max.month == 12:
        fecha_fin_mes = date(fecha_max.year + 1, 1, 1) - timedelta(days=1)
    else:
        fecha_fin_mes = date(fecha_max.year, fecha_max.month + 1, 1) - timedelta(days=1)

    dias_mes = [fecha_inicio_mes + timedelta(days=i)
                for i in range((fecha_fin_mes - fecha_inicio_mes).days + 1)]

    # Normalizar IDs institucionales del catálogo
    # Usamos ID_Biometrico_SIRA como ID primario (puede coincidir con ID_Institucional)
    registros = []

    for _, empleado in catalogo.iterrows():
        id_bio = str(empleado["ID_Biometrico_SIRA"]).strip()
        nombre = str(empleado.get("Nombre_Completo", "")).strip()
        tipo   = str(empleado.get("Tipo", "")).strip()
        especialidad = str(empleado.get("Especialidad_Base", "")).strip()
        estatus_emp = str(empleado.get("Estatus", "")).strip().upper()

        # Solo procesar empleados activos
        if estatus_emp not in ["ACTIVO", "ACTIVE", "1", "SI", "SÍ", ""]:
            continue

        # Usar ID_Biometrico_SIRA como ID_Institucional para cruce
        id_inst = id_bio

        for dia in dias_mes:
            notas = ""

            # ── PRIORIDAD 1: ¿Existe incidencia registrada? ──────────────
            if (id_inst, dia) in mapa_incidencias:
                estatus_dia = mapa_incidencias[(id_inst, dia)]
                notas = "Incidencia registrada"

            # ── PRIORIDAD 2: ¿Tiene turno programado ese día? ────────────
            elif (id_inst, dia) not in set_guardias:
                estatus_dia = "NO_LABORABLE"

            # ── PRIORIDAD 3: Evaluar asistencia con escáner ──────────────
            else:
                marca = indice_scanner.get((id_inst, dia))

                if marca is None:
                    estatus_dia = "FALTA"
                    notas = "Sin registro biométrico"
                else:
                    checkin  = marca.get("checkin")
                    estatus_dia = _evaluar_asistencia(
                        checkin, hora_entrada_default, tolerancia_min
                    )
                    if checkin:
                        notas = f"CheckIn: {checkin}"
                    if marca.get("checkout"):
                        notas += f" | CheckOut: {marca['checkout']}"

            color_info = COLOR_MAP.get(estatus_dia, COLOR_MAP["SIN_DATOS"])

            registros.append({
                "ID_Institucional":  id_inst,
                "Nombre_Completo":   nombre,
                "Tipo":              tipo,
                "Especialidad_Base": especialidad,
                "Fecha":             dia,
                "Estatus":           estatus_dia,
                "Color_Hex":         color_info["hex"],
                "Label":             color_info["label"],
                "Emoji":             color_info["emoji"],
                "Notas":             notas,
            })

    if not registros:
        raise ValueError("No se generaron registros. Verifique que el catálogo tenga empleados activos.")

    df_resultado = pd.DataFrame(registros)
    df_resultado["Fecha"] = pd.to_datetime(df_resultado["Fecha"])
    return df_resultado


# ──────────────────────────────────────────────
# Estadísticas Resumen
# ──────────────────────────────────────────────
def calcular_resumen(df_resultado: pd.DataFrame) -> dict:
    """
    Calcula estadísticas agregadas del período procesado.

    Retorna dict con:
        conteos: {estatus: cantidad}
        porcentajes: {estatus: %}
        total_dias_laborables: int
        empleados: int
    """
    # Solo días laborables para el resumen de asistencia
    laborables = df_resultado[df_resultado["Estatus"] != "NO_LABORABLE"]
    conteos = laborables["Estatus"].value_counts().to_dict()
    total = len(laborables)

    porcentajes = {k: round((v / total * 100), 1) if total > 0 else 0
                   for k, v in conteos.items()}

    return {
        "conteos": conteos,
        "porcentajes": porcentajes,
        "total_dias_laborables": total,
        "empleados": df_resultado["ID_Institucional"].nunique(),
        "mes": df_resultado["Fecha"].dt.to_period("M").iloc[0] if len(df_resultado) else None,
    }


def filtrar_por_empleado(df_resultado: pd.DataFrame,
                          termino: str) -> pd.DataFrame:
    """
    Filtra el DataFrame de resultados por ID o nombre (búsqueda parcial).
    """
    termino = termino.strip().lower()
    mask = (
        df_resultado["ID_Institucional"].str.lower().str.contains(termino, na=False) |
        df_resultado["Nombre_Completo"].str.lower().str.contains(termino, na=False)
    )
    return df_resultado[mask]
