"""
SGAM - core.py  v2.1 (Auditado)
Motor de cruce: algoritmo semáforo hermético basado en única fuente de verdad.
"""

import pandas as pd
from datetime import date, timedelta, datetime, time
from typing import Optional
import calendar as cal

from ingestion import GUARDIAS_ORDEN, INCIDENCIA_TIPO_MAP

# ──────────────────────────────────────────────
# Paleta de colores
# ──────────────────────────────────────────────
COLOR_MAP = {
    "ASISTENCIA":    {"hex": "FFD966", "label": "Asistencia",   "emoji": "🟡"},
    "RETARDO":       {"hex": "FF8C00", "label": "Retardo",      "emoji": "🟠"},
    "FALTA":         {"hex": "FF4B4B", "label": "Falta",        "emoji": "🔴"},
    "NO_LABORABLE":  {"hex": "D9D9D9", "label": "No Laborable", "emoji": "⚪"},
    "VACACIONES":    {"hex": "70AD47", "label": "Vacaciones",   "emoji": "🟢"},
    "INCAPACIDAD":   {"hex": "FFE699", "label": "Incapacidad",  "emoji": "💛"},
    "PERMISO":       {"hex": "9DC3E6", "label": "Permiso",      "emoji": "🔵"},
    "COMISION":      {"hex": "F4B942", "label": "Comisión",     "emoji": "🟤"},
    "ROTACION":      {"hex": "FFB6C1", "label": "Rotación",     "emoji": "🔄"},
    "OTRO":          {"hex": "E2EFDA", "label": "Otro",         "emoji": "⚫"},
    "SIN_DATOS":     {"hex": "FFFFFF", "label": "Sin datos",    "emoji": "◻️"},
}

UMBRAL_TURNO_NOCTURNO_HORA = 6

# ──────────────────────────────────────────────
# Helpers de guardias cíclicas
# ──────────────────────────────────────────────
def calcular_guardia_del_dia(rotacion_inicial: str, dia_del_mes: int) -> str:
    if rotacion_inicial not in GUARDIAS_ORDEN:
        return "A"
    idx_inicio = GUARDIAS_ORDEN.index(rotacion_inicial)
    idx_hoy = (idx_inicio + (dia_del_mes - 1)) % len(GUARDIAS_ORDEN)
    return GUARDIAS_ORDEN[idx_hoy]

def construir_calendario_guardias(df_rol: pd.DataFrame, dias_del_mes: list[date]) -> dict[tuple, str]:
    calendario: dict[tuple, str] = {}
    for _, row in df_rol.iterrows():
        id_emp = str(row["ID"]).strip()
        rot_ini = str(row["Rotación"]).strip().upper()
        if rot_ini not in GUARDIAS_ORDEN:
            rot_ini = "A"

        for fecha in dias_del_mes:
            guardia_hoy = calcular_guardia_del_dia(rot_ini, fecha.day)
            calendario[(id_emp, fecha)] = guardia_hoy
    return calendario

# ──────────────────────────────────────────────
# Helpers de incidencias
# ──────────────────────────────────────────────
def _expandir_incidencias(df_inc: pd.DataFrame) -> dict[tuple, dict]:
    mapa: dict[tuple, dict] = {}
    for _, row in df_inc.iterrows():
        id_emp   = str(row["ID"]).strip()
        ausencia = str(row.get("Ausencia Justificada", "")).lower().strip()
        tipo     = INCIDENCIA_TIPO_MAP.get(ausencia, "OTRO")
        nota     = str(row.get("Notas_Motivo", "")).strip()
        destino  = str(row.get("Destino", "")).strip()
        f_ini    = row["Fecha Inicio"]
        f_fin    = row["Fecha Termino"]

        if pd.isna(f_ini) or pd.isna(f_fin): continue
        if isinstance(f_ini, datetime): f_ini = f_ini.date()
        if isinstance(f_fin, datetime): f_fin = f_fin.date()

        try:
            delta = (f_fin - f_ini).days + 1
        except Exception:
            continue

        for i in range(max(delta, 1)):
            dia = f_ini + timedelta(days=i)
            # [CORRECCIÓN] En caso de incidencia superpuesta, conservamos y concatenamos notas
            if (id_emp, dia) in mapa:
                existente = mapa[(id_emp, dia)]
                mapa[(id_emp, dia)] = {
                    "tipo": tipo,
                    "nota": f"{existente['nota']} | {nota}".strip(" |"),
                    "destino": destino or existente['destino']
                }
            else:
                mapa[(id_emp, dia)] = {"tipo": tipo, "nota": nota, "destino": destino}
    return mapa

# ──────────────────────────────────────────────
# Helpers de escáner
# ──────────────────────────────────────────────
def _construir_indice_scanner(df_scanner: pd.DataFrame, id_map: dict[str, str]) -> dict[tuple, list[dict]]:
    indice: dict[tuple, list[dict]] = {}
    for _, row in df_scanner.iterrows():
        bio   = str(row["ID_Biometrico"]).strip()
        fecha = row["Fecha"]
        if isinstance(fecha, datetime): fecha = fecha.date()
        if pd.isna(fecha): continue

        id_cat   = id_map.get(bio, bio)
        checkin  = row.get("Hora_CheckIn")
        checkout = row.get("Hora_CheckOut")
        checkin  = checkin  if (checkin  is not None and not pd.isna(checkin))  else None
        checkout = checkout if (checkout is not None and not pd.isna(checkout)) else None

        nocturno = False
        if checkout is not None:
            if getattr(checkout, "hour", 99) < UMBRAL_TURNO_NOCTURNO_HORA:
                nocturno = True

        marca = {"checkin": checkin, "checkout": checkout, "nocturno": nocturno}
        indice.setdefault((id_cat, fecha), []).append(marca)

        if nocturno:
            sig = fecha + timedelta(days=1)
            cont = {"checkin": None, "checkout": checkout, "nocturno": True, "continuacion": True}
            indice.setdefault((id_cat, sig), []).append(cont)
    return indice

def _min_desde_medianoche(t: Optional[time]) -> Optional[int]:
    if t is None: return None
    return t.hour * 60 + t.minute

# ──────────────────────────────────────────────
# Evaluación de turno
# ──────────────────────────────────────────────
def _evaluar_turno(marcas: list[dict], hora_esperada: time, tol_retardo: int, tol_falta: int) -> tuple[str, str]:
    if not marcas:
        return "FALTA", "Sin registro biométrico"

    reales      = [m for m in marcas if not m.get("continuacion")]
    cont        = [m for m in marcas if m.get("continuacion")]

    if not reales and cont:
        co = cont[0].get("checkout")
        return "ASISTENCIA", f"Continuación turno nocturno | CheckOut: {co}"

    checkins = [m["checkin"] for m in reales if m.get("checkin") is not None]
    if not checkins:
        notas = " | ".join(f"CheckOut:{m['checkout']}" for m in reales if m.get("checkout"))
        return "ASISTENCIA", notas or "Sin CheckIn registrado"

    checkin_ppal = min(checkins, key=lambda t: _min_desde_medianoche(t))
    diff = (_min_desde_medianoche(checkin_ppal) or 0) - (_min_desde_medianoche(hora_esperada) or 0)

    if diff <= tol_retardo:
        estatus = "ASISTENCIA"
    elif diff <= tol_falta:
        estatus = "RETARDO"
    else:
        estatus = "FALTA"

    notas_partes = []
    for i, m in enumerate(reales, 1):
        pfx = f"T{i}" if len(reales) > 1 else ""
        if m.get("checkin"):  notas_partes.append(f"{pfx}CheckIn:{m['checkin']}")
        if m.get("checkout"): notas_partes.append(f"{pfx}CheckOut:{m['checkout']}")
        if m.get("nocturno"): notas_partes.append("(nocturno)")
    if diff > 0: notas_partes.append(f"+{diff}min")

    return estatus, " | ".join(notas_partes)

# ──────────────────────────────────────────────
# Motor Principal
# ──────────────────────────────────────────────
def procesar_asistencias(datos: dict, reglas: dict) -> pd.DataFrame:
    catalogo     = datos["catalogo"]
    df_rol       = datos.get("rol_guardias", pd.DataFrame())
    df_inc       = datos["incidencias"]
    df_scanner   = datos["scanner"]

    tol_retardo = reglas.get("tolerancia_retardo_min", 10)
    tol_falta   = reglas.get("tolerancia_falta_min",   15)
    horarios    = reglas.get("horarios_guardia", {})
    hora_default = time(7, 0)

    # ── [CORRECCIÓN CRÍTICA] Barrera estricta Anti-Huérfanos ─────────
    ids_catalogo = set(catalogo[catalogo["_activo"]]["ID"].astype(str).str.strip().unique())
    id_map: dict[str, str] = {id_: id_ for id_ in ids_catalogo}
    df_scanner = df_scanner[df_scanner["ID_Biometrico"].astype(str).str.strip().isin(ids_catalogo)].copy()

    # ── [CORRECCIÓN CRÍTICA] Límites Mensuales Robustos (Moda) ────────
    s_fechas = pd.to_datetime(df_scanner["Fecha"], errors="coerce").dropna()
    if s_fechas.empty:
        raise ValueError("El reporte del escáner no tiene fechas válidas aplicables al catálogo actual.")

    mes_objetivo = s_fechas.dt.to_period("M").mode()[0]
    df_scanner = df_scanner[s_fechas.dt.to_period("M") == mes_objetivo].copy()

    fecha_inicio_mes = date(mes_objetivo.year, mes_objetivo.month, 1)
    ultimo_dia = cal.monthrange(mes_objetivo.year, mes_objetivo.month)[1]
    fecha_fin_mes = date(mes_objetivo.year, mes_objetivo.month, ultimo_dia)
    dias_mes = [fecha_inicio_mes + timedelta(days=i) for i in range((fecha_fin_mes - fecha_inicio_mes).days + 1)]

    # ── Pre-calcular estructuras ──────────────────────────────────────
    mapa_incidencias  = _expandir_incidencias(df_inc)
    calendario_guard  = construir_calendario_guardias(df_rol, dias_mes)
    indice_scanner    = _construir_indice_scanner(df_scanner, id_map)

    registros = []

    for _, empleado in catalogo.iterrows():
        if not empleado.get("_activo", False):
            continue

        id_emp    = str(empleado["ID"]).strip()
        nombre    = str(empleado.get("Nombre completo", "")).strip()
        tipo      = str(empleado.get("Tipo de personal", "")).strip()
        esp       = str(empleado.get("Especialidad", "")).strip()
        subesp    = str(empleado.get("Subespecialidad", "")).strip()
        alta_esp  = str(empleado.get("Alta Especialidad", "")).strip()
        grado     = str(empleado.get("Grado", "")).strip()
        periodo   = str(empleado.get("Periodo ingreso", "")).strip()
        foto      = str(empleado.get("Foto_Ruta", "")).strip()

        rot_row = df_rol[df_rol["ID"].astype(str).str.strip() == id_emp]
        rot_inicial = str(rot_row.iloc[0]["Rotación"]).strip().upper() if not rot_row.empty else "A"

        guardia_anterior = None # Validacion de saltos

        for dia in dias_mes:
            notas          = ""
            guardia_tipo   = ""
            turno_servicio = ""

            if (id_emp, dia) in mapa_incidencias:
                inc_info    = mapa_incidencias[(id_emp, dia)]
                estatus_dia = inc_info["tipo"]
                partes = ["Incidencia registrada"]
                if inc_info["destino"]: partes.append(f"Destino: {inc_info['destino']}")
                if inc_info["nota"]:    partes.append(inc_info["nota"])
                notas = " | ".join(partes)
                guardia_tipo = calcular_guardia_del_dia(rot_inicial, dia.day) # Mantener tracking matemático

            else:
                if (id_emp, dia) in calendario_guard:
                    guardia_tipo   = calendario_guard[(id_emp, dia)]
                    if not rot_row.empty:
                        area = str(rot_row.iloc[0].get("Area asignada", "")).strip()
                        irp  = str(rot_row.iloc[0].get("IRP", "")).strip()
                        turno_servicio = " · ".join(filter(None, [area, irp]))
                else:
                    guardia_tipo   = calcular_guardia_del_dia("A", dia.day)

                hora_esp = horarios.get(guardia_tipo, {}).get("hora_inicio", hora_default)
                marcas   = indice_scanner.get((id_emp, dia), [])
                estatus_dia, notas = _evaluar_turno(marcas, hora_esp, tol_retardo, tol_falta)

            # [VALIDACIÓN] Doble guardia al cruzar meses o modificar plantillas mid-ciclo
            if guardia_anterior and guardia_anterior == guardia_tipo and estatus_dia not in ["VACACIONES", "INCAPACIDAD", "PERMISO", "COMISION", "ROTACION"]:
                notas = f"ALERTA SECUENCIA: Doble guardia ({guardia_tipo}) detectada | " + notas
            guardia_anterior = guardia_tipo

            color_info = COLOR_MAP.get(estatus_dia, COLOR_MAP["SIN_DATOS"])

            registros.append({
                "ID":              id_emp,
                "Nombre_Completo": nombre,
                "Tipo":            tipo,
                "Especialidad":    esp,
                "Subespecialidad": subesp,
                "Alta_Especialidad": alta_esp,
                "Grado":           grado,
                "Periodo_Ingreso": periodo,
                "Foto_Ruta":       foto,
                "Rot_Inicial":     rot_inicial,
                "Fecha":           dia,
                "Guardia_Tipo":    guardia_tipo,
                "Turno_Servicio":  turno_servicio,
                "Estatus":         estatus_dia,
                "Color_Hex":       color_info["hex"],
                "Label":           color_info["label"],
                "Emoji":           color_info["emoji"],
                "Notas":           notas,
            })

    if not registros:
        raise ValueError("No se generaron registros. Verifique que haya empleados activos y datos de escáner en el mismo periodo.")

    df = pd.DataFrame(registros)
    df["Fecha"] = pd.to_datetime(df["Fecha"])
    return df

# ──────────────────────────────────────────────
# Utilidades de filtrado y resumen
# ──────────────────────────────────────────────
def calcular_resumen(df: pd.DataFrame) -> dict:
    lab     = df[df["Estatus"] != "NO_LABORABLE"]
    conteos = lab["Estatus"].value_counts().to_dict()
    total   = len(lab)
    return {
        "conteos":               conteos,
        "porcentajes":           {k: round(v/total*100, 1) if total else 0 for k, v in conteos.items()},
        "total_dias_laborables": total,
        "empleados":             df["ID"].nunique(),
        "mes":                   df["Fecha"].dt.to_period("M").iloc[0] if len(df) else None,
    }

def filtrar_por_empleado(df: pd.DataFrame, termino: str) -> pd.DataFrame:
    t = termino.strip().lower()
    mask = (df["ID"].astype(str).str.lower().str.contains(t, na=False) | df["Nombre_Completo"].str.lower().str.contains(t, na=False))
    return df[mask]

def obtener_tipos_unicos(df: pd.DataFrame) -> list[str]:
    OBSOLETOS = {"médico adscrito", "medico adscrito"}
    return sorted(t for t in df["Tipo"].dropna().unique() if str(t).lower() not in OBSOLETOS)

def obtener_especialidades_unicas(df: pd.DataFrame) -> list[str]:
    return sorted(df["Especialidad"].replace("", pd.NA).dropna().unique().tolist())