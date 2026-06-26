"""
SGAM - core.py  v3.3 (Filtro Exclusión Guardias + Optimizaciones de Rendimiento y Huérfanos)
Motor de cruce unificado: Algoritmo semáforo hermético basado en única fuente de verdad.
"""

import pandas as pd
from datetime import date, timedelta, datetime, time
from typing import Optional
import calendar as cal

# Asumimos que GUARDIAS_ORDEN e INCIDENCIA_TIPO_MAP vienen de ingestion
from ingestion import GUARDIAS_ORDEN, INCIDENCIA_TIPO_MAP

# ──────────────────────────────────────────────
# Paleta de colores Institucional (Actualizada)
# ──────────────────────────────────────────────
COLOR_MAP = {
    "ASISTENCIA":    {"hex": "FFD966", "label": "Asistencia",     "emoji": "🟡"},
    "RETARDO":       {"hex": "FF8C00", "label": "Retardo",        "emoji": "🟠"},
    "FALTA":         {"hex": "FF4B4B", "label": "Falta",          "emoji": "🔴"},
    "GUARDIA":       {"hex": "8EA9DB", "label": "Guardia",        "emoji": "🏥"},
    "POST_GUARDIA":  {"hex": "D9E1F2", "label": "Post-Guardia",   "emoji": "🛌"},
    "FALTA_GUARDIA": {"hex": "C00000", "label": "Falta Guardia",  "emoji": "❌"},
    "NO_LABORABLE":  {"hex": "D9D9D9", "label": "No Laborable",   "emoji": "⚪"},
    "VACACIONES":    {"hex": "70AD47", "label": "Vacaciones",     "emoji": "🟢"},
    "INCAPACIDAD":   {"hex": "FFE699", "label": "Incapacidad",    "emoji": "💛"},
    "PERMISO":       {"hex": "9DC3E6", "label": "Permiso",        "emoji": "🔵"},
    "COMISION":      {"hex": "F4B942", "label": "Comisión",       "emoji": "🟤"},
    "ROTACION":      {"hex": "FFB6C1", "label": "Rotación",       "emoji": "🔄"},
    "FESTIVO":       {"hex": "ADD8E6", "label": "Festivo",        "emoji": "🎉"}, 
    "OTRO":          {"hex": "E2EFDA", "label": "Otro",           "emoji": "⚫"},
    "SIN_DATOS":     {"hex": "FFFFFF", "label": "Sin datos",      "emoji": "◻️"},
}

UMBRAL_TURNO_NOCTURNO_HORA = 6

# ──────────────────────────────────────────────
# Helpers de guardias cíclicas (Generador A-B-C-D)
# ──────────────────────────────────────────────
def construir_calendario_hospital(letra_inicial_mes: str, dias_del_mes: list[date]) -> dict[date, str]:
    """Genera un diccionario cíclico mapeando cada fecha a la letra de guardia (A, B, C, D)."""
    if letra_inicial_mes not in GUARDIAS_ORDEN:
        letra_inicial_mes = "A"
    
    idx_inicio = GUARDIAS_ORDEN.index(letra_inicial_mes)
    cal_hospital = {}
    
    for i, dia in enumerate(dias_del_mes):
        cal_hospital[dia] = GUARDIAS_ORDEN[(idx_inicio + i) % len(GUARDIAS_ORDEN)]
        
    return cal_hospital

def determinar_estado_ciclo(grupo_medico: str, dia: date, cal_hospital: dict[date, str]) -> str:
    """Calcula individualmente si al médico le toca Guardia, Post-Guardia u Ordinario."""
    letra_hoy = cal_hospital.get(dia)
    
    # Determinar la letra del día de ayer para el post-guardia
    idx_hoy = GUARDIAS_ORDEN.index(letra_hoy)
    letra_ayer = GUARDIAS_ORDEN[(idx_hoy - 1) % len(GUARDIAS_ORDEN)]
    
    if letra_hoy == grupo_medico:
        return 'Guardia'
    elif letra_ayer == grupo_medico:
        return 'Post-Guardia'
    else:
        return 'Ordinario'

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
            if (id_emp, dia) in mapa:
                existente = mapa[(id_emp, dia)]
                # [PRESERVADO v3.0] Alerta explícita de conflictos en lugar de sobrescritura silenciosa
                tipo_final = existente["tipo"]
                conflicto = ""
                if tipo_final != tipo:
                    conflicto = f"⚠ CONFLICTO DE INCIDENCIAS: {existente['tipo']} vs {tipo} (se conservó {tipo_final})"
                partes_nota = [p for p in (existente['nota'], nota, conflicto) if p]
                mapa[(id_emp, dia)] = {
                    "tipo": tipo_final,
                    "nota": " | ".join(partes_nota),
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
# Evaluación de turnos y Guardias
# ──────────────────────────────────────────────
def _evaluar_turno_ordinario(marcas: list[dict], hora_esperada: time, tol_retardo: int, tol_falta: int) -> tuple[str, str]:
    """Evalúa jornadas ordinarias de 07:00 AM a 17:00 PM."""
    if not marcas:
        return "FALTA", "Sin registro biométrico"

    reales = [m for m in marcas if not m.get("continuacion")]
    cont   = [m for m in marcas if m.get("continuacion")]

    if not reales and cont:
        co = cont[0].get("checkout")
        return "ASISTENCIA", f"Continuación turno nocturno | CheckOut: {co}"

    checkins = [m["checkin"] for m in reales if m.get("checkin") is not None]
    if not checkins:
        notas = " | ".join(f"CheckOut:{m['checkout']}" for m in reales if m.get("checkout"))
        return "ASISTENCIA", notas or "Sin CheckIn registrado"

    checkin_ppal = min(checkins, key=lambda t: _min_desde_medianoche(t))
    
    min_registro = _min_desde_medianoche(checkin_ppal) or 0
    min_esperado = _min_desde_medianoche(hora_esperada) or 0
    
    limite_asistencia = min_esperado + 15
    limite_retardo    = min_esperado + 30

    if min_registro <= limite_asistencia:
        estatus = "ASISTENCIA"
    elif min_registro <= limite_retardo:
        estatus = "RETARDO"
    else:
        estatus = "FALTA"

    diff = min_registro - min_esperado
    notas_partes = []
    for i, m in enumerate(reales, 1):
        pfx = f"T{i}" if len(reales) > 1 else ""
        if m.get("checkin"):  notas_partes.append(f"{pfx}CheckIn:{m['checkin']}")
        if m.get("checkout"): notas_partes.append(f"{pfx}CheckOut:{m['checkout']}")
        if m.get("nocturno"): notas_partes.append("(nocturno)")
    
    if diff > 0 and estatus != "FALTA": 
        notas_partes.append(f"+{diff}min")

    return estatus, " | ".join(notas_partes)

def _evaluar_turno_guardia(marcas_hoy: list[dict], marcas_manana: list[dict], hora_esperada: time) -> tuple[str, str]:
    """Evalúa el ciclo de Guardia de 24 horas (07:00 AM a 07:00 AM del día siguiente)."""
    if not marcas_hoy and not marcas_manana:
        return "FALTA_GUARDIA", "Ausencia absoluta (Sin marcas de entrada ni salida)"

    # 1. Evaluar registro de salida al día siguiente
    tiempos_manana = []
    for m in marcas_manana:
        if m.get("checkin"): tiempos_manana.append(m["checkin"])
        if m.get("checkout"): tiempos_manana.append(m["checkout"])
        
    if not tiempos_manana:
        return "FALTA_GUARDIA", "Falta registro de salida (Post-Guardia ausente)"
        
    tiempos_manana.sort(key=lambda t: _min_desde_medianoche(t))
    salida_real = tiempos_manana[0] 
    min_salida = _min_desde_medianoche(salida_real) or 0
    
    # Tolerancia inferior estricta: 06:45 AM (405 min)
    if min_salida < 405:
        return "FALTA_GUARDIA", f"Ciclo invalidado: Salida prematura a las {salida_real}"

    # 2. Evaluar registro de entrada (Hoy)
    checkins_hoy = []
    for m in [m for m in marcas_hoy if not m.get("continuacion")]:
        if m.get("checkin"): checkins_hoy.append(m["checkin"])
        
    if not checkins_hoy:
        return "FALTA_GUARDIA", f"Sin registro de entrada. Salida aislada a las {salida_real}"
        
    checkin_ppal = min(checkins_hoy, key=lambda t: _min_desde_medianoche(t))
    min_esperado = _min_desde_medianoche(hora_esperada) or 0
    min_registro = _min_desde_medianoche(checkin_ppal) or 0
    
    # Evaluación de puntualidad de entrada
    if min_registro <= (min_esperado + 30):
        nota_retardo = f" (+{min_registro - min_esperado}min)" if min_registro > (min_esperado + 15) else ""
        return "GUARDIA", f"Ciclo completado | Ent: {checkin_ppal}{nota_retardo} | Sal: {salida_real}"
    else:
        return "FALTA_GUARDIA", f"Llegada fuera de tolerancia ({checkin_ppal})"

# ──────────────────────────────────────────────
# Motor Principal
# ──────────────────────────────────────────────
def procesar_asistencias(datos: dict, reglas: dict) -> pd.DataFrame:
    catalogo     = datos["catalogo"]
    df_rol       = datos.get("rol_guardias", pd.DataFrame())
    df_inc       = datos["incidencias"]
    df_scanner   = datos["scanner"]
    dic_festivos = datos.get("festivos", {})

    tol_retardo = reglas.get("tolerancia_retardo_min", 10)
    tol_falta   = reglas.get("tolerancia_falta_min",   15)
    horarios    = reglas.get("horarios_guardia", {})
    hora_default = time(7, 0)
    
    letra_inicial_mes = reglas.get("letra_inicial_mes", "A")

    ids_catalogo = set(catalogo[catalogo["_activo"]]["ID"].astype(str).str.strip().unique())
    id_map: dict[str, str] = {id_: id_ for id_ in ids_catalogo}

    # [PRESERVADO v3.0] Extracción segura de IDs huérfanos para alertar en la UI
    huerfanos_ids = sorted(
        set(df_scanner["ID_Biometrico"].astype(str).str.strip().unique()) - ids_catalogo
    )
    n_registros_huerfanos = int(
        (~df_scanner["ID_Biometrico"].astype(str).str.strip().isin(ids_catalogo)).sum()
    )
    df_scanner = df_scanner[df_scanner["ID_Biometrico"].astype(str).str.strip().isin(ids_catalogo)].copy()

    s_fechas = pd.to_datetime(df_scanner["Fecha"], errors="coerce").dropna()
    if s_fechas.empty:
        raise ValueError("El reporte del escáner no tiene fechas válidas aplicables al catálogo actual.")

    mes_objetivo = s_fechas.dt.to_period("M").mode()[0]

    fecha_inicio_mes = date(mes_objetivo.year, mes_objetivo.month, 1)
    ultimo_dia = cal.monthrange(mes_objetivo.year, mes_objetivo.month)[1]
    fecha_fin_mes = date(mes_objetivo.year, mes_objetivo.month, ultimo_dia)
    dias_mes = [fecha_inicio_mes + timedelta(days=i) for i in range((fecha_fin_mes - fecha_inicio_mes).days + 1)]

    # [PRESERVADO v3.0] Ventana crítica de ±1 día para procesar el CheckOut del último día del mes
    ventana_ini = fecha_inicio_mes - timedelta(days=1)
    ventana_fin = fecha_fin_mes + timedelta(days=1)
    s_fechas_full = pd.to_datetime(df_scanner["Fecha"], errors="coerce")
    df_scanner = df_scanner[(s_fechas_full.dt.date >= ventana_ini) & (s_fechas_full.dt.date <= ventana_fin)].copy()

    mapa_incidencias = _expandir_incidencias(df_inc)
    cal_hospital     = construir_calendario_hospital(letra_inicial_mes, dias_mes)
    indice_scanner   = _construir_indice_scanner(df_scanner, id_map)

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

        # ── NUEVA REGLA DE NEGOCIO: FILTRO DE ASIGNACIÓN Y EXCLUSIÓN DE HOJA 2_ROL_GUARDIAS ──
        rot_row = df_rol[df_rol["ID"].astype(str).str.strip() == id_emp]
        aplica_ciclo_guardias = False
        grupo_medico = "N/A"
        
        if not rot_row.empty:
            rot_inicial = str(rot_row.iloc[0].get("Rotación", "")).strip().upper()
            if rot_inicial in GUARDIAS_ORDEN:
                grupo_medico = rot_inicial
                aplica_ciclo_guardias = True

        guardia_anterior = None 
        estatus_dia_anterior = None # [INTEGRADO v3.1] Rastreador de estatus para evaluar el día de mañana

        for dia in dias_mes:
            eval_notas        = ""
            notas_adicionales = ""
            turno_servicio    = ""
            es_festivo        = dia in dic_festivos
            
            # Determinación de estado de ciclo dinámico vs Exclusión estricta
            if aplica_ciclo_guardias:
                estado_ciclo = determinar_estado_ciclo(grupo_medico, dia, cal_hospital)
                
                # [INTEGRADO v3.1] IMPLEMENTACIÓN DE LA REGLA DE NEGOCIO FALTANTE
                if estado_ciclo == 'Post-Guardia' and estatus_dia_anterior == 'FALTA_GUARDIA':
                    estado_ciclo = 'Ordinario'
                    notas_adicionales = "⚠️ Descanso revocado por Falta en Guardia previa. "
            else:
                # Totalmente excluido: No entra a ciclos, no tiene post-guardia, es ordinario de Lunes a Domingo
                estado_ciclo = 'Ordinario'

            hora_esp = horarios.get(estado_ciclo, {}).get("hora_inicio", hora_default)

            if (id_emp, dia) in mapa_incidencias:
                inc_info    = mapa_incidencias[(id_emp, dia)]
                estatus_dia = inc_info["tipo"]
                partes = ["Incidencia registrada"]
                if inc_info["destino"]: partes.append(f"Destino: {inc_info['destino']}")
                if inc_info["nota"]:    partes.append(inc_info["nota"])
                eval_notas = " | ".join(partes)
                
                if es_festivo:
                    eval_notas = f"Día Festivo ({dic_festivos[dia]}) | {eval_notas}"
            else:
                if not rot_row.empty:
                    area = str(rot_row.iloc[0].get("Area asignada", "")).strip()
                    irp  = str(rot_row.iloc[0].get("IRP", "")).strip()
                    turno_servicio = " · ".join(filter(None, [area, irp]))

                marcas_hoy = indice_scanner.get((id_emp, dia), [])
                marcas_manana = indice_scanner.get((id_emp, dia + timedelta(days=1)), [])
                
                if estado_ciclo == 'Guardia':
                    estatus_dia, eval_notas = _evaluar_turno_guardia(marcas_hoy, marcas_manana, hora_esp)
                    if es_festivo:
                        eval_notas = f"Guardia en Festivo ({dic_festivos[dia]}) | {eval_notas}"

                elif estado_ciclo == 'Post-Guardia':
                    estatus_dia = "POST_GUARDIA"
                    eval_notas = "Descanso Post-Guardia (Salida procesada en turno anterior)"

                else: # Ordinario (Aquí entra también si se revocó la Post-Guardia o si está excluido)
                    if es_festivo:
                        if marcas_hoy:
                            estatus_dia, eval_notas = _evaluar_turno_ordinario(marcas_hoy, hora_esp, tol_retardo, tol_falta)
                            eval_notas = f"Asistencia en Festivo ({dic_festivos[dia]}) | {eval_notas}"
                        else:
                            estatus_dia = "FESTIVO"
                            eval_notas = dic_festivos[dia]
                    else:
                        estatus_dia, eval_notas = _evaluar_turno_ordinario(marcas_hoy, hora_esp, tol_retardo, tol_falta)

            # Recomponer notas concatenando la alerta de revocación si aplica
            notas = notas_adicionales + eval_notas

            # Alerta de secuencia lógica (Solo aplica a quienes hacen guardias)
            if aplica_ciclo_guardias and guardia_anterior and guardia_anterior == estado_ciclo and estado_ciclo == 'Guardia' and estatus_dia not in ["VACACIONES", "INCAPACIDAD", "PERMISO", "COMISION", "ROTACION", "FESTIVO"]:
                notas = f"ALERTA SECUENCIA: Doble guardia detectada | " + notas
            
            # Guardar histórico para el ciclo del día siguiente
            guardia_anterior = estado_ciclo
            estatus_dia_anterior = estatus_dia

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
                "Rot_Inicial":     grupo_medico,
                "Fecha":           dia,
                "Guardia_Tipo":    estado_ciclo,
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
    
    # [PRESERVADO v3.0] Exponer metadatos de huérfanos sin romper la integridad del DataFrame
    df.attrs["huerfanos_ids"] = huerfanos_ids
    df.attrs["huerfanos_registros"] = n_registros_huerfanos
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