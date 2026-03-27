"""
SGAM - Sistema de Gestión de Asistencias Médicas
Módulo: core.py
Responsabilidad: Motor de cruce de datos y asignación de estatus por día/empleado.

Mejoras v1.1:
  - Soporte para dobles turnos (dos guardias el mismo día)
  - Turnos nocturnos: checkout en el día siguiente se asocia al día de entrada
  - Perfil enriquecido: Tipo y Especialidad incluidos en todos los registros
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
    "vacacion":   "VACACIONES",
    "vacaciones": "VACACIONES",
    "incapacidad": "INCAPACIDAD",
    "permiso":    "PERMISO",
    "comision":   "COMISION",
    "comisión":   "COMISION",
}

# Hora de corte para detectar turnos nocturnos:
# Un checkout cuya hora sea MENOR que este umbral se considera del día siguiente.
UMBRAL_TURNO_NOCTURNO_HORA = 6   # Si checkout < 06:00 → se interpreta como madrugada


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
        id_inst  = str(row["ID_Institucional"]).strip()
        tipo_raw = str(row.get("Tipo_Incidencia", "OTRO")).lower().strip()
        tipo     = INCIDENCIA_TIPO_MAP.get(tipo_raw, "OTRO")
        f_ini    = row["Fecha_Inicio"]
        f_fin    = row["Fecha_Fin"]

        if not f_ini or not f_fin:
            continue

        if isinstance(f_ini, datetime):
            f_ini = f_ini.date()
        if isinstance(f_fin, datetime):
            f_fin = f_fin.date()

        delta = (f_fin - f_ini).days + 1
        for i in range(delta):
            dia = f_ini + timedelta(days=i)
            mapa[(id_inst, dia)] = tipo

    return mapa


def _construir_dict_guardias(df_guardias: pd.DataFrame) -> dict[tuple, list[dict]]:
    """
    Construye un dict {(id_institucional, fecha): [lista de turnos del día]}.

    Cada turno es un dict con las claves del rol: TIPO, Servicio_Cubierto.
    Permite manejar DOBLES TURNOS: el mismo empleado puede tener dos entradas
    para el mismo día (p.ej. turno matutino + turno vespertino o guardia extra).
    """
    guardias: dict[tuple, list[dict]] = {}
    for _, row in df_guardias.iterrows():
        id_inst = str(row["ID_Institucional"]).strip()
        fecha   = row["Fecha_Guardia"]
        if isinstance(fecha, datetime):
            fecha = fecha.date()
        if not fecha:
            continue

        clave = (id_inst, fecha)
        turno = {
            "tipo_turno": str(row.get("TIPO", "")).strip(),
            "servicio":   str(row.get("Servicio_Cubierto", "")).strip(),
        }
        guardias.setdefault(clave, []).append(turno)

    return guardias


def _construir_indice_scanner(df_scanner: pd.DataFrame,
                               catalogo: pd.DataFrame) -> dict[tuple, list[dict]]:
    """
    Construye dict {(id_institucional, fecha): [lista de marcas del día]}.

    Maneja DOBLES MARCAS (dos checkins el mismo día) y TURNOS NOCTURNOS:
    si el checkout tiene hora < UMBRAL_TURNO_NOCTURNO_HORA se asume que
    pertenece al turno iniciado el día anterior, por lo que se registra
    también como checkout del día previo.

    Cada marca es un dict: {'checkin': time|None, 'checkout': time|None, 'nocturno': bool}
    """
    # Mapa ID biométrico → ID institucional
    bio_to_id: dict[str, str] = {}
    for _, row in catalogo.iterrows():
        bio     = str(row["ID_Biometrico_SIRA"]).strip()
        id_inst = str(row.get("ID_Institucional", row["ID_Biometrico_SIRA"])).strip()
        bio_to_id[bio] = id_inst

    # Acumulador: clave → lista de marcas
    indice: dict[tuple, list[dict]] = {}

    for _, row in df_scanner.iterrows():
        bio  = str(row["ID_Biometrico"]).strip()
        fecha = row["Fecha"]
        if isinstance(fecha, datetime):
            fecha = fecha.date()

        id_inst  = bio_to_id.get(bio, bio)
        checkin  = row.get("Hora_CheckIn")
        checkout = row.get("Hora_CheckOut")

        checkin  = checkin  if (checkin  is not None and not _is_na(checkin))  else None
        checkout = checkout if (checkout is not None and not _is_na(checkout)) else None

        # ── Detectar turno nocturno ──────────────────────────────────────
        # Si el checkout es en madrugada (<06:00) se asume que el médico
        # entró hoy y salió "mañana". Guardamos el checkout también en el
        # día siguiente para cruce correcto.
        es_nocturno = False
        if checkout is not None:
            hora_salida = checkout.hour if hasattr(checkout, "hour") else 0
            if hora_salida < UMBRAL_TURNO_NOCTURNO_HORA:
                es_nocturno = True

        marca = {
            "checkin":   checkin,
            "checkout":  checkout,
            "nocturno":  es_nocturno,
        }

        # Agregar marca al día de checkin
        clave_hoy = (id_inst, fecha)
        indice.setdefault(clave_hoy, []).append(marca)

        # Si es turno nocturno, el checkout también se "ve" en el día siguiente
        # para que ese día no cuente como Falta si hay guardia programada.
        if es_nocturno:
            dia_siguiente = fecha + timedelta(days=1)
            clave_sig = (id_inst, dia_siguiente)
            marca_sig = {
                "checkin":  None,          # No hay checkin en el día siguiente
                "checkout": checkout,      # Solo hay checkout (continuación nocturna)
                "nocturno": True,
                "continuacion": True,      # Marca informativa: es continuación
            }
            indice.setdefault(clave_sig, []).append(marca_sig)

    return indice


def _is_na(valor) -> bool:
    """Wrapper seguro para pd.isna que no falla con objetos time."""
    try:
        return pd.isna(valor)
    except (TypeError, ValueError):
        return False


def _minutos_desde_medianoche(t: Optional[time]) -> Optional[int]:
    """Convierte un objeto time a minutos desde medianoche."""
    if t is None:
        return None
    return t.hour * 60 + t.minute


def _evaluar_turno(marcas: list[dict],
                    hora_entrada_esperada: time,
                    tolerancia_min: int) -> tuple[str, str]:
    """
    Evalúa una lista de marcas del día para un empleado con turno programado.

    Soporta DOBLES MARCAS: si hay dos checkins, se toma el mejor (más temprano).
    Maneja continuaciones nocturnas: si la única marca es una continuación
    (checkout de madrugada, sin checkin propio), se considera ASISTENCIA por
    continuidad de guardia.

    Retorna: (estatus: str, notas: str)
    """
    if not marcas:
        return "FALTA", "Sin registro biométrico"

    # Separar marcas reales de continuaciones nocturnas
    marcas_reales       = [m for m in marcas if not m.get("continuacion")]
    marcas_continuacion = [m for m in marcas if m.get("continuacion")]

    # ── Solo hay continuación nocturna (sin checkin propio) ──────────────
    if not marcas_reales and marcas_continuacion:
        checkout = marcas_continuacion[0].get("checkout")
        nota = f"Continuación turno nocturno | CheckOut: {checkout}" if checkout else "Continuación turno nocturno"
        return "ASISTENCIA", nota

    # ── Hay una o más marcas reales ──────────────────────────────────────
    # En caso de doble turno, tomamos el checkin más temprano para evaluar puntualidad
    checkins_validos = [
        m["checkin"] for m in marcas_reales
        if m.get("checkin") is not None
    ]

    if not checkins_validos:
        # Hay registros pero sin checkin (solo checkout). Contar como asistencia.
        notas_partes = []
        for m in marcas_reales:
            if m.get("checkout"):
                notas_partes.append(f"CheckOut: {m['checkout']}")
        return "ASISTENCIA", " | ".join(notas_partes) or "Sin CheckIn registrado"

    # Ordenar checkins y tomar el más temprano (mejor caso para el médico)
    checkin_principal = min(checkins_validos,
                             key=lambda t: _minutos_desde_medianoche(t))

    minutos_real     = _minutos_desde_medianoche(checkin_principal)
    minutos_esperado = _minutos_desde_medianoche(hora_entrada_esperada)
    diferencia       = minutos_real - minutos_esperado

    estatus = "ASISTENCIA" if diferencia <= tolerancia_min else "RETARDO"

    # Construir nota descriptiva con todas las marcas del día
    notas_partes = []
    for idx, m in enumerate(marcas_reales, start=1):
        prefijo = f"T{idx}" if len(marcas_reales) > 1 else ""
        if m.get("checkin"):
            notas_partes.append(f"{prefijo}CheckIn:{m['checkin']}")
        if m.get("checkout"):
            notas_partes.append(f"{prefijo}CheckOut:{m['checkout']}")
        if m.get("nocturno"):
            notas_partes.append("(nocturno)")

    return estatus, " | ".join(notas_partes)


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
            Grado, Fecha, Estatus, Color_Hex, Label, Emoji, Notas,
            Turno_Tipo, Turno_Servicio
    """
    catalogo       = datos["catalogo"]
    df_incidencias = datos["incidencias"]
    df_guardias    = datos["guardias"]
    df_scanner     = datos["scanner"]

    tolerancia_min   = reglas.get("tolerancia_retardo_min", 10)
    horarios_guardia = reglas.get("horarios_guardia", {})
    # hora_entrada_default se usará solo si no hay horario específico por tipo
    hora_entrada_default = time(8, 0, 0)

    # ── Pre-calcular estructuras de búsqueda rápida ──────────────────────
    mapa_incidencias = _expandir_incidencias(df_incidencias)
    dict_guardias    = _construir_dict_guardias(df_guardias)  # {clave: [turnos]}
    indice_scanner   = _construir_indice_scanner(df_scanner, catalogo)  # {clave: [marcas]}

    # ── Determinar rango de fechas del mes ───────────────────────────────
    todas_fechas = (
        [f for (_, f) in dict_guardias.keys()] +
        [f for (_, f) in indice_scanner.keys()]
    )
    if not todas_fechas:
        raise ValueError("No se encontraron fechas en guardias ni en el escáner.")

    fecha_min        = min(todas_fechas)
    fecha_max        = max(todas_fechas)
    fecha_inicio_mes = fecha_min.replace(day=1)
    if fecha_max.month == 12:
        fecha_fin_mes = date(fecha_max.year + 1, 1, 1) - timedelta(days=1)
    else:
        fecha_fin_mes = date(fecha_max.year, fecha_max.month + 1, 1) - timedelta(days=1)

    dias_mes = [
        fecha_inicio_mes + timedelta(days=i)
        for i in range((fecha_fin_mes - fecha_inicio_mes).days + 1)
    ]

    # ── Procesamiento por empleado y día ─────────────────────────────────
    registros = []

    for _, empleado in catalogo.iterrows():
        id_bio       = str(empleado["ID_Biometrico_SIRA"]).strip()
        nombre       = str(empleado.get("Nombre_Completo", "")).strip()
        tipo         = str(empleado.get("Tipo", "")).strip()
        especialidad = str(empleado.get("Especialidad_Base", "")).strip()
        grado        = str(empleado.get("Grado", "")).strip()
        periodo      = str(empleado.get("Periodo_Ingreso", "")).strip()
        foto_ruta    = str(empleado.get("Foto_Ruta", "")).strip()
        estatus_emp  = str(empleado.get("Estatus", "")).strip().upper()

        # Ignorar empleados inactivos
        if estatus_emp not in ["ACTIVO", "ACTIVE", "1", "SI", "SÍ", ""]:
            continue

        id_inst = id_bio   # ID_Biometrico_SIRA es la clave de cruce

        for dia in dias_mes:
            notas           = ""
            turno_tipo      = ""
            turno_servicio  = ""

            # ── PRIORIDAD 1: Incidencia registrada ───────────────────────
            if (id_inst, dia) in mapa_incidencias:
                estatus_dia = mapa_incidencias[(id_inst, dia)]
                # Buscar notas de la incidencia original
                notas = "Incidencia registrada"

            # ── PRIORIDAD 2: Sin turno en Rol de Guardias ────────────────
            elif (id_inst, dia) not in dict_guardias:
                estatus_dia = "NO_LABORABLE"

            # ── PRIORIDAD 3: Evaluar asistencia ─────────────────────────
            else:
                turnos_del_dia = dict_guardias[(id_inst, dia)]

                # Metadatos del turno
                turno_tipo     = " + ".join(
                    t["tipo_turno"] for t in turnos_del_dia if t["tipo_turno"]
                )
                turno_servicio = " / ".join(
                    t["servicio"] for t in turnos_del_dia if t["servicio"]
                )

                # Determinar hora de entrada esperada según tipo de guardia
                # Si hay múltiples turnos, usar el primero con horario definido
                tipo_turno_ppal = turnos_del_dia[0].get("tipo_turno", "").upper()
                horario_turno   = horarios_guardia.get(tipo_turno_ppal, {})
                hora_esperada   = horario_turno.get("hora_inicio", hora_entrada_default)

                marcas = indice_scanner.get((id_inst, dia), [])
                estatus_dia, notas = _evaluar_turno(marcas, hora_esperada, tolerancia_min)

            color_info = COLOR_MAP.get(estatus_dia, COLOR_MAP["SIN_DATOS"])

            registros.append({
                "ID_Institucional":  id_inst,
                "Nombre_Completo":   nombre,
                "Tipo":              tipo,
                "Especialidad_Base": especialidad,
                "Grado":             grado,
                "Periodo_Ingreso":   periodo,
                "Foto_Ruta":         foto_ruta,
                "Fecha":             dia,
                "Estatus":           estatus_dia,
                "Color_Hex":         color_info["hex"],
                "Label":             color_info["label"],
                "Emoji":             color_info["emoji"],
                "Notas":             notas,
                "Turno_Tipo":        turno_tipo,
                "Turno_Servicio":    turno_servicio,
            })

    if not registros:
        raise ValueError(
            "No se generaron registros. Verifique que el catálogo tenga empleados activos."
        )

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
        mes: Period
    """
    laborables = df_resultado[df_resultado["Estatus"] != "NO_LABORABLE"]
    conteos    = laborables["Estatus"].value_counts().to_dict()
    total      = len(laborables)

    porcentajes = {
        k: round((v / total * 100), 1) if total > 0 else 0
        for k, v in conteos.items()
    }

    return {
        "conteos":               conteos,
        "porcentajes":           porcentajes,
        "total_dias_laborables": total,
        "empleados":             df_resultado["ID_Institucional"].nunique(),
        "mes":                   (
            df_resultado["Fecha"].dt.to_period("M").iloc[0]
            if len(df_resultado) else None
        ),
    }


# ──────────────────────────────────────────────
# Filtros de consulta
# ──────────────────────────────────────────────
def filtrar_por_empleado(df_resultado: pd.DataFrame,
                          termino: str) -> pd.DataFrame:
    """Filtra el DataFrame por ID o nombre (búsqueda parcial, sin distinción de mayúsculas)."""
    termino = termino.strip().lower()
    mask = (
        df_resultado["ID_Institucional"].str.lower().str.contains(termino, na=False) |
        df_resultado["Nombre_Completo"].str.lower().str.contains(termino, na=False)
    )
    return df_resultado[mask]


def filtrar_por_tipo(df_resultado: pd.DataFrame,
                      tipo: str) -> pd.DataFrame:
    """
    Filtra por Tipo de personal (ej. 'Residente', 'Interno', 'Médico Adscrito').
    La búsqueda es parcial e insensible a mayúsculas.
    """
    tipo = tipo.strip().lower()
    mask = df_resultado["Tipo"].str.lower().str.contains(tipo, na=False)
    return df_resultado[mask]


def filtrar_por_especialidad(df_resultado: pd.DataFrame,
                              especialidad: str) -> pd.DataFrame:
    """
    Filtra por Especialidad (búsqueda parcial, insensible a mayúsculas).
    """
    esp = especialidad.strip().lower()
    mask = df_resultado["Especialidad_Base"].str.lower().str.contains(esp, na=False)
    return df_resultado[mask]


def obtener_tipos_unicos(df_resultado: pd.DataFrame) -> list[str]:
    """Retorna la lista de tipos de personal presentes en los datos."""
    return sorted(df_resultado["Tipo"].dropna().unique().tolist())


def obtener_especialidades_unicas(df_resultado: pd.DataFrame) -> list[str]:
    """Retorna la lista de especialidades presentes en los datos."""
    return sorted(df_resultado["Especialidad_Base"].dropna().unique().tolist())
