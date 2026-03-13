"""
SGAM - Sistema de Gestión de Asistencias Médicas
Módulo: export.py
Responsabilidad: Generación de reportes Excel institucionales con openpyxl.
"""

import calendar
from datetime import date
from pathlib import Path
from typing import Optional

import pandas as pd
import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage


# ──────────────────────────────────────────────
# Estilos globales
# ──────────────────────────────────────────────
FUENTE_TITULO  = Font(name="Calibri", bold=True, size=14, color="1F3864")
FUENTE_HEADER  = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
FUENTE_CELDA   = Font(name="Calibri", size=9)
FUENTE_FIRMA   = Font(name="Calibri", size=9, italic=True, color="595959")

FILL_HEADER    = PatternFill("solid", fgColor="1F3864")
FILL_SUBHEADER = PatternFill("solid", fgColor="2E75B6")

ALIN_CENTRO    = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIN_IZQ       = Alignment(horizontal="left",   vertical="center", wrap_text=True)

BORDE_FINO = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

BORDE_MEDIO = Border(
    left=Side(style="medium"),
    right=Side(style="medium"),
    top=Side(style="medium"),
    bottom=Side(style="medium"),
)

# Nombres cortos de los meses en español
MESES_ES = {
    1: "Enero", 2: "Febrero",  3: "Marzo",     4: "Abril",
    5: "Mayo",  6: "Junio",    7: "Julio",      8: "Agosto",
    9: "Sept.", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}

DIAS_ES = ["Lu", "Ma", "Mi", "Ju", "Vi", "Sá", "Do"]


# ──────────────────────────────────────────────
# Helpers internos
# ──────────────────────────────────────────────
def _aplicar_borde(celda, borde=BORDE_FINO):
    celda.border = borde


def _celda_estilo(ws, fila, col, valor="", fill=None, fuente=None,
                   alineacion=ALIN_CENTRO, borde=BORDE_FINO):
    """Aplica valor y estilo completo a una celda."""
    c = ws.cell(row=fila, column=col, value=valor)
    if fill:
        c.fill = fill
    if fuente:
        c.font = fuente
    else:
        c.font = FUENTE_CELDA
    c.alignment = alineacion
    c.border = borde
    return c


def _nombre_archivo(nombre_empleado: str, mes: int, anio: int) -> str:
    """Genera el nombre estandarizado del archivo."""
    nombre_limpio = nombre_empleado.replace(" ", "_").replace(",", "")
    mes_str = MESES_ES.get(mes, str(mes))
    return f"SGAM_Reporte_{nombre_limpio}_{mes_str}{anio}.xlsx"


# ──────────────────────────────────────────────
# Construcción de la hoja de reporte
# ──────────────────────────────────────────────
def _insertar_logo(ws, ruta_logo: Optional[str], celda: str = "A1"):
    """Inserta logo institucional si existe el archivo."""
    if not ruta_logo:
        return
    ruta = Path(ruta_logo)
    if not ruta.exists():
        return
    try:
        img = XLImage(str(ruta))
        img.width = 160
        img.height = 60
        ws.add_image(img, celda)
    except Exception:
        pass   # Si falla, continúa sin logo


def _seccion_encabezado(ws, empleado: dict, mes: int, anio: int,
                          ruta_logo: Optional[str]):
    """
    Construye el bloque de encabezado institucional:
    Logo | Nombre del sistema | Datos del empleado
    """
    # Fila 1: Logo + Título del sistema
    ws.row_dimensions[1].height = 50
    ws.merge_cells("A1:C1")
    _insertar_logo(ws, ruta_logo, "A1")

    ws.merge_cells("D1:K1")
    c_titulo = ws.cell(row=1, column=4,
                        value="SGAM – Sistema de Gestión de Asistencias Médicas")
    c_titulo.font = FUENTE_TITULO
    c_titulo.alignment = ALIN_CENTRO
    c_titulo.fill = PatternFill("solid", fgColor="EBF3FB")

    # Fila 2: Nombre del hospital (personalizable)
    ws.merge_cells("A2:K2")
    c_hosp = ws.cell(row=2, column=1,
                      value="HOSPITAL GENERAL – Departamento de Recursos Humanos")
    c_hosp.font = Font(name="Calibri", bold=True, size=11, color="1F3864")
    c_hosp.alignment = ALIN_CENTRO
    c_hosp.fill = PatternFill("solid", fgColor="DEEAF1")

    # Fila 3: vacía como separador visual
    ws.row_dimensions[3].height = 6

    # Fila 4-7: Datos del empleado
    campos = [
        ("Nombre:",         empleado.get("Nombre_Completo", "")),
        ("ID Biométrico:",  empleado.get("ID_Institucional", "")),
        ("Tipo:",           empleado.get("Tipo", "")),
        ("Especialidad:",   empleado.get("Especialidad_Base", "")),
        ("Período:",        f"{MESES_ES.get(mes, mes)} {anio}"),
    ]

    for i, (etiqueta, valor) in enumerate(campos):
        fila = 4 + i
        ws.row_dimensions[fila].height = 16

        c_et = ws.cell(row=fila, column=1, value=etiqueta)
        c_et.font = Font(name="Calibri", bold=True, size=9, color="1F3864")
        c_et.alignment = ALIN_IZQ

        ws.merge_cells(f"B{fila}:E{fila}")
        c_val = ws.cell(row=fila, column=2, value=valor)
        c_val.font = FUENTE_CELDA
        c_val.alignment = ALIN_IZQ
        c_val.border = Border(bottom=Side(style="thin"))


def _seccion_calendario(ws, df_empleado: pd.DataFrame,
                          mes: int, anio: int, fila_inicio: int = 11) -> int:
    """
    Genera la tabla calendario mensual con celdas coloreadas.
    Retorna la siguiente fila disponible después del calendario.
    """
    dias_del_mes = calendar.monthrange(anio, mes)[1]

    # ── Fila de encabezado de días ──────────────────────────────────────
    ws.row_dimensions[fila_inicio].height = 20
    c_mes = ws.cell(row=fila_inicio, column=1,
                     value=f"CALENDARIO – {MESES_ES.get(mes, mes).upper()} {anio}")
    c_mes.font = FUENTE_HEADER
    c_mes.fill = FILL_HEADER
    c_mes.alignment = ALIN_CENTRO
    ws.merge_cells(
        start_row=fila_inicio, start_column=1,
        end_row=fila_inicio, end_column=dias_del_mes + 1
    )

    # ── Fila de números de día ──────────────────────────────────────────
    fila_dias = fila_inicio + 1
    ws.row_dimensions[fila_dias].height = 18
    ws.cell(row=fila_dias, column=1, value="Día →").font = FUENTE_HEADER
    ws.cell(row=fila_dias, column=1).fill = FILL_SUBHEADER
    ws.cell(row=fila_dias, column=1).alignment = ALIN_CENTRO

    for d in range(1, dias_del_mes + 1):
        col = d + 1
        ws.column_dimensions[get_column_letter(col)].width = 5
        c_dia = ws.cell(row=fila_dias, column=col, value=d)
        c_dia.font = FUENTE_HEADER
        c_dia.fill = FILL_SUBHEADER
        c_dia.alignment = ALIN_CENTRO
        c_dia.border = BORDE_FINO

        # Mini día de semana debajo
        fecha_d = date(anio, mes, d)
        nombre_dia = DIAS_ES[fecha_d.weekday()]
        c_dds = ws.cell(row=fila_dias + 1, column=col, value=nombre_dia)
        c_dds.font = Font(name="Calibri", size=7, color="404040")
        c_dds.alignment = ALIN_CENTRO
        c_dds.border = BORDE_FINO

    # ── Fila de estatus coloreado ───────────────────────────────────────
    fila_estatus = fila_dias + 2
    ws.row_dimensions[fila_estatus].height = 22

    # Construir índice {dia: estatus/color}
    mapa_dia = {}
    for _, row in df_empleado.iterrows():
        dia = pd.to_datetime(row["Fecha"]).day
        mapa_dia[dia] = {
            "label": row.get("Label", ""),
            "color": row.get("Color_Hex", "FFFFFF"),
            "notas": row.get("Notas", ""),
        }

    c_label = ws.cell(row=fila_estatus, column=1, value="Estatus")
    c_label.font = Font(name="Calibri", bold=True, size=9, color="FFFFFF")
    c_label.fill = FILL_SUBHEADER
    c_label.alignment = ALIN_CENTRO

    for d in range(1, dias_del_mes + 1):
        col = d + 1
        info = mapa_dia.get(d, {"label": "", "color": "FFFFFF", "notas": ""})
        color_hex = info["color"]
        label_corto = (info["label"] or "")[:3]   # Máx 3 chars para la celda

        c = ws.cell(row=fila_estatus, column=col, value=label_corto)
        c.fill = PatternFill("solid", fgColor=color_hex)
        c.font = Font(name="Calibri", size=7, bold=True, color="1F1F1F")
        c.alignment = ALIN_CENTRO
        c.border = BORDE_FINO

        # Tooltip simulado como comentario
        if info["notas"]:
            from openpyxl.comments import Comment
            comentario = Comment(info["notas"], "SGAM")
            comentario.width = 150
            comentario.height = 60
            c.comment = comentario

    # ── Leyenda de colores ──────────────────────────────────────────────
    fila_leyenda = fila_estatus + 2
    ws.row_dimensions[fila_leyenda].height = 14
    ws.merge_cells(
        start_row=fila_leyenda, start_column=1,
        end_row=fila_leyenda, end_column=dias_del_mes + 1
    )
    c_ley = ws.cell(row=fila_leyenda, column=1,
                     value="LEYENDA: 🟡 Asistencia  🔴 Retardo  🟠 Falta  ⚪ No Laborable  "
                           "🟢 Vacaciones  🔵 Incapacidad")
    c_ley.font = Font(name="Calibri", size=8, italic=True, color="404040")
    c_ley.alignment = ALIN_IZQ

    return fila_leyenda + 2


def _seccion_firma(ws, fila_inicio: int, dias_del_mes: int):
    """Genera espacio para firma de RH."""
    ws.row_dimensions[fila_inicio].height = 40
    ws.row_dimensions[fila_inicio + 1].height = 16
    ws.row_dimensions[fila_inicio + 2].height = 16

    ws.merge_cells(
        start_row=fila_inicio, start_column=1,
        end_row=fila_inicio, end_column=10
    )

    # Línea de firma
    fila_linea = fila_inicio + 1
    ws.merge_cells(f"A{fila_linea}:E{fila_linea}")
    c_lin = ws.cell(row=fila_linea, column=1, value="")
    c_lin.border = Border(bottom=Side(style="medium"))

    ws.merge_cells(f"G{fila_linea}:K{fila_linea}")
    c_lin2 = ws.cell(row=fila_linea, column=7, value="")
    c_lin2.border = Border(bottom=Side(style="medium"))

    # Etiquetas debajo de firma
    fila_et = fila_inicio + 2
    ws.merge_cells(f"A{fila_et}:E{fila_et}")
    c_rh = ws.cell(row=fila_et, column=1,
                    value="Firma y Sello – Recursos Humanos")
    c_rh.font = FUENTE_FIRMA
    c_rh.alignment = ALIN_CENTRO

    ws.merge_cells(f"G{fila_et}:K{fila_et}")
    c_med = ws.cell(row=fila_et, column=7,
                     value="Firma – Médico / Residente")
    c_med.font = FUENTE_FIRMA
    c_med.alignment = ALIN_CENTRO


# ──────────────────────────────────────────────
# Función principal de exportación
# ──────────────────────────────────────────────
def exportar_reporte_empleado(df_empleado: pd.DataFrame,
                               directorio_salida: str,
                               ruta_logo: Optional[str] = None) -> str:
    """
    Genera el reporte Excel institucional para un empleado específico.

    Parámetros:
        df_empleado: DataFrame filtrado para un solo empleado (de core.procesar_asistencias)
        directorio_salida: Carpeta donde se guardará el archivo
        ruta_logo: Ruta al PNG del logo institucional (opcional)

    Retorna:
        Ruta completa del archivo generado
    """
    if df_empleado.empty:
        raise ValueError("No hay datos para exportar. El DataFrame está vacío.")

    # Extraer metadatos del empleado
    primera = df_empleado.iloc[0]
    empleado = {
        "Nombre_Completo":   primera.get("Nombre_Completo", "Desconocido"),
        "ID_Institucional":  primera.get("ID_Institucional", ""),
        "Tipo":              primera.get("Tipo", ""),
        "Especialidad_Base": primera.get("Especialidad_Base", ""),
    }

    fechas = pd.to_datetime(df_empleado["Fecha"])
    mes    = fechas.dt.month.mode()[0]
    anio   = fechas.dt.year.mode()[0]
    dias_del_mes = calendar.monthrange(anio, mes)[1]

    # Filtrar solo el mes actual
    df_mes = df_empleado[
        (fechas.dt.month == mes) & (fechas.dt.year == anio)
    ].copy()

    # Crear workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{MESES_ES.get(mes, mes)} {anio}"

    # Ajustar columna A
    ws.column_dimensions["A"].width = 14

    # ── Secciones del reporte ────────────────────────────────────────────
    _seccion_encabezado(ws, empleado, mes, anio, ruta_logo)
    fila_sig = _seccion_calendario(ws, df_mes, mes, anio, fila_inicio=11)
    _seccion_firma(ws, fila_inicio=fila_sig, dias_del_mes=dias_del_mes)

    # ── Congelar filas de encabezado ─────────────────────────────────────
    ws.freeze_panes = "A12"

    # ── Nombre de archivo y guardado ─────────────────────────────────────
    nombre = _nombre_archivo(empleado["Nombre_Completo"], mes, anio)
    ruta_salida = Path(directorio_salida) / nombre
    Path(directorio_salida).mkdir(parents=True, exist_ok=True)

    wb.save(str(ruta_salida))
    return str(ruta_salida)


def exportar_todos(df_resultado: pd.DataFrame,
                    directorio_salida: str,
                    ruta_logo: Optional[str] = None,
                    callback=None) -> list[str]:
    """
    Exporta un reporte por cada empleado en el DataFrame.

    Parámetros:
        callback: función(progreso: float, nombre: str) para actualizar UI
    """
    ids = df_resultado["ID_Institucional"].unique()
    archivos_generados = []
    total = len(ids)

    for i, id_emp in enumerate(ids):
        df_emp = df_resultado[df_resultado["ID_Institucional"] == id_emp].copy()
        nombre = df_emp.iloc[0].get("Nombre_Completo", id_emp)

        try:
            ruta = exportar_reporte_empleado(df_emp, directorio_salida, ruta_logo)
            archivos_generados.append(ruta)
        except Exception as e:
            print(f"[ERROR] No se pudo exportar {nombre}: {e}")

        if callback:
            callback(progreso=(i + 1) / total, nombre=nombre)

    return archivos_generados
