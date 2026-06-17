"""
SGAM - Generador de archivos de prueba v2.1
Genera plantillas de ejemplo con IDs que coinciden con el escáner real del hospital.

IDs del escáner real (Diciembre 2025):
  Residentes MR-25: 2025047, 2025071, 2025072, 2025073, 2025074,
                    2025075, 2025076, 2025077, 2025078, 2025079, 2025080, 2025081
  Residentes MR-23: 1032353
  Residentes MR-21: 2021001
  Otros:            902373, 12345, 1111, 160260, ...
"""

import random
from datetime import date, time
from pathlib import Path
import pandas as pd

random.seed(42)


def _excel_template(ruta, catalogo_rows, rol_rows,
                    inc_rows=None, vac_rows=None, rot_rows=None):
    """Escribe una plantilla SGAM con las 5 hojas."""
    Path(ruta).parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(ruta, engine="openpyxl") as w:
        pd.DataFrame(catalogo_rows).to_excel(
            w, sheet_name="1_Catalogo_Personal", index=False)
        pd.DataFrame(rol_rows).to_excel(
            w, sheet_name="2_Rol_Guardias", index=False)
        pd.DataFrame(inc_rows  or []).to_excel(
            w, sheet_name="3_Registro_Incidencias", index=False)
        pd.DataFrame(vac_rows  or []).to_excel(
            w, sheet_name="4_Vacaciones", index=False)
        pd.DataFrame(rot_rows  or []).to_excel(
            w, sheet_name="5_Rotaciones", index=False)


def generar_plantilla_internos(ruta="data/Plantilla_Internos_Ejemplo.xlsx"):
    """
    Internos de ejemplo — IDs tomados del escáner real del hospital.
    Internos no tienen especialidad.
    """
    catalogo = [
        {
            "ID": 160260,
            "Nombre completo": "INTERNO EJEMPLO UNO",
            "Estatus": "activo",
            "Tipo de personal": "interno",
            "Universidad": "BUAP",
            "Especialidad": "",
            "Subespecialidad": "",
            "Alta Especialidad": "",
            "Periodo ingreso": "B",
        },
        {
            "ID": 405682,
            "Nombre completo": "INTERNO EJEMPLO DOS",
            "Estatus": "activo",
            "Tipo de personal": "interno",
            "Universidad": "UPAEP",
            "Especialidad": "",
            "Subespecialidad": "",
            "Alta Especialidad": "",
            "Periodo ingreso": "A",
        },
        {
            "ID": 406866,
            "Nombre completo": "INTERNO EJEMPLO TRES",
            "Estatus": "activo",
            "Tipo de personal": "interno",
            "Universidad": "UNAM",
            "Especialidad": "",
            "Subespecialidad": "",
            "Alta Especialidad": "",
            "Periodo ingreso": "B",
        },
        {
            "ID": 414159,
            "Nombre completo": "INTERNO EJEMPLO CUATRO",
            "Estatus": "no activo",   # inactivo — no se procesa
            "Tipo de personal": "interno",
            "Universidad": "UAM",
            "Especialidad": "",
            "Subespecialidad": "",
            "Alta Especialidad": "",
            "Periodo ingreso": "A",
        },
    ]
    guardias = [
        {"ID": 160260, "Area asignada": "Urgencias",        "Rotación": "A", "Observacion": ""},
        {"ID": 405682, "Area asignada": "Hospitalización",  "Rotación": "C", "Observacion": ""},
        {"ID": 406866, "Area asignada": "Consulta Externa", "Rotación": "B", "Observacion": ""},
    ]
    incidencias = [
        {
            "ID": 405682,
            "Ausencia Justificada": "incapacidad",
            "Fecha Inicio":  "08/12/2025",
            "Fecha Termino": "10/12/2025",
            "Observacion":   "Incapacidad por enfermedad general",
        },
    ]
    vacaciones = [
        {
            "ID": 160260,
            "Ausencia Justificada": "vacaciones",
            "Fecha Inicio":  "22/12/2025",
            "Fecha Termino": "26/12/2025",
            "Observacion":   "Vacaciones fin de año",
        },
    ]
    rotaciones = []

    _excel_template(ruta, catalogo, guardias, incidencias, vacaciones, rotaciones)
    print(f"✅ Internos: {ruta}")
    return ruta


def generar_plantilla_residentes(ruta="data/Plantilla_Residentes_Ejemplo.xlsx"):
    """
    Residentes de ejemplo — IDs que coinciden con el escáner real (Diciembre 2025).
    IDs MR-25: 2025047 … 2025081
    IDs MR-23: 1032353
    IDs MR-21: 2021001
    """
    catalogo = [
        {
            "ID": 2021001,
            "Nombre completo": "MR-21 ALFREDO ISMAEL MEDRANO",
            "Estatus": "r4",            # R4 = activo automático
            "Tipo de personal": "residente",
            "Universidad": "BUAP",
            "Especialidad": "Medicina Interna",
            "Subespecialidad": "No aplica",
            "Alta Especialidad": "No aplica",
            "Periodo ingreso": "A",
        },
        {
            "ID": 1032353,
            "Nombre completo": "MR-23 CARLOS ERIK BARRIOS",
            "Estatus": "r2",
            "Tipo de personal": "residente",
            "Universidad": "UPAEP",
            "Especialidad": "Pediatría",
            "Subespecialidad": "No aplica",
            "Alta Especialidad": "No aplica",
            "Periodo ingreso": "A",
        },
        {
            "ID": 2025047,
            "Nombre completo": "MR-25 LUIS HUMBERTO TORRES",
            "Estatus": "r1",
            "Tipo de personal": "residente",
            "Universidad": "UNAM",
            "Especialidad": "Urgencias",
            "Subespecialidad": "No aplica",
            "Alta Especialidad": "No aplica",
            "Periodo ingreso": "B",
        },
        {
            "ID": 2025071,
            "Nombre completo": "MR-25 ALEJANDRO NOE SALINAS",
            "Estatus": "r1",
            "Tipo de personal": "residente",
            "Universidad": "IPN",
            "Especialidad": "Urgencias",
            "Subespecialidad": "No aplica",
            "Alta Especialidad": "No aplica",
            "Periodo ingreso": "B",
        },
        {
            "ID": 2025072,
            "Nombre completo": "MR-25 DAVID ALFREDO GONZALEZ",
            "Estatus": "r1",
            "Tipo de personal": "residente",
            "Universidad": "UAM",
            "Especialidad": "Ginecología",
            "Subespecialidad": "No aplica",
            "Alta Especialidad": "No aplica",
            "Periodo ingreso": "B",
        },
        {
            "ID": 2025073,
            "Nombre completo": "MR-25 MARIA FERNANDA LOPEZ",
            "Estatus": "r1",
            "Tipo de personal": "residente",
            "Universidad": "UV",
            "Especialidad": "Pediatría",
            "Subespecialidad": "No aplica",
            "Alta Especialidad": "No aplica",
            "Periodo ingreso": "B",
        },
        {
            "ID": 2025074,
            "Nombre completo": "MR-25 LUIS ARMANDO MARTINEZ",
            "Estatus": "r1",
            "Tipo de personal": "residente",
            "Universidad": "BUAP",
            "Especialidad": "Cirugía General",
            "Subespecialidad": "No aplica",
            "Alta Especialidad": "No aplica",
            "Periodo ingreso": "B",
        },
    ]
    guardias = [
        {"ID": 2021001,  "Area asignada": "Med. Interna",   "Rotación": "C", "Observacion": ""},
        {"ID": 1032353,  "Area asignada": "Pediatría",       "Rotación": "A", "Observacion": ""},
        {"ID": 2025047,  "Area asignada": "Urgencias",       "Rotación": "B", "Observacion": ""},
        {"ID": 2025071,  "Area asignada": "Urgencias",       "Rotación": "D", "Observacion": ""},
        {"ID": 2025072,  "Area asignada": "Ginecología",     "Rotación": "A", "Observacion": ""},
        {"ID": 2025073,  "Area asignada": "Pediatría",       "Rotación": "C", "Observacion": ""},
        {"ID": 2025074,  "Area asignada": "Cirugía",         "Rotación": "B", "Observacion": ""},
    ]
    incidencias = [
        {
            "ID": 2025071,
            "Ausencia Justificada": "permisos",
            "Fecha Inicio":  "06/12/2025",
            "Fecha Termino": "06/12/2025",
            "Observacion":   "Permiso por asuntos personales",
        },
    ]
    vacaciones = [
        {
            "ID": 2021001,
            "Ausencia Justificada": "vacaciones",
            "Fecha Inicio":  "22/12/2025",
            "Fecha Termino": "26/12/2025",
            "Observacion":   "Vacaciones navideñas",
        },
    ]
    rotaciones = [
        {
            "ID": 2025074,
            "Ausencia Justificada": "Rotacion externa",
            "Fecha Inicio":  "15/12/2025",
            "Fecha Termino": "19/12/2025",
            "Destino":       "Hospital Regional Sur – Cirugía General",
            "Observaciones": "Rotación programada en convenio institucional",
        },
    ]

    _excel_template(ruta, catalogo, guardias, incidencias, vacaciones, rotaciones)
    print(f"✅ Residentes: {ruta}")
    return ruta


if __name__ == "__main__":
    generar_plantilla_internos()
    generar_plantilla_residentes()
    print("\n📂 Plantillas de ejemplo listas en data/")
    print("   Para el escáner, usa el archivo real del hospital (.xls) directamente.")
    print("   El sistema lo procesa con cargar_reporte_scanner_hospital().")
