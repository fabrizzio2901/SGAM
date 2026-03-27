"""
SGAM - Generador de Archivos de Ejemplo
Crea Estructura_Maestra_Hospital2.xlsx y Reporte_Scanner_Ejemplo.xlsx
con datos ficticios para pruebas.
"""

import pandas as pd
from datetime import date, timedelta
import random
import openpyxl
from pathlib import Path


def generar_maestro(ruta_salida: str = "data/Estructura_Maestra_Hospital2.xlsx"):
    Path(ruta_salida).parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:

        # ── Hoja 1: Reglas ──────────────────────────────────────────
        df_reglas = pd.DataFrame({
            "Regla de Negocio": [
                "Tolerancia Retardo",
                "Tolerancia Salida Anticipada",
                "Horario Guardia A",
                "Horario Guardia B",
                "Horario Guardia C",
            ],
            "Descripción": [
                "Minutos permitidos de retraso antes de marcar retardo: 10 minutos",
                "Minutos antes del fin de turno que se permite salir: 10 minutos",
                "Turno A Matutino: 08:00 - 15:00",
                "Turno B Vespertino: 15:00 - 21:00",
                "Turno C Nocturno: 21:00 - 08:00",
            ],
        })
        df_reglas.to_excel(writer, sheet_name="1_Reglas_ID", index=False)

        # ── Hoja 2: Catálogo Personal (campos nuevos: Periodo_Ingreso, Foto_Ruta) ──
        personal = []
        datos = [
            # (nombre, tipo, especialidad, grado, periodo)
            ("Dra. Laura Mendoza Torres",   "Médico Adscrito", "Urgencias",        "Especialista", "A"),
            ("Dr. Carlos Ríos Alvarado",    "Residente",       "Pediatría",        "R2",           "A"),
            ("Dra. Sofía Hernández Paz",    "Residente",       "Ginecología",      "R1",           "B"),
            ("Dr. Javier Morales Cruz",     "Médico Adscrito", "Medicina Interna", "Especialista", "A"),
            ("Dr. Andrés López Gutiérrez",  "Interno",         "",                 "Intern.",      "B"),
            ("Dra. Patricia Vega Sánchez",  "Médico Adscrito", "Cardiología",      "Especialista", "A"),
            ("Dr. Fernando Castillo Ruiz",  "Residente",       "Traumatología",    "R3",           "B"),
            ("Dra. Isabel Martínez Flores", "Interno",         "",                 "Intern.",      "A"),
        ]
        for i, (nombre, tipo, esp, grado, periodo) in enumerate(datos, start=1001):
            personal.append({
                "ID_Biometrico_SIRA": str(i),
                "Nombre_Completo":    nombre,
                "Tipo":               tipo,
                "Especialidad_Base":  esp,          # Vacío para Internos
                "Grado":              grado,
                "Universidad":        random.choice(["UNAM", "IPN", "UAM", "UV"]),
                "Estatus":            "Activo",
                "Vigencia":           "2025-12-31",
                "Periodo_Ingreso":    periodo,      # NUEVO: A = Primavera, B = Otoño
                "Foto_Ruta":          "",           # NUEVO: ruta a imagen (vacío en ejemplo)
            })

        pd.DataFrame(personal).to_excel(writer, sheet_name="2_Catalogo_Personal", index=False)

        # ── Hoja 3: Incidencias (con Notas_Motivo y rotación con destino) ──
        incidencias = [
            {
                "ID_Institucional":  "1001",
                "Tipo_Incidencia":   "Vacaciones",
                "Fecha_Inicio":      date(2025, 1, 13),
                "Fecha_Fin":         date(2025, 1, 17),
                "Destino_o_Servicio": "",
                "Notas_Motivo":      "Período vacacional aprobado por jefatura.",
            },
            {
                "ID_Institucional":  "1003",
                "Tipo_Incidencia":   "Incapacidad",
                "Fecha_Inicio":      date(2025, 1, 8),
                "Fecha_Fin":         date(2025, 1, 10),
                "Destino_o_Servicio": "Médico tratante externo",
                "Notas_Motivo":      "Incapacidad por síndrome gripal. Certificado IMSS adjunto.",
            },
            {
                "ID_Institucional":  "1005",
                "Tipo_Incidencia":   "Rotación",           # NUEVO tipo — requiere destino
                "Fecha_Inicio":      date(2025, 1, 20),
                "Fecha_Fin":         date(2025, 1, 24),
                "Destino_o_Servicio": "Hospital Regional Sur – Servicio de Cirugía",
                "Notas_Motivo":      "Rotación programada en convenio institucional.",
            },
            {
                "ID_Institucional":  "1002",
                "Tipo_Incidencia":   "Permiso",
                "Fecha_Inicio":      date(2025, 1, 6),
                "Fecha_Fin":         date(2025, 1, 6),
                "Destino_o_Servicio": "",
                "Notas_Motivo":      "Permiso por asuntos personales. Aprobado verbalmente.",
            },
        ]
        pd.DataFrame(incidencias).to_excel(
            writer, sheet_name="3_Registro_Incidencias", index=False)

        # ── Hoja 4: Rol de Guardias (tipos A, B, C) ──────────────────
        guardias = []
        ids_personal = [str(i) for i in range(1001, 1009)]
        servicios = ["Urgencias", "Hospitalización", "Consulta Externa", "UCI"]

        for dia in range(1, 32):
            try:
                fecha = date(2025, 1, dia)
            except ValueError:
                break
            for id_emp in ids_personal:
                if fecha.weekday() < 5:
                    # Turno A (matutino) entre semana para todos
                    guardias.append({
                        "Fecha_Guardia":    fecha,
                        "ID_Institucional": id_emp,
                        "Servicio_Cubierto": random.choice(servicios),
                        "TIPO": "A",
                    })
                else:
                    # Fines de semana: rotación B y C
                    if random.random() < 0.3:
                        tipo_g = random.choice(["B", "C"])
                        guardias.append({
                            "Fecha_Guardia":    fecha,
                            "ID_Institucional": id_emp,
                            "Servicio_Cubierto": random.choice(servicios),
                            "TIPO": tipo_g,
                        })

        pd.DataFrame(guardias).to_excel(writer, sheet_name="4_Rol_Guardias", index=False)

    print(f"✅ Archivo maestro generado: {ruta_salida}")
    return ruta_salida


def generar_scanner(ruta_salida: str = "data/Reporte_Scanner_Enero2025.xlsx"):
    """Genera un reporte del escáner biométrico con datos ficticios."""
    Path(ruta_salida).parent.mkdir(parents=True, exist_ok=True)

    from datetime import time

    registros = []
    ids_personal = [str(i) for i in range(1001, 1009)]

    for dia in range(1, 32):
        try:
            fecha = date(2025, 1, dia)
        except ValueError:
            break

        if fecha.weekday() >= 5:
            continue   # Sin registros en fines de semana (simplificado)

        for id_emp in ids_personal:
            # Simular ~15% de faltas, ~10% de retardos
            prob = random.random()

            if prob < 0.15:
                # Falta (sin registro)
                continue
            elif prob < 0.25:
                # Retardo (llega tarde)
                minutos_tarde = random.randint(11, 45)
                hora_entrada = time(8, minutos_tarde)
            else:
                # Asistencia puntual (±9 min de tolerancia)
                minutos = random.randint(0, 9)
                hora_entrada = time(7, 55 + minutos if 55 + minutos < 60 else minutos,
                                    random.randint(0, 59))

            hora_salida = time(15, random.randint(0, 30))

            registros.append({
                "ID_Biometrico": id_emp,
                "Fecha":         fecha,
                "Hora_CheckIn":  hora_entrada.strftime("%H:%M:%S"),
                "Hora_CheckOut": hora_salida.strftime("%H:%M:%S"),
            })

    df_scanner = pd.DataFrame(registros)
    df_scanner.to_excel(ruta_salida, index=False)
    print(f"✅ Archivo escáner generado: {ruta_salida} ({len(df_scanner)} registros)")
    return ruta_salida


if __name__ == "__main__":
    generar_maestro()
    generar_scanner()
    print("\n📂 Archivos de ejemplo listos en la carpeta 'data/'")
