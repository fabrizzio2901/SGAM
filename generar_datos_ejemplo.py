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
                "Jornada Estándar",
                "Guardia 24h",
            ],
            "Descripción": [
                "Minutos permitidos de retraso antes de marcar retardo: 10 minutos",
                "Minutos antes del fin de turno que se permite salir: 10 minutos",
                "Jornada regular de 8:00 a 15:00 horas",
                "Guardia especial de 08:00 a 08:00 del día siguiente",
            ],
        })
        df_reglas.to_excel(writer, sheet_name="1_Reglas_ID", index=False)

        # ── Hoja 2: Catálogo Personal ────────────────────────────────
        personal = []
        nombres = [
            "Dra. Laura Mendoza Torres",    "Dr. Carlos Ríos Alvarado",
            "Dra. Sofía Hernández Paz",     "Dr. Javier Morales Cruz",
            "Dr. Andrés López Gutiérrez",   "Dra. Patricia Vega Sánchez",
            "Dr. Fernando Castillo Ruiz",   "Dra. Isabel Martínez Flores",
        ]
        tipos = ["Médico Adscrito", "Residente", "Residente", "Médico Adscrito",
                 "Interno",         "Médico Adscrito", "Residente", "Interno"]
        especialidades = [
            "Urgencias", "Pediatría", "Ginecología", "Medicina Interna",
            "Cirugía General", "Cardiología", "Traumatología", "Urgencias"
        ]

        for i, (nombre, tipo, esp) in enumerate(zip(nombres, tipos, especialidades), start=1001):
            personal.append({
                "ID_Biometrico_SIRA": str(i),
                "Nombre_Completo":    nombre,
                "Tipo":               tipo,
                "Especialidad_Base":  esp,
                "Grado":              "Especialista" if tipo == "Médico Adscrito" else "R2" if tipo == "Residente" else "Intern.",
                "Universidad":        random.choice(["UNAM", "IPN", "UAM", "UV"]),
                "Estatus":            "Activo",
                "Vigencia":           "2025-12-31",
            })

        df_catalogo = pd.DataFrame(personal)
        df_catalogo.to_excel(writer, sheet_name="2_Catalogo_Personal", index=False)

        # ── Hoja 3: Incidencias ──────────────────────────────────────
        incidencias = [
            {
                "ID_Institucional": "1001",
                "Tipo_Incidencia":  "Vacaciones",
                "Fecha_Inicio":     date(2025, 1, 13),
                "Fecha_Fin":        date(2025, 1, 17),
                "Destino_o_Servicio": "Descanso anual",
                "Notas_Motivo":     "Período vacacional aprobado",
            },
            {
                "ID_Institucional": "1003",
                "Tipo_Incidencia":  "Incapacidad",
                "Fecha_Inicio":     date(2025, 1, 8),
                "Fecha_Fin":        date(2025, 1, 10),
                "Destino_o_Servicio": "Médico tratante",
                "Notas_Motivo":     "Incapacidad por enfermedad general",
            },
            {
                "ID_Institucional": "1005",
                "Tipo_Incidencia":  "Comisión",
                "Fecha_Inicio":     date(2025, 1, 20),
                "Fecha_Fin":        date(2025, 1, 21),
                "Destino_o_Servicio": "Congreso Nacional de Cirugía",
                "Notas_Motivo":     "Representación institucional",
            },
        ]
        df_incidencias = pd.DataFrame(incidencias)
        df_incidencias.to_excel(writer, sheet_name="3_Registro_Incidencias", index=False)

        # ── Hoja 4: Rol de Guardias ──────────────────────────────────
        guardias = []
        ids_personal = [str(i) for i in range(1001, 1009)]
        servicios = ["Urgencias", "Hospitalización", "Consulta Externa", "UCI"]

        # Enero 2025 — días laborables (lun-vie, más algunos fines de semana)
        for dia in range(1, 32):
            try:
                fecha = date(2025, 1, dia)
            except ValueError:
                break

            # Todos tienen turno entre semana; guardia rotativa fines de semana
            for id_emp in ids_personal:
                if fecha.weekday() < 5:   # Lunes a Viernes
                    guardias.append({
                        "Fecha_Guardia":    fecha,
                        "ID_Institucional": id_emp,
                        "Servicio_Cubierto": random.choice(servicios),
                        "TIPO": "Turno Normal",
                    })
                elif random.random() < 0.3:   # 30% probabilidad guardia fin de semana
                    guardias.append({
                        "Fecha_Guardia":    fecha,
                        "ID_Institucional": id_emp,
                        "Servicio_Cubierto": random.choice(servicios),
                        "TIPO": "Guardia 24h",
                    })

        df_guardias = pd.DataFrame(guardias)
        df_guardias.to_excel(writer, sheet_name="4_Rol_Guardias", index=False)

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
