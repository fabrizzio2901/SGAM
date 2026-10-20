"""Regresión de conteos compartidos, descansos y exportación de estadísticas."""

import tempfile
import unittest
from pathlib import Path

import openpyxl
import pandas as pd

from core import calcular_metricas
from utils import (
    estadisticas_por_empleado, resumen_por_tipo, resumen_por_especialidad,
    calcular_tendencia_semanal, dias_criticos, exportar_estadisticas_excel,
)


def registros(estados, id_emp="1", tipo="Residente", especialidad="Urgencias"):
    return pd.DataFrame({
        "ID": id_emp,
        "Nombre_Completo": f"Persona {id_emp}",
        "Tipo": tipo,
        "Especialidad": especialidad,
        "Grado": "R1",
        "Periodo_Ingreso": "A",
        "Fecha": pd.date_range("2026-04-01", periods=len(estados)),
        "Estatus": estados,
    })


class MetricasTests(unittest.TestCase):
    def setUp(self):
        self.df = registros([
            "ASISTENCIA", "GUARDIA", "POST_GUARDIA", "RETARDO", "FALTA",
            "FALTA_GUARDIA", "DESCANSO", "VACACIONES", "FESTIVO", "NO_LABORABLE",
        ])

    def test_guardias_y_faltas_guardia_cuentan_en_desempeno(self):
        m = calcular_metricas(self.df)
        self.assertEqual((m["asistencias"], m["faltas"], m["retardos"]), (3, 2, 1))
        self.assertEqual((m["guardias"], m["post_guardias"], m["faltas_guardia"]), (1, 1, 1))
        self.assertEqual(m["dias_evaluados"], 6)
        self.assertEqual((m["pct_asistencia"], m["pct_retardo"], m["pct_falta"]), (50.0, 16.7, 33.3))
        self.assertEqual(m["pct_presencia"], 50.0)

    def test_descansos_e_incidencias_no_penalizan(self):
        base = registros(["GUARDIA", "POST_GUARDIA"])
        neutral = registros([
            "DESCANSO", "NO_LABORABLE", "VACACIONES", "INCAPACIDAD", "PERMISO",
            "COMISION", "ROTACION", "FESTIVO", "OTRO", "SIN_DATOS",
        ])
        m = calcular_metricas(pd.concat([base, neutral], ignore_index=True))
        self.assertEqual(m["dias_evaluados"], 2)
        self.assertEqual(m["pct_asistencia"], 100.0)
        self.assertEqual(m["pct_presencia"], 50.0)
        self.assertEqual(m["descansos"], 1)
        solo = calcular_metricas(neutral)
        self.assertEqual(solo["dias_evaluados"], 0)
        self.assertEqual(solo["pct_asistencia"], 0.0)

    def test_agregados_conservan_totales_y_personas_sin_dias_puntuables(self):
        df = pd.concat([self.df, registros(["DESCANSO"], id_emp="2")], ignore_index=True)
        esperado = calcular_metricas(df)
        individuales = estadisticas_por_empleado(df)
        for clave in ("asistencias", "faltas", "retardos", "descansos", "dias_evaluados"):
            self.assertEqual(individuales[clave].sum(), esperado[clave])
        for tabla in (resumen_por_tipo(df), resumen_por_especialidad(df)):
            fila = tabla.iloc[0]
            self.assertEqual(fila["Empleados"], 2)
            self.assertEqual(fila["Asistencias"], 3)
            self.assertEqual(fila["Faltas"], 2)
            self.assertEqual(fila["Descansos"], 2)
            self.assertEqual(fila["Días Evaluados"], 6)
            self.assertEqual(fila["% Asistencia"], esperado["pct_asistencia"])
        semanas = calcular_tendencia_semanal(df)
        self.assertEqual(semanas["Asistencias"].sum(), esperado["asistencias"])
        self.assertEqual(semanas["Faltas"].sum(), esperado["faltas"])
        self.assertEqual(semanas["Días Evaluados"].sum(), esperado["dias_evaluados"])

    def test_dias_criticos_incluyen_falta_guardia(self):
        criticos = dias_criticos(self.df, top_n=30)
        self.assertEqual(criticos["Faltas"].sum(), 2)
        self.assertEqual(criticos["Retardos"].sum(), 1)
        self.assertEqual(criticos["Total_Incidencias"].sum(), 3)

    def test_solo_descansos_y_sin_especialidad(self):
        df = registros(["DESCANSO"], especialidad="N/A")
        self.assertTrue(resumen_por_especialidad(df).empty)
        self.assertTrue(dias_criticos(df).empty)
        self.assertEqual(resumen_por_tipo(df).iloc[0]["Empleados"], 1)
        self.assertEqual(calcular_tendencia_semanal(df).iloc[0]["Días Evaluados"], 0)

    def test_semanas_en_orden_al_cambiar_de_ano(self):
        df = registros(["GUARDIA", "FALTA_GUARDIA", "DESCANSO"])
        df["Fecha"] = pd.to_datetime(["2025-12-28", "2025-12-29", "2026-01-05"])
        semanas = calcular_tendencia_semanal(df)
        self.assertEqual(semanas["Semana"].tolist(), ["Semana 1", "Semana 2", "Semana 3"])
        self.assertEqual(semanas["Asistencias"].tolist(), [1, 0, 0])
        self.assertEqual(semanas["Faltas"].tolist(), [0, 1, 0])
        self.assertEqual(semanas["Descansos"].tolist(), [0, 0, 1])

    def test_excel_conserva_metricas_y_todos_los_encabezados(self):
        with tempfile.TemporaryDirectory() as temporal:
            archivo = exportar_estadisticas_excel(self.df, temporal)
            self.assertTrue(Path(archivo).exists())
            libro = openpyxl.load_workbook(archivo, data_only=True)
            hoja = libro["Resumen_Individual"]
            encabezados = [celda.value for celda in hoja[2]]
            fila = dict(zip(encabezados, [celda.value for celda in hoja[3]]))
            self.assertEqual(fila["Asist./G./PG"], 3)
            self.assertEqual(fila["Guardias"], 1)
            self.assertEqual(fila["Post-G."], 1)
            self.assertEqual(fila["Faltas totales"], 2)
            self.assertEqual(fila["Falta G."], 1)
            self.assertEqual(fila["Días Eval."], 6)
            self.assertEqual(fila["% Asist."], 50)
            for nombre in ("Por_Tipo", "Por_Especialidad", "Tendencia_Semanal"):
                hoja = libro[nombre]
                self.assertTrue(all(c.value is not None for c in hoja[2]))
            libro.close()


if __name__ == "__main__":
    unittest.main()
