"""Pruebas del cargador y el motor: descansos exclusivos de un ID."""

import gc
import tempfile
import unittest
from datetime import date, time
from pathlib import Path

import pandas as pd

from core import COLOR_MAP, calcular_metricas, procesar_asistencias
from ingestion import cargar_plantilla, extraer_reglas


class DescansosPorIdTests(unittest.TestCase):
    def procesar(self, descansos, marcas_extra=(), grupo=""):
        catalogo = pd.DataFrame({
            "ID": ["123", "456"],
            "Nombre completo": ["Persona 123", "Persona 456"],
            "Estatus": ["Activo", "Activo"],
            "Tipo de personal": ["Interno", "Interno"],
            "Periodo ingreso": ["A", "A"],
        })
        rol = pd.DataFrame({"ID": ["123", "456"], "Rotación": [grupo, ""]})
        incidencias = pd.DataFrame(columns=["ID", "Ausencia Justificada", "Fecha Inicio", "Fecha Termino"])
        with tempfile.TemporaryDirectory() as temporal:
            plantilla = Path(temporal) / "plantilla.xlsx"
            with pd.ExcelWriter(plantilla, engine="openpyxl") as writer:
                catalogo.to_excel(writer, sheet_name="1_Catalogo_Personal", index=False)
                rol.to_excel(writer, sheet_name="2_Rol_Guardias", index=False)
                for hoja in ("3_Registro_Incidencias", "4_Vacaciones", "5_Rotaciones"):
                    incidencias.to_excel(writer, sheet_name=hoja, index=False)
                pd.DataFrame(columns=["Motivo", "Fecha"]).to_excel(writer, sheet_name="6_Dias_Festivos", index=False)
                pd.DataFrame(descansos, columns=["ID", "Dia", "Fecha(Opcional)"]).to_excel(writer, sheet_name="7_Descansos", index=False)
            datos = cargar_plantilla(str(plantilla))
            marcas = [
                {"ID_Biometrico": id_emp, "Fecha": date(2026, 4, 1), "Hora_CheckIn": time(7), "Hora_CheckOut": time(17)}
                for id_emp in ("123", "456")
            ]
            datos["scanner"] = pd.DataFrame(marcas + list(marcas_extra))
            resultado = procesar_asistencias(datos, extraer_reglas(pd.DataFrame()))
            # Libera los lectores de Excel antes de retirar el archivo temporal en Windows.
            gc.collect()
            return resultado

    def dia(self, resultado, id_emp, fecha):
        return resultado[(resultado["ID"] == id_emp) & (resultado["Fecha"] == pd.Timestamp(fecha))].iloc[0]

    def test_descanso_semanal_solo_para_id_asignado_y_normalizado(self):
        resultado = self.procesar([("00123", "Domingo", None)])
        for fecha in ("2026-04-05", "2026-04-12", "2026-04-19", "2026-04-26"):
            asignado = self.dia(resultado, "123", fecha)
            self.assertEqual(asignado["Estatus"], "DESCANSO")
            self.assertEqual(asignado["Label"], "Descanso")
            self.assertEqual(asignado["Color_Hex"], COLOR_MAP["DESCANSO"]["hex"])
            self.assertEqual(self.dia(resultado, "456", fecha)["Estatus"], "FALTA")
        self.assertEqual(calcular_metricas(resultado[resultado["ID"] == "123"])["descansos"], 4)
        self.assertEqual(calcular_metricas(resultado[resultado["ID"] == "456"])["descansos"], 0)

    def test_fecha_puntual_solo_en_fecha_e_id_indicados(self):
        resultado = self.procesar([("123.0", "Otro", date(2026, 4, 8))])
        self.assertEqual(self.dia(resultado, "123", "2026-04-08")["Estatus"], "DESCANSO")
        self.assertEqual(self.dia(resultado, "456", "2026-04-08")["Estatus"], "FALTA")
        self.assertEqual(self.dia(resultado, "123", "2026-04-07")["Estatus"], "FALTA")
        self.assertEqual(self.dia(resultado, "123", "2026-04-09")["Estatus"], "FALTA")
        self.assertEqual((resultado["Estatus"] == "DESCANSO").sum(), 1)

    def test_marca_en_descanso_individual_se_evalua_como_asistencia(self):
        resultado = self.procesar([("123", "Domingo", None)], marcas_extra=[{
            "ID_Biometrico": "123", "Fecha": date(2026, 4, 5),
            "Hora_CheckIn": time(7), "Hora_CheckOut": time(17),
        }])
        asignado = self.dia(resultado, "123", "2026-04-05")
        self.assertEqual(asignado["Estatus"], "ASISTENCIA")
        self.assertIn("Asistencia en Descanso", asignado["Notas"])
        self.assertEqual(self.dia(resultado, "123", "2026-04-12")["Estatus"], "DESCANSO")
        self.assertEqual(self.dia(resultado, "456", "2026-04-05")["Estatus"], "FALTA")

    def test_generico_se_acumula_sin_extender_el_descanso_individual(self):
        resultado = self.procesar([("generico", "Sábado", None), ("123", "Domingo", None)])
        self.assertEqual(self.dia(resultado, "123", "2026-04-04")["Estatus"], "DESCANSO")
        self.assertEqual(self.dia(resultado, "456", "2026-04-04")["Estatus"], "DESCANSO")
        self.assertEqual(self.dia(resultado, "456", "2026-04-05")["Estatus"], "FALTA")
        self.assertEqual(calcular_metricas(resultado[resultado["ID"] == "123"])["descansos"], 8)
        self.assertEqual(calcular_metricas(resultado[resultado["ID"] == "456"])["descansos"], 4)

    def test_guardia_conserva_prioridad_sobre_descanso_individual(self):
        resultado = self.procesar([("123", "Otro", date(2026, 4, 5))], grupo="A")
        asignado = self.dia(resultado, "123", "2026-04-05")
        self.assertEqual(asignado["Guardia_Tipo"], "Guardia")
        self.assertEqual(asignado["Estatus"], "FALTA_GUARDIA")
        self.assertEqual(self.dia(resultado, "456", "2026-04-05")["Estatus"], "FALTA")


if __name__ == "__main__":
    unittest.main()
