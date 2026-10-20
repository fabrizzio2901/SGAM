"""Genera una demostración completamente ficticia para SGAM.

Uso: python examples/generate_demo.py [--output demo_output]
"""

import argparse
from datetime import date, timedelta
from pathlib import Path

import pandas as pd


def generate(output: Path) -> tuple[Path, Path]:
    output.mkdir(parents=True, exist_ok=True)
    template = output / "demo_plantilla.xlsx"
    scanner = output / "demo_scanner.xlsx"
    if template.exists() or scanner.exists():
        raise FileExistsError("Los archivos demo ya existen. Elige otra carpeta con --output.")

    staff = pd.DataFrame([
        {"ID": 1001, "Nombre completo": "PERSONA DEMO A", "Estatus": "Activo",
         "Tipo de personal": "Interno", "Periodo ingreso": "A"},
        {"ID": 1002, "Nombre completo": "PERSONA DEMO B", "Estatus": "R1",
         "Tipo de personal": "Residente", "Especialidad": "Medicina Interna"},
    ])
    # Sin grupo A-D: ambos se evalúan en jornada ordinaria.
    duty = pd.DataFrame([{"ID": 1001, "Rotación": ""}, {"ID": 1002, "Rotación": ""}])
    incident_columns = ["ID", "Ausencia Justificada", "Fecha Inicio", "Fecha Termino"]
    vacation = pd.DataFrame([{
        "ID": 1002, "Ausencia Justificada": "vacaciones",
        "Fecha Inicio": date(2026, 4, 13), "Fecha Termino": date(2026, 4, 15),
    }])
    with pd.ExcelWriter(template, engine="openpyxl") as writer:
        staff.to_excel(writer, sheet_name="1_Catalogo_Personal", index=False)
        duty.to_excel(writer, sheet_name="2_Rol_Guardias", index=False)
        pd.DataFrame(columns=incident_columns).to_excel(
            writer, sheet_name="3_Registro_Incidencias", index=False)
        vacation.to_excel(writer, sheet_name="4_Vacaciones", index=False)
        pd.DataFrame(columns=incident_columns).to_excel(
            writer, sheet_name="5_Rotaciones", index=False)
        pd.DataFrame([{"Motivo": "Festivo demo", "Fecha": date(2026, 4, 10)}]).to_excel(
            writer, sheet_name="6_Dias_Festivos", index=False)
        pd.DataFrame([{"ID": "generico", "Dia": "Domingo"}]).to_excel(
            writer, sheet_name="7_Descansos", index=False)

    marks = []
    for offset in range(30):
        day = date(2026, 4, 1) + timedelta(days=offset)
        if day.weekday() >= 5:
            continue
        for person_id in (1001, 1002):
            if person_id == 1002 and date(2026, 4, 13) <= day <= date(2026, 4, 15):
                continue
            # Variaciones ficticias para mostrar asistencia, retardo y falta.
            if person_id == 1001 and day == date(2026, 4, 8):
                continue
            check_in = "07:20" if person_id == 1002 and day == date(2026, 4, 7) else "07:00"
            marks.append({"ID": person_id, "Fecha": day,
                          "Entrada": check_in, "Salida": "16:00"})
    pd.DataFrame(marks).to_excel(scanner, index=False)
    return template, scanner


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=Path("demo_output"))
    args = parser.parse_args()
    try:
        template_path, scanner_path = generate(args.output)
    except FileExistsError as exc:
        parser.exit(1, f"Error: {exc}\n")
    print(f"Plantilla: {template_path}\nEscáner: {scanner_path}")
