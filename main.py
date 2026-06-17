"""
SGAM - Sistema de Gestión de Asistencias Médicas
Módulo: main.py
Punto de entrada principal de la aplicación.
"""

import sys
import os

# ── Asegurar que el directorio del proyecto esté en el PATH ─────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

# ── Importación diferida para mostrar errores claros de dependencias ─────
def _verificar_dependencias():
    dependencias = {
        "customtkinter": "customtkinter",
        "pandas":        "pandas",
        "openpyxl":      "openpyxl",
        "matplotlib":    "matplotlib",
        "xlrd":          "xlrd",
        "PIL":           "Pillow",
    }
    faltantes = []
    for nombre_import, nombre_pip in dependencias.items():
        try:
            __import__(nombre_import)
        except ImportError:
            faltantes.append(nombre_pip)

    if faltantes:
        print("❌ Dependencias faltantes. Instálalas con:\n")
        print(f"   pip install {' '.join(faltantes)}\n")
        sys.exit(1)


def main():
    print("=" * 55)
    print("  SGAM – Sistema de Gestión de Asistencias Médicas")
    print("  Versión 1.0.0  |  Iniciando...")
    print("=" * 55)

    _verificar_dependencias()

    from ui import iniciar_app
    iniciar_app()


if __name__ == "__main__":
    main()
