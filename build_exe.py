"""
SGAM - Script de construcción del ejecutable .exe
Usa PyInstaller para empaquetar la aplicación.
Ejecutar con: python build_exe.py
"""

import subprocess
import sys
from pathlib import Path

APP_NAME    = "SGAM"
MAIN_SCRIPT = "main.py"
ICON_PATH   = "assets/icono.ico"   # Crear o dejar comentado si no existe
BASE_DIR    = Path(__file__).parent


def build():
    args = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--onefile",                      # Un solo ejecutable
        "--windowed",                     # Sin ventana de consola
        f"--name={APP_NAME}",
        "--add-data=assets;assets",       # Incluir carpeta de recursos
        "--add-data=data;data",           # Incluir datos de ejemplo (opcional)
        "--hidden-import=customtkinter",
        "--hidden-import=PIL._tkinter_finder",
        "--hidden-import=openpyxl",
        "--hidden-import=pandas",
        "--hidden-import=matplotlib",
        "--collect-data=customtkinter",
        "--collect-data=matplotlib",
    ]

    # Agregar ícono si existe
    if Path(ICON_PATH).exists():
        args.append(f"--icon={ICON_PATH}")

    args.append(MAIN_SCRIPT)

    print("🔨 Construyendo ejecutable SGAM...")
    print(f"   Comando: {' '.join(args)}\n")

    result = subprocess.run(args, cwd=BASE_DIR)

    if result.returncode == 0:
        exe_path = BASE_DIR / "dist" / f"{APP_NAME}.exe"
        print(f"\n✅ Ejecutable generado en: {exe_path}")
    else:
        print("\n❌ Error durante la construcción. Revise la salida de PyInstaller.")
        sys.exit(1)


if __name__ == "__main__":
    build()
