"""Compila dist/SGAM.exe usando la configuración portable de SGAM.spec.

Uso: python build_exe.py [--clean]
"""

import platform
import subprocess
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
SPEC_FILE = BASE_DIR / "SGAM.spec"


def _asegurar_pyinstaller():
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])


def build():
    if platform.system() != "Windows":
        print("Para generar el ejecutable de Windows, compila este proyecto en Windows.")
        sys.exit(1)
    if not SPEC_FILE.exists():
        print(f"No se encontró {SPEC_FILE}")
        sys.exit(1)

    _asegurar_pyinstaller()
    args = [sys.executable, "-m", "PyInstaller", "--noconfirm"]
    if "--clean" in sys.argv[1:]:
        args.append("--clean")
    args.append(str(SPEC_FILE))

    print("Construyendo SGAM; el empaquetado puede tardar varios minutos.", flush=True)
    result = subprocess.run(args, cwd=BASE_DIR)
    if result.returncode:
        print("La compilación falló. Revisa los mensajes de PyInstaller.")
        sys.exit(result.returncode)
    print(f"Ejecutable generado en: {BASE_DIR / 'dist' / 'SGAM.exe'}")


if __name__ == "__main__":
    build()
