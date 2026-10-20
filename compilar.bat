@echo off
REM Compila SGAM en un .exe. Doble clic para ejecutar.
cd /d "%~dp0"
if not exist venv (
    echo Creando entorno virtual...
    python -m venv venv
)
call venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
python build_exe.py
echo.
echo Si todo salio bien, el ejecutable esta en: dist\SGAM.exe
pause
