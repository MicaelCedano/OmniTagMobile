@echo off
cd /d "%~dp0"
echo Iniciando aplicacion...
".venv\Scripts\python.exe" "etiqueta_iphone_2025.py"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Ocurrio un error al ejecutar la aplicacion.
    echo Por favor, revisa el mensaje de arriba.
    pause
)
