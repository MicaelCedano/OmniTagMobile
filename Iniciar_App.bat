@echo off
cd /d "%~dp0"
echo Iniciando OmniTag Mobile...
".venv\Scripts\python.exe" "omnitag_mobile.py"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Ocurrio un error al ejecutar la aplicacion.
    echo Por favor, revisa el mensaje de arriba.
    pause
)
