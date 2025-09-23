@echo off
echo ===============================
echo   INICIANDO APLICATIVO FIRMA
echo ===============================

REM Verifica que Python esté en el PATH
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo ❌ Python no está instalado o no está en el PATH.
    pause
    exit /b
)

REM Ejecutar el programa
python app.py

echo ===============================
echo   EJECUCIÓN FINALIZADA
echo ===============================
pause
