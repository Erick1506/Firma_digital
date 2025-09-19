@echo off
echo ===============================
echo   LIMPIANDO CARPETAS ANTIGUAS
echo ===============================

rmdir /s /q build
rmdir /s /q dist
del app.spec

echo ===============================
echo   COMPILANDO EJECUTABLE .EXE
echo ===============================

pyinstaller --onefile --noconsole --icon=icono.png ^
--hidden-import docx2pdf ^
--hidden-import comtypes ^
--hidden-import win32com ^
app.py

echo.
echo ✅ Compilación completada.
echo ===============================
echo Copiando archivos necesarios a la carpeta "dist"...
echo ===============================

REM Crear estructura de carpetas en dist
mkdir dist\firma

REM Copiar logo, certificado e icono
xcopy /E /I /Y firma dist\firma
copy /Y WolfangAlbertoLatorreMartinez.pfx dist\
copy /Y icono.png dist\

echo.
echo Archivos copiados con exito!
echo El ejecutable y los archivos estan en la carpeta "dist".
echo ===============================

pause
