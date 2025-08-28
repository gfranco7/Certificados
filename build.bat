@echo off
echo ==========================================
echo   🚀 Compilando Certificados App (PyInstaller)
echo ==========================================

:: Nombre del ejecutable
set APP_NAME=Certificados

:: Cierra procesos abiertos del exe anterior
echo [1/4] Cerrando procesos de %APP_NAME%.exe ...
taskkill /F /IM %APP_NAME%.exe >nul 2>&1

:: Borra carpetas previas de build y dist
echo [2/4] Eliminando carpetas anteriores ...
rmdir /S /Q build dist __pycache__ >nul 2>&1

:: Ejecuta PyInstaller
echo [3/4] Compilando con PyInstaller ...
pyinstaller --onefile --noconsole ^
 --add-data "templates;templates" ^
 --add-data "static;static" ^
 app.py --name %APP_NAME%

:: Verifica si se creó el exe
if exist dist\%APP_NAME%.exe (
    echo [4/4]  Compilación exitosa: dist\%APP_NAME%.exe
) else (
    echo [4/4]  Error: No se generó el ejecutable.
)

pause
