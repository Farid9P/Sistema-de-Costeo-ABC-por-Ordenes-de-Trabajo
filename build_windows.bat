@echo off
setlocal
cd /d "%~dp0"
python build_exe.py
if errorlevel 1 (
    echo.
    echo No se pudo generar el ejecutable. Revise los mensajes anteriores.
    exit /b %errorlevel%
)
echo.
if exist "dist\Sistema_Costos_ABC_Frenosa.exe" (
    echo Ejecutable unico generado: dist\Sistema_Costos_ABC_Frenosa.exe
) else (
    echo No se pudo generar el ejecutable. Revise los mensajes anteriores.
    exit /b 1
)
endlocal
