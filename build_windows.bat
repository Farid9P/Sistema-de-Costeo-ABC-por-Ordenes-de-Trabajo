@echo off
setlocal
cd /d "%~dp0"
python -m pip install -r requirements.txt
python build_exe.py
echo.
echo Ejecutable esperado: dist\Sistema_Costos_ABC_Frenosa.exe
endlocal

