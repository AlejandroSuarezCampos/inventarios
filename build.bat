@echo off
REM Compila InventarioPC.exe en Windows (requiere Python 3.10+ instalado SOLO en el equipo que compila).
python -m pip install --upgrade pip
python -m pip install -r requirements-build.txt || goto :error
python -m pytest -q tests || goto :error
python -m PyInstaller --noconfirm --clean InventarioPC.spec || goto :error
echo.
echo Listo: dist\InventarioPC.exe
pause
exit /b 0
:error
echo Ha fallado la compilacion.
pause
exit /b 1
