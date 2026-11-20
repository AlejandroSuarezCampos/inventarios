@echo off
REM Compila InventarioPC.exe (y el instalador si está Inno Setup 6) en Windows.
REM Requiere Python 3.10+ instalado SOLO en el equipo que compila.
python -m pip install --upgrade pip
python -m pip install -r requirements-build.txt || goto :error
python -m pytest -q tests || goto :error
python -m PyInstaller --noconfirm --clean InventarioPC.spec || goto :error
for /f "tokens=2 delims==\" " %%v in ('findstr /c:"__version__" inventario\__init__.py') do set VERSION=%%~v
set ISCC="%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if exist %ISCC% (
  %ISCC% /DMyAppVersion=%VERSION% installer\InventarioPC.iss || goto :error
  echo Instalador: dist\InventarioPC-Setup-%VERSION%.exe
) else (
  echo Inno Setup 6 no esta instalado: se omite el instalador.
)
echo.
echo Listo: dist\InventarioPC.exe
pause
exit /b 0
:error
echo Ha fallado la compilacion.
pause
exit /b 1
