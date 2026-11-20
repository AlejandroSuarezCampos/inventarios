# -*- mode: python ; coding: utf-8 -*-
# Compilar con:  pyinstaller --noconfirm InventarioPC.spec
# Resultado:     dist/InventarioPC.exe  (un único archivo, no necesita Python instalado)
import re
from pathlib import Path

# La versión se toma de inventario/__init__.py y se escribe en los metadatos del .exe
VERSION = re.search(r'__version__\s*=\s*"([^"]+)"', Path("inventario/__init__.py").read_text(encoding="utf-8")).group(1)
partes = [int(p) for p in re.findall(r"\d+", VERSION)[:3]] + [0]
partes = (partes + [0, 0, 0, 0])[:4]
Path("version_info.txt").write_text(f"""VSVersionInfo(
  ffi=FixedFileInfo(filevers={tuple(partes)}, prodvers={tuple(partes)}, mask=0x3f, flags=0x0,
                    OS=0x40004, fileType=0x1, subtype=0x0, date=(0, 0)),
  kids=[
    StringFileInfo([StringTable('0C0A04B0', [
      StringStruct('CompanyName', 'InventarioPC'),
      StringStruct('FileDescription', 'Inventario de Equipos'),
      StringStruct('FileVersion', '{VERSION}'),
      StringStruct('InternalName', 'InventarioPC'),
      StringStruct('OriginalFilename', 'InventarioPC.exe'),
      StringStruct('ProductName', 'Inventario de Equipos'),
      StringStruct('ProductVersion', '{VERSION}')])]),
    VarFileInfo([VarStruct('Translation', [0x0C0A, 1200])])
  ]
)
""", encoding="utf-8")

a = Analysis(
    ["app.py"],
    pathex=[],
    binaries=[],
    datas=[("assets/icono.ico", "assets"), ("assets/icono.png", "assets")],
    hiddenimports=["qrcode", "qrcode.util", "qrcode.base", "qrcode.constants"],
    hookspath=[],
    runtime_hooks=[],
    excludes=["numpy", "pandas", "PIL", "matplotlib", "pytest", "unittest", "pydoc_data", "png"],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="InventarioPC",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    icon="assets/icono.ico",
    version="version_info.txt",
)
