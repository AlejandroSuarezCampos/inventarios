# -*- mode: python ; coding: utf-8 -*-
# Compilar con:  pyinstaller --noconfirm InventarioPC.spec
# Resultado:     dist/InventarioPC.exe  (un único archivo, no necesita Python instalado)

a = Analysis(
    ["app.py"],
    pathex=[],
    binaries=[],
    datas=[("assets/icono.ico", "assets"), ("assets/icono.png", "assets")],
    hiddenimports=[],
    hookspath=[],
    runtime_hooks=[],
    excludes=["numpy", "pandas", "PIL", "matplotlib", "pytest", "unittest", "pydoc_data"],
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
