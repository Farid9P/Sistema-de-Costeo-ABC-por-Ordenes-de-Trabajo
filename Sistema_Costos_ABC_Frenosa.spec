# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path
import sys


ROOT = Path.cwd()
PYTHON_ROOT = Path(sys.base_prefix)
PYTHON_DLLS = PYTHON_ROOT / "DLLs"
PYTHON_LIB = PYTHON_ROOT / "Lib"
PYTHON_TCL = PYTHON_ROOT / "tcl"

a = Analysis(
    [str(ROOT / "app" / "main.py")],
    pathex=[],
    binaries=[
        (str(PYTHON_DLLS / "_tkinter.pyd"), "."),
        (str(PYTHON_DLLS / "tcl86t.dll"), "."),
        (str(PYTHON_DLLS / "tk86t.dll"), "."),
    ],
    datas=[
        (str(ROOT / "app" / "schema.sql"), "app"),
        (str(ROOT / "images_models"), "images_models"),
        (str(PYTHON_LIB / "tkinter"), "tkinter"),
        (str(PYTHON_TCL / "tcl8.6"), "tcl/tcl8.6"),
        (str(PYTHON_TCL / "tk8.6"), "tcl/tk8.6"),
        (str(PYTHON_TCL / "tcl8"), "tcl/tcl8"),
        (str(PYTHON_TCL / "dde1.4"), "tcl/dde1.4"),
        (str(PYTHON_TCL / "reg1.3"), "tcl/reg1.3"),
    ],
    hiddenimports=[
        "tkinter",
        "_tkinter",
        "tkinter.ttk",
        "tkinter.filedialog",
        "tkinter.messagebox",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[str(ROOT / "app" / "pyinstaller_tk_runtime.py")],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="Sistema_Costos_ABC_Frenosa",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
