# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path
import sys


ROOT = Path.cwd()
PYTHON_ROOT = Path(sys.base_prefix)
PYTHON_DLLS = PYTHON_ROOT / "DLLs"
PYTHON_LIB = PYTHON_ROOT / "Lib"
PYTHON_TCL = PYTHON_ROOT / "tcl"
SOURCE_MODULES = [
    "app/importers.py",
    "app/costing_mod.py",
    "app/costing_abc.py",
    "app/reports.py",
    "app/documents.py",
    "app/tests_runner.py",
]

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
        (str(ROOT / "app" / "agente_ia" / "preguntas_base.txt"), "app/agente_ia"),
        (str(ROOT / "images_models"), "images_models"),
        (str(PYTHON_LIB / "tkinter"), "tkinter"),
        (str(PYTHON_TCL / "tcl8.6"), "tcl/tcl8.6"),
        (str(PYTHON_TCL / "tk8.6"), "tcl/tk8.6"),
        (str(PYTHON_TCL / "tcl8"), "tcl/tcl8"),
        (str(PYTHON_TCL / "dde1.4"), "tcl/dde1.4"),
        (str(PYTHON_TCL / "reg1.3"), "tcl/reg1.3"),
        *[(str(ROOT / source), "app_source/app") for source in SOURCE_MODULES],
    ],
    hiddenimports=[
        "tkinter",
        "_tkinter",
        "tkinter.ttk",
        "tkinter.filedialog",
        "tkinter.messagebox",
        "matplotlib.backends.backend_tkagg",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[str(ROOT / "app" / "pyinstaller_tk_runtime.py")],
    excludes=[
        "pandas",
        "lxml.isoschematron",
    ],
    noarchive=False,
    optimize=1,
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
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
