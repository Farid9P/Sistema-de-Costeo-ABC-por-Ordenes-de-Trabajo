from __future__ import annotations

import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parent
EXE_NAME = "Sistema_Costos_ABC_Frenosa"
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


def verify_onefile_contents(exe_path: Path) -> None:
    result = subprocess.run(
        [
            sys.executable,
            "-m",
            "PyInstaller.utils.cliutils.archive_viewer",
            str(exe_path),
            "--list",
        ],
        check=True,
        capture_output=True,
        text=True,
    )
    content = result.stdout.lower().replace("\\\\", "\\")
    required_entries = [
        "_tkinter.pyd",
        "tcl86t.dll",
        "tk86t.dll",
        "tkinter\\__init__.py",
        "tcl\\tcl8.6\\init.tcl",
        "images_models\\modelo1.png",
        "images_models\\modelo2.png",
        "app_source\\app\\importers.py",
        "app_source\\app\\costing_abc.py",
        "app\\agente_ia\\preguntas_base.txt",
    ]
    missing = [entry for entry in required_entries if entry.lower() not in content]
    if missing:
        raise SystemExit(
            "El ejecutable se genero incompleto. Faltan recursos: "
            + ", ".join(missing)
        )


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    temp_root = Path(tempfile.gettempdir()) / "costos_abc_frenosa_build" / stamp
    workpath = temp_root / "build"
    specpath = temp_root / "spec"
    distpath = ROOT / "dist"
    workpath.mkdir(parents=True, exist_ok=True)
    specpath.mkdir(parents=True, exist_ok=True)
    distpath.mkdir(parents=True, exist_ok=True)
    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--windowed",
        "--noupx",
        "--optimize",
        "1",
        "--name",
        EXE_NAME,
        "--runtime-hook",
        str(ROOT / "app" / "pyinstaller_tk_runtime.py"),
        "--workpath",
        str(workpath),
        "--specpath",
        str(specpath),
        "--distpath",
        str(distpath),
        "--add-data",
        f"{ROOT / 'app' / 'schema.sql'};app",
        "--add-data",
        f"{ROOT / 'app' / 'agente_ia' / 'preguntas_base.txt'};app/agente_ia",
        "--add-data",
        f"{ROOT / 'images_models'};images_models",
        "--add-data",
        f"{PYTHON_LIB / 'tkinter'};tkinter",
        "--add-data",
        f"{PYTHON_TCL / 'tcl8.6'};tcl/tcl8.6",
        "--add-data",
        f"{PYTHON_TCL / 'tk8.6'};tcl/tk8.6",
        "--add-data",
        f"{PYTHON_TCL / 'tcl8'};tcl/tcl8",
        "--add-data",
        f"{PYTHON_TCL / 'dde1.4'};tcl/dde1.4",
        "--add-data",
        f"{PYTHON_TCL / 'reg1.3'};tcl/reg1.3",
        "--add-binary",
        f"{PYTHON_DLLS / '_tkinter.pyd'};.",
        "--add-binary",
        f"{PYTHON_DLLS / 'tcl86t.dll'};.",
        "--add-binary",
        f"{PYTHON_DLLS / 'tk86t.dll'};.",
        "--hidden-import",
        "tkinter",
        "--hidden-import",
        "_tkinter",
        "--hidden-import",
        "tkinter.ttk",
        "--hidden-import",
        "tkinter.filedialog",
        "--hidden-import",
        "tkinter.messagebox",
        "--hidden-import",
        "matplotlib.backends.backend_tkagg",
        "--exclude-module",
        "pandas",
        "--exclude-module",
        "lxml.isoschematron",
    ]
    for source_module in SOURCE_MODULES:
        cmd.extend(["--add-data", f"{ROOT / source_module};app_source/app"])
    cmd.append(str(ROOT / "app" / "main.py"))
    print("Compilando ejecutable offline...")
    print(" ".join(str(part) for part in cmd))
    subprocess.check_call(cmd, cwd=ROOT)
    exe_path = distpath / f"{EXE_NAME}.exe"
    if not exe_path.exists():
        raise SystemExit("PyInstaller termino, pero no se encontro el ejecutable esperado.")
    verify_onefile_contents(exe_path)
    print(f"Ejecutable unico generado: {exe_path}")


if __name__ == "__main__":
    main()
