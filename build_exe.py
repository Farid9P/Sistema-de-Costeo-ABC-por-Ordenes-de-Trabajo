from __future__ import annotations

import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent
EXE_NAME = "Sistema_Costos_ABC_Frenosa"


def main() -> None:
    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--windowed",
        "--name",
        EXE_NAME,
        "--add-data",
        f"{ROOT / 'app' / 'schema.sql'};app",
        str(ROOT / "app" / "main.py"),
    ]
    print("Compilando ejecutable offline...")
    print(" ".join(str(part) for part in cmd))
    subprocess.check_call(cmd, cwd=ROOT)
    exe_path = ROOT / "dist" / f"{EXE_NAME}.exe"
    if not exe_path.exists():
        raise SystemExit("PyInstaller termino, pero no se encontro el ejecutable esperado.")
    print(f"Ejecutable generado: {exe_path}")


if __name__ == "__main__":
    main()

