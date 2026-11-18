from __future__ import annotations

import os
import sys
from pathlib import Path


def _resource_root() -> Path:
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)  # type: ignore[attr-defined]
    return Path(__file__).resolve().parent.parent


base = _resource_root()
tcl_library = base / "tcl" / "tcl8.6"
tk_library = base / "tcl" / "tk8.6"

if tcl_library.exists():
    os.environ["TCL_LIBRARY"] = str(tcl_library)
if tk_library.exists():
    os.environ["TK_LIBRARY"] = str(tk_library)
