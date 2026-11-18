from __future__ import annotations

import sys
from pathlib import Path


if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    from app.ui.main_window import run_app
else:
    from .ui.main_window import run_app


def main() -> None:
    run_app()


if __name__ == "__main__":
    main()

