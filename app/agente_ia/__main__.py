"""Consola del agente.

    python -m app.agente_ia "¿cual es el producto con mas ventas?"
    python -m app.agente_ia --propuestas
    python -m app.agente_ia            (modo interactivo)
"""

from __future__ import annotations

import argparse
import sys

from ..constants import DEFAULT_DB_PATH
from .agente import AVISO_SOLO_LECTURA, AgenteFinanzas


def main() -> None:
    parser = argparse.ArgumentParser(description="Agente IA de Finanzas y Contabilidad (solo lectura)")
    parser.add_argument("pregunta", nargs="*", help="Pregunta en lenguaje natural")
    parser.add_argument("--db", default=str(DEFAULT_DB_PATH), help="Base SQLite del sistema")
    parser.add_argument("--propuestas", action="store_true", help="Mostrar propuestas de mejora")
    parser.add_argument("--solo-interpretar", action="store_true", help="Mostrar solo la interpretacion")
    args = parser.parse_args()
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    agente = AgenteFinanzas(args.db)
    if args.propuestas:
        print(agente.propuestas_texto())
        return
    if args.pregunta:
        print(agente.consultar(" ".join(args.pregunta), responder=not args.solo_interpretar).texto())
        return
    print(agente.nombre + "\n" + AVISO_SOLO_LECTURA + "\nEscriba su pregunta (vacio para salir).")
    while True:
        try:
            pregunta = input("\n> ").strip()
        except (EOFError, KeyboardInterrupt):
            break
        if not pregunta:
            break
        print(agente.consultar(pregunta).texto())


if __name__ == "__main__":
    main()
