"""Acceso de SOLO LECTURA a la base SQLite del sistema.

El agente no puede hacer cambios de ningun tipo. Se aplican tres candados:
1. La base se abre con la URI ``mode=ro`` (el motor SQLite rechaza escrituras).
2. ``PRAGMA query_only = ON`` en la conexion.
3. Solo se aceptan sentencias que empiezan con SELECT o WITH.
"""

from __future__ import annotations

import re
import sqlite3
from pathlib import Path

from ..constants import DEFAULT_DB_PATH


class OperacionNoPermitida(PermissionError):
    """Se intento una operacion que modificaria datos."""


_SOLO_CONSULTA = re.compile(r"^\s*(select|with)\b", re.IGNORECASE)
_PALABRAS_ESCRITURA = re.compile(
    r"\b(insert|update|delete|drop|alter|create|replace|attach|detach|vacuum|reindex|pragma)\b",
    re.IGNORECASE,
)


class LectorSoloLectura:
    def __init__(self, db_path: str | Path | None = None):
        self.db_path = Path(db_path) if db_path else DEFAULT_DB_PATH
        self._conn: sqlite3.Connection | None = None

    # -- conexion ---------------------------------------------------------
    @property
    def disponible(self) -> bool:
        return self.db_path.exists()

    def conectar(self) -> sqlite3.Connection:
        if self._conn is None:
            if not self.db_path.exists():
                raise FileNotFoundError(f"No existe la base de datos: {self.db_path}")
            uri = self.db_path.resolve().as_uri() + "?mode=ro"
            conn = sqlite3.connect(uri, uri=True)
            conn.row_factory = sqlite3.Row
            conn.execute("PRAGMA query_only = ON")
            self._conn = conn
        return self._conn

    def cerrar(self) -> None:
        if self._conn is not None:
            self._conn.close()
            self._conn = None

    def __enter__(self) -> "LectorSoloLectura":
        return self

    def __exit__(self, *exc) -> None:
        self.cerrar()

    # -- consultas ----------------------------------------------------------
    def _validar(self, sql: str) -> None:
        sin_textos = re.sub(r"'[^']*'|\"[^\"]*\"", "''", sql)
        if not _SOLO_CONSULTA.match(sin_textos) or _PALABRAS_ESCRITURA.search(sin_textos) or ";" in sin_textos.strip().rstrip(";"):
            raise OperacionNoPermitida("El agente de Finanzas y Contabilidad solo puede leer datos (SELECT).")

    def filas(self, sql: str, params: tuple | list = ()) -> list[dict]:
        self._validar(sql)
        return [dict(row) for row in self.conectar().execute(sql, params).fetchall()]

    def fila(self, sql: str, params: tuple | list = ()) -> dict | None:
        rows = self.filas(sql, params)
        return rows[0] if rows else None

    def valor(self, sql: str, params: tuple | list = (), default=None):
        row = self.fila(sql, params)
        if not row:
            return default
        value = next(iter(row.values()))
        return default if value is None else value

    def existe_tabla(self, tabla: str) -> bool:
        return bool(self.valor("SELECT COUNT(*) AS n FROM sqlite_master WHERE type = 'table' AND name = ?", (tabla,), 0))

    def tabla(self, tabla: str, where: str = "", params: tuple | list = (), order: str = "") -> list[dict]:
        if not self.existe_tabla(tabla):
            return []
        nombre = '"' + tabla.replace('"', '""') + '"'
        sql = f"SELECT * FROM {nombre}"
        if where:
            sql += f" WHERE {where}"
        if order:
            sql += f" ORDER BY {order}"
        return self.filas(sql, params)

    def ultimo_run(self) -> int | None:
        if not self.existe_tabla("calculation_runs"):
            return None
        return self.valor(
            "SELECT run_id FROM calculation_runs WHERE status = 'COMPLETADO' ORDER BY run_id DESC LIMIT 1"
        )

    def calculado(self, tabla: str, order: str = "") -> list[dict]:
        """Filas de una tabla calculada para la ultima ejecucion completada."""
        run_id = self.ultimo_run()
        if run_id is None:
            return []
        return self.tabla(tabla, "run_id = ?", (run_id,), order)
