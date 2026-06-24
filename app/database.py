from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Iterable, Sequence

from .constants import DEFAULT_DB_PATH, ROOT_DIR, SHEET_TO_TABLE


SCHEMA_PATH = Path(__file__).with_name("schema.sql")
LEGACY_IMPORTED_TABLE_RENAMES = {
    "periodos": "periodo",
    "pd_actividades": "Personal_directos",
    "costos_indirectos": "costo_otros_recursos(CI)",
    "soporte_mantenimiento": "Soporte_Mantenimiento(CI)",
    "soporte_control_calidad": "Soporte_Control_Calidad(CI)",
    "soporte_ingenieria": "Soporte_Ingenieria(CI)",
}


def quote_identifier(identifier: str) -> str:
    if not identifier or "\x00" in identifier:
        raise ValueError("Identificador SQLite invalido.")
    return '"' + identifier.replace('"', '""') + '"'


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def resolve_db_path(db_path: str | Path | None = None) -> Path:
    return Path(db_path) if db_path else DEFAULT_DB_PATH


def connect(db_path: str | Path | None = None) -> sqlite3.Connection:
    path = resolve_db_path(db_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(path)
    conn.row_factory = sqlite3.Row
    return conn


@contextmanager
def db_session(db_path: str | Path | None = None):
    conn = connect(db_path)
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def initialize_database(db_path: str | Path | None = None) -> Path:
    path = resolve_db_path(db_path)
    with db_session(path) as conn:
        conn.executescript(SCHEMA_PATH.read_text(encoding="utf-8"))
        migrate_legacy_imported_table_names(conn)
    return path


def table_exists(conn: sqlite3.Connection, table_name: str) -> bool:
    row = conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?",
        (table_name,),
    ).fetchone()
    return row is not None


def migrate_legacy_imported_table_names(conn: sqlite3.Connection) -> None:
    for old_name, new_name in LEGACY_IMPORTED_TABLE_RENAMES.items():
        if not table_exists(conn, old_name):
            continue
        if table_exists(conn, new_name):
            old_count = fetch_one_value(conn, f"SELECT COUNT(*) FROM {quote_identifier(old_name)}", default=0)
            new_count = fetch_one_value(conn, f"SELECT COUNT(*) FROM {quote_identifier(new_name)}", default=0)
            if old_count and not new_count:
                old_cols = table_columns(conn, old_name)
                new_cols = table_columns(conn, new_name)
                common_cols = [col for col in new_cols if col in old_cols]
                quoted_cols = ", ".join(common_cols)
                conn.execute(
                    f"""
                    INSERT INTO {quote_identifier(new_name)} ({quoted_cols})
                    SELECT {quoted_cols}
                    FROM {quote_identifier(old_name)}
                    """
                )
            conn.execute(
                "UPDATE source_trace SET table_name = ? WHERE table_name = ?",
                (new_name, old_name),
            )
        conn.execute(f"DROP TABLE {quote_identifier(old_name)}")
    add_missing_columns(
        conn,
        "periodo",
        {"fecha_inicio": "TEXT", "fecha_fin": "TEXT", "origen_periodo": "TEXT"},
    )


def add_missing_columns(conn: sqlite3.Connection, table_name: str, columns: dict[str, str]) -> None:
    if not table_exists(conn, table_name):
        return
    existing = set(table_columns(conn, table_name))
    for column, column_type in columns.items():
        if column not in existing:
            conn.execute(
                f"ALTER TABLE {quote_identifier(table_name)} ADD COLUMN {quote_identifier(column)} {column_type}"
            )


def create_import_batch(conn: sqlite3.Connection, excel_path: str | Path, status: str = "INICIADO", notes: str = "") -> int:
    cur = conn.execute(
        "INSERT INTO import_batches (excel_path, imported_at, status, notes) VALUES (?, ?, ?, ?)",
        (str(excel_path), now_iso(), status, notes),
    )
    return int(cur.lastrowid)


def update_import_batch(conn: sqlite3.Connection, batch_id: int, status: str, notes: str = "") -> None:
    conn.execute(
        "UPDATE import_batches SET status = ?, notes = ? WHERE batch_id = ?",
        (status, notes, batch_id),
    )


def add_import_issue(
    conn: sqlite3.Connection,
    batch_id: int | None,
    severity: str,
    message: str,
    sheet_name: str | None = None,
    row_number: int | None = None,
    column_name: str | None = None,
) -> None:
    conn.execute(
        """
        INSERT INTO import_errors
            (batch_id, severity, sheet_name, row_number, column_name, message)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (batch_id, severity, sheet_name, row_number, column_name, message),
    )


def start_calculation_run(conn: sqlite3.Connection, model_code: str, notes: str = "") -> int:
    cur = conn.execute(
        "INSERT INTO calculation_runs (model_code, started_at, status, notes) VALUES (?, ?, ?, ?)",
        (model_code, now_iso(), "INICIADO", notes),
    )
    return int(cur.lastrowid)


def finish_calculation_run(conn: sqlite3.Connection, run_id: int, status: str, notes: str = "") -> None:
    conn.execute(
        "UPDATE calculation_runs SET finished_at = ?, status = ?, notes = ? WHERE run_id = ?",
        (now_iso(), status, notes, run_id),
    )


def table_columns(conn: sqlite3.Connection, table_name: str) -> list[str]:
    rows = conn.execute(f"PRAGMA table_info({quote_identifier(table_name)})").fetchall()
    return [row["name"] for row in rows]


def clear_imported_tables(conn: sqlite3.Connection) -> None:
    for table in set(SHEET_TO_TABLE.values()):
        conn.execute(f"DELETE FROM {quote_identifier(table)}")
    conn.execute("DELETE FROM source_trace")


def clear_calculated_tables(conn: sqlite3.Connection) -> None:
    tables = [
        "kardex_promedio",
        "cmd_ot",
        "csd_ot",
        "mod_actividad",
        "otros_recursos_ac_detalle",
        "otros_recursos_actividad",
        "soporte_mantenimiento_ac",
        "soporte_control_calidad_ac",
        "soporte_ingenieria_ac",
        "soporte_actividad",
        "actividad_total",
        "produccion_ot_detalle",
        "produccion_ot",
        "soporte_ot_detalle",
        "soporte_ot",
        "hoja_costos_ot",
        "estado_resultados_ot",
        "dashboard_alertas",
    ]
    for table in tables:
        conn.execute(f"DELETE FROM {table}")


def insert_rows(conn: sqlite3.Connection, table_name: str, rows: list[dict], batch_id: int | None = None) -> int:
    if not rows:
        return 0
    available = table_columns(conn, table_name)
    insertable = [col for col in available if col != "batch_id" or batch_id is not None]
    filtered_rows = []
    for row in rows:
        item = {col: row.get(col) for col in insertable if col != "batch_id"}
        if "batch_id" in insertable:
            item["batch_id"] = batch_id
        filtered_rows.append(item)
    columns = list(filtered_rows[0].keys())
    placeholders = ", ".join("?" for _ in columns)
    sql = f"INSERT INTO {quote_identifier(table_name)} ({', '.join(columns)}) VALUES ({placeholders})"
    conn.executemany(sql, [[row.get(col) for col in columns] for row in filtered_rows])
    return len(filtered_rows)


def fetch_all(conn: sqlite3.Connection, table_name: str, order_by: str | None = None) -> list[dict]:
    sql = f"SELECT * FROM {quote_identifier(table_name)}"
    if order_by:
        sql += f" ORDER BY {order_by}"
    return [dict(row) for row in conn.execute(sql).fetchall()]


def fetch_one_value(conn: sqlite3.Connection, sql: str, params: Sequence | None = None, default=None):
    row = conn.execute(sql, params or []).fetchone()
    if row is None:
        return default
    return row[0]


def latest_run_id(conn: sqlite3.Connection) -> int | None:
    return fetch_one_value(
        conn,
        "SELECT run_id FROM calculation_runs WHERE status = 'COMPLETADO' ORDER BY run_id DESC LIMIT 1",
    )


def project_file(path: str | Path) -> Path:
    return ROOT_DIR / path
