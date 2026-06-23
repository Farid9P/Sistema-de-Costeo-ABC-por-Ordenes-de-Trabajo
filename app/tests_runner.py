from __future__ import annotations

import argparse
from pathlib import Path
from typing import Callable

from .constants import DEFAULT_DB_PATH, DEFAULT_EXCEL_PATH, OUTPUT_DOCUMENTS, OUTPUT_REPORTS, REPORTS_DIR
from .costing_abc import run_model_1_calculation
from .database import db_session, initialize_database, latest_run_id, quote_identifier
from .documents import generate_all_documents
from .importers import import_excel_to_sqlite, validate_workbook
from .model_availability import evaluate_models
from .reports import generate_all_reports


def _count(conn, table: str, run_id: int | None = None) -> int:
    quoted = quote_identifier(table)
    if run_id is None:
        return conn.execute(f"SELECT COUNT(*) FROM {quoted}").fetchone()[0]
    return conn.execute(f"SELECT COUNT(*) FROM {quoted} WHERE run_id = ?", (run_id,)).fetchone()[0]


def _record(results: list[dict], name: str, func: Callable[[], str | None]) -> None:
    try:
        detail = func() or "Conforme"
        results.append({"prueba": name, "ok": True, "detalle": detail})
    except Exception as exc:
        results.append({"prueba": name, "ok": False, "detalle": str(exc)})


def run_functional_tests(
    excel_path: str | Path = DEFAULT_EXCEL_PATH,
    db_path: str | Path = DEFAULT_DB_PATH,
    reports_dir: str | Path = REPORTS_DIR,
    documents_dir: str | Path = REPORTS_DIR,
) -> list[dict]:
    excel_path = Path(excel_path)
    db_path = Path(db_path)
    reports_dir = Path(reports_dir)
    documents_dir = Path(documents_dir)
    initialize_database(db_path)
    results: list[dict] = []
    state: dict[str, int] = {}

    _record(
        results,
        "Validacion de hojas obligatorias",
        lambda: _assert_validation(excel_path),
    )
    _record(
        results,
        "Importacion correcta del Excel",
        lambda: _assert_import(excel_path, db_path),
    )
    _record(
        results,
        "Bloqueo de Modelo 2 y Modelo 3",
        lambda: _assert_model_blocking(db_path),
    )
    _record(
        results,
        "Ejecucion de calculo ABC",
        lambda: _assert_calculation(db_path, state),
    )
    for name, table in [
        ("Calculo MOD -> AC", "mod_actividad"),
        ("Calculo otros recursos -> AC", "otros_recursos_actividad"),
        ("Calculo soporte -> AC", "soporte_actividad"),
        ("Calculo AC -> OT", "produccion_ot"),
        ("Calculo hoja de costos", "hoja_costos_ot"),
    ]:
        _record(results, name, lambda table=table: _assert_generated_rows(db_path, table, state["run_id"]))
    _record(
        results,
        "Generacion de reportes Excel",
        lambda: _assert_reports(db_path, reports_dir),
    )
    _record(
        results,
        "Generacion de documentos Word",
        lambda: _assert_documents(db_path, documents_dir, results),
    )
    return results


def _assert_validation(excel_path: Path) -> str:
    result = validate_workbook(excel_path)
    if not result.ok:
        raise AssertionError("; ".join(issue.message for issue in result.errors))
    return f"Excel valido con {len(result.warnings)} advertencias."


def _assert_import(excel_path: Path, db_path: Path) -> str:
    result = import_excel_to_sqlite(excel_path, db_path)
    if not result.ok:
        raise AssertionError("; ".join(issue.message for issue in result.errors))
    with db_session(db_path) as conn:
        empresas = _count(conn, "empresas")
        ots = _count(conn, "ordenes_trabajo")
        actividades = _count(conn, "actividades")
    if empresas == 0 or ots == 0 or actividades == 0:
        raise AssertionError("No se importaron empresas, ordenes o actividades.")
    return f"Importadas {sum(result.imported_rows.values())} filas no vacias."


def _assert_model_blocking(db_path: Path) -> str:
    models = evaluate_models(db_path)
    model_2 = next(item for item in models if item.code == "MODELO_2")
    model_3 = next(item for item in models if item.code == "MODELO_3")
    if model_2.available or model_3.available:
        raise AssertionError("Modelo 2 o Modelo 3 no deben estar disponibles.")
    return "Modelo 2 y Modelo 3 bloqueados por datos faltantes."


def _assert_calculation(db_path: Path, state: dict[str, int]) -> str:
    run_id = run_model_1_calculation(db_path)
    state["run_id"] = run_id
    if not run_id:
        raise AssertionError("No se genero run_id.")
    return f"Calculo completado con run_id {run_id}."


def _assert_generated_rows(db_path: Path, table: str, run_id: int) -> str:
    with db_session(db_path) as conn:
        count = _count(conn, table, run_id)
    if count == 0:
        raise AssertionError(f"La tabla {table} no tiene resultados.")
    return f"{table}: {count} filas."


def _assert_reports(db_path: Path, reports_dir: Path) -> str:
    paths = generate_all_reports(db_path, reports_dir)
    missing = [name for name in OUTPUT_REPORTS if not (reports_dir / name).exists()]
    if missing:
        raise AssertionError("Reportes faltantes: " + ", ".join(missing))
    return f"Reportes generados: {len(paths)}."


def _assert_documents(db_path: Path, documents_dir: Path, test_results: list[dict]) -> str:
    paths = generate_all_documents(db_path, documents_dir, test_results)
    missing = [name for name in OUTPUT_DOCUMENTS if not (documents_dir / name).exists()]
    if missing:
        raise AssertionError("Documentos faltantes: " + ", ".join(missing))
    return f"Documentos generados: {len(paths)}."


def main() -> None:
    parser = argparse.ArgumentParser(description="Ejecuta pruebas funcionales del Sistema de Costos ABC Frenosa.")
    parser.add_argument("--excel", default=str(DEFAULT_EXCEL_PATH))
    parser.add_argument("--db", default=str(DEFAULT_DB_PATH))
    parser.add_argument("--reports-dir", default=str(REPORTS_DIR))
    args = parser.parse_args()
    results = run_functional_tests(args.excel, args.db, args.reports_dir, args.reports_dir)
    for item in results:
        status = "OK" if item["ok"] else "ERROR"
        print(f"[{status}] {item['prueba']}: {item['detalle']}")
    if not all(item["ok"] for item in results):
        raise SystemExit(1)


if __name__ == "__main__":
    main()
