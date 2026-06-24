from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.utils.exceptions import InvalidFileException

from .constants import OPTIONAL_SHEET_SPECS, REQUIRED_FILES, SHEET_SPECS, SHEET_TO_TABLE
from .database import (
    add_import_issue,
    clear_imported_tables,
    create_import_batch,
    db_session,
    initialize_database,
    insert_rows,
    update_import_batch,
)
from .utils import is_blank_row, normalize_header, safe_float


NUMERIC_COLUMNS = {
    "peso_kg",
    "cantidad_piezas",
    "cantidad_producida",
    "cantidad_kg",
    "costo_unitario_soles",
    "kg_producidos_periodo",
    "horas_reales_trabajadas_periodo",
    "cantidad_operarios",
    "salario_mensual_operario_soles",
    "tiempo_normal_min_kg",
    "tolerancia_personal_pct",
    "tolerancia_fatiga_pct",
    "tolerancia_pie_pct",
    "tolerancia_encorvado_pct",
    "redondear_te_decimales",
    "redondear_porcentaje_decimales",
    "horas_estandar_totales_periodo",
    "tiempo_real_total_minutos_periodo",
    "importe_soles",
    "porcentaje_inductor",
    "costo_equipo_soles",
    "cantidad",
    "porcentaje_distribucion",
    "ventas_netas_soles",
}


@dataclass
class ImportIssue:
    severity: str
    message: str
    sheet_name: str | None = None
    row_number: int | None = None
    column_name: str | None = None


@dataclass
class ImportResult:
    ok: bool
    batch_id: int | None = None
    imported_rows: dict[str, int] = field(default_factory=dict)
    issues: list[ImportIssue] = field(default_factory=list)

    @property
    def errors(self) -> list[ImportIssue]:
        return [issue for issue in self.issues if issue.severity.upper() == "ERROR"]

    @property
    def warnings(self) -> list[ImportIssue]:
        return [issue for issue in self.issues if issue.severity.upper() == "ADVERTENCIA"]


def check_required_files(root: str | Path = ".") -> list[ImportIssue]:
    base = Path(root)
    issues: list[ImportIssue] = []
    for rel in REQUIRED_FILES:
        if not (base / rel).exists():
            issues.append(ImportIssue("ERROR", f"No se encontro el archivo obligatorio: {rel}"))
    return issues


def _headers_match(actual: list[Any], expected: list[str]) -> tuple[bool, list[str]]:
    normalized_actual = [normalize_header(item) for item in actual[: len(expected)]]
    normalized_expected = [normalize_header(item) for item in expected]
    missing = [
        expected_item
        for expected_item, normalized in zip(expected, normalized_expected)
        if normalized not in normalized_actual
    ]
    return not missing, missing


def validate_workbook(excel_path: str | Path) -> ImportResult:
    path = Path(excel_path)
    issues: list[ImportIssue] = []
    if not path.exists():
        return ImportResult(False, issues=[ImportIssue("ERROR", f"No existe el archivo Excel: {path}")])
    if path.suffix.lower() != ".xlsx":
        return ImportResult(False, issues=[ImportIssue("ERROR", "El importador acepta archivos .xlsx.")])

    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except PermissionError:
        return ImportResult(
            False,
            issues=[
                ImportIssue(
                    "ERROR",
                    "No se puede abrir el Excel maestro porque esta bloqueado o sin permiso de lectura. "
                    "Cierre el archivo en Excel, cierre la vista previa de OneDrive/Explorador y vuelva a validar.",
                )
            ],
        )
    except (InvalidFileException, OSError) as exc:
        return ImportResult(
            False,
            issues=[
                ImportIssue(
                    "ERROR",
                    f"No se pudo leer el Excel maestro. Verifique que sea un archivo .xlsx valido y no este danado. Detalle: {exc}",
                )
            ],
        )
    sheet_names = set(wb.sheetnames)

    for sheet, headers in SHEET_SPECS.items():
        if sheet not in sheet_names:
            issues.append(ImportIssue("ERROR", f"Falta la hoja obligatoria: {sheet}", sheet))
            continue
        ws = wb[sheet]
        actual = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
        ok, missing = _headers_match(actual, headers)
        if not ok:
            issues.append(
                ImportIssue(
                    "ERROR",
                    f"La hoja {sheet} no contiene las columnas obligatorias: {', '.join(missing)}",
                    sheet,
                )
            )

    for sheet in sorted(sheet_names - set(SHEET_SPECS) - set(OPTIONAL_SHEET_SPECS)):
        if sheet.lower() in {"proceso_1", "proceso_2", "proceso_3", "hoja_costos_ot", "resumen_costos", "dashboard"}:
            issues.append(
                ImportIssue(
                    "ADVERTENCIA",
                    f"La hoja {sheet} parece ser de salida; sera ignorada como entrada.",
                    sheet,
                )
            )
        else:
            issues.append(ImportIssue("ADVERTENCIA", f"Hoja adicional no mapeada: {sheet}", sheet))

    for sheet, headers in {**SHEET_SPECS, **OPTIONAL_SHEET_SPECS}.items():
        if sheet not in sheet_names:
            continue
        ws = wb[sheet]
        actual = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
        normalized_actual = [normalize_header(value) for value in actual]
        for expected in headers:
            if normalize_header(expected) not in normalized_actual:
                issues.append(
                    ImportIssue(
                        "ADVERTENCIA" if sheet in OPTIONAL_SHEET_SPECS else "ERROR",
                        f"Columna esperada no encontrada: {expected}",
                        sheet,
                        1,
                        expected,
                    )
                )

    errors = [issue for issue in issues if issue.severity == "ERROR"]
    return ImportResult(ok=not errors, issues=issues)


def _read_sheet_rows(workbook: openpyxl.Workbook, sheet_name: str, expected_headers: list[str]) -> list[dict[str, Any]]:
    ws = workbook[sheet_name]
    raw_headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    header_map: dict[int, str] = {}
    for idx, raw in enumerate(raw_headers):
        for expected in expected_headers:
            if normalize_header(raw) == normalize_header(expected):
                header_map[idx] = expected
                break
    rows: list[dict[str, Any]] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        item = {header: row[idx] if idx < len(row) else None for idx, header in header_map.items()}
        if is_blank_row(item):
            continue
        for col in list(item):
            if col in NUMERIC_COLUMNS:
                item[col] = safe_float(item[col])
        rows.append(item)
    return rows


def import_excel_to_sqlite(excel_path: str | Path, db_path: str | Path) -> ImportResult:
    initialize_database(db_path)
    validation = validate_workbook(excel_path)
    with db_session(db_path) as conn:
        batch_id = create_import_batch(conn, excel_path)
        for issue in validation.issues:
            add_import_issue(
                conn,
                batch_id,
                issue.severity,
                issue.message,
                issue.sheet_name,
                issue.row_number,
                issue.column_name,
            )
        if not validation.ok:
            update_import_batch(conn, batch_id, "ERROR", "Validacion de estructura fallida.")
            validation.batch_id = batch_id
            return validation

        clear_imported_tables(conn)
        try:
            wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
        except PermissionError:
            message = (
                "No se puede abrir el Excel maestro porque esta bloqueado o sin permiso de lectura. "
                "Cierre el archivo en Excel, cierre la vista previa de OneDrive/Explorador y vuelva a importar."
            )
            add_import_issue(conn, batch_id, "ERROR", message)
            update_import_batch(conn, batch_id, "ERROR", "No se pudo abrir el Excel maestro.")
            validation.ok = False
            validation.batch_id = batch_id
            validation.issues.append(ImportIssue("ERROR", message))
            return validation
        imported_rows: dict[str, int] = {}
        for sheet, table in SHEET_TO_TABLE.items():
            if sheet not in wb.sheetnames:
                continue
            expected_headers = SHEET_SPECS.get(sheet) or OPTIONAL_SHEET_SPECS.get(sheet) or []
            rows = _read_sheet_rows(wb, sheet, expected_headers)
            count = insert_rows(conn, table, rows, batch_id)
            imported_rows[sheet] = count
            for index, row in enumerate(rows, start=2):
                business_key = (
                    row.get("empresa_id")
                    or row.get("periodo_id")
                    or row.get("modelo_id")
                    or row.get("ot_id")
                    or row.get("actividad_id")
                    or row.get("elemento_costo_id")
                    or row.get("soporte_item_id")
                    or row.get("venta_id")
                    or row.get("fuente_id")
                    or row.get("hoja")
                )
                conn.execute(
                    """
                    INSERT INTO source_trace
                        (batch_id, table_name, sheet_name, source_row, business_key, source_text, observation)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        batch_id,
                        table,
                        sheet,
                        index,
                        str(business_key or ""),
                        str(row.get("fuente") or row.get("fuente_cuadro") or ""),
                        str(row.get("observacion") or ""),
                    ),
                )

        _validate_loaded_relations(conn, batch_id)
        errors = conn.execute(
            "SELECT COUNT(*) FROM import_errors WHERE batch_id = ? AND severity = 'ERROR'",
            (batch_id,),
        ).fetchone()[0]
        status = "ERROR" if errors else "COMPLETADO"
        notes = "Importacion completada." if not errors else "Importacion con errores de integridad."
        update_import_batch(conn, batch_id, status, notes)
        validation.ok = not errors
        validation.batch_id = batch_id
        validation.imported_rows = imported_rows
        validation.issues = [
            ImportIssue(row["severity"], row["message"], row["sheet_name"], row["row_number"], row["column_name"])
            for row in conn.execute(
                "SELECT severity, message, sheet_name, row_number, column_name FROM import_errors WHERE batch_id = ?",
                (batch_id,),
            ).fetchall()
        ]
        return validation


def _validate_loaded_relations(conn, batch_id: int) -> None:
    relation_checks = [
        (
            "Ordenes_Trabajo",
            "objeto_costo_id",
            "SELECT ot_id, objeto_costo_id FROM ordenes_trabajo WHERE objeto_costo_id NOT IN (SELECT objeto_costo_id FROM objetos_costo)",
        ),
        (
            "Materiales_Directos",
            "ot_id",
            "SELECT material_id, ot_id FROM materiales_directos WHERE ot_id NOT IN (SELECT ot_id FROM ordenes_trabajo)",
        ),
        (
            "EC_AC_Inductores",
            "actividad_id",
            "SELECT elemento_costo_id, actividad_id FROM ec_ac_inductores WHERE actividad_id NOT IN (SELECT actividad_id FROM actividades)",
        ),
        (
            "AC_OT_Pesos_Produccion",
            "ot_id",
            "SELECT actividad_id, ot_id FROM ac_ot_pesos_produccion WHERE ot_id NOT IN (SELECT ot_id FROM ordenes_trabajo)",
        ),
        (
            "AC_OT_Pesos_Soporte",
            "ot_id",
            "SELECT actividad_id, ot_id FROM ac_ot_pesos_soporte WHERE ot_id NOT IN (SELECT ot_id FROM ordenes_trabajo)",
        ),
    ]
    for sheet, column, sql in relation_checks:
        for row in conn.execute(sql).fetchall():
            add_import_issue(
                conn,
                batch_id,
                "ERROR",
                f"Relacion invalida en {sheet}: {dict(row)}",
                sheet,
                None,
                column,
            )

    percentage_checks = [
        ("EC_AC_Inductores", "elemento_costo_id", "porcentaje_inductor", "ec_ac_inductores"),
        ("AC_OT_Pesos_Produccion", "actividad_id", "porcentaje_inductor", "ac_ot_pesos_produccion"),
        ("AC_OT_Pesos_Soporte", "actividad_id", "porcentaje_inductor", "ac_ot_pesos_soporte"),
        ("Soporte_CC_ING_Inductores", "proceso_soporte", "porcentaje_distribucion", "soporte_cc_ing_inductores"),
    ]
    for sheet, group_col, pct_col, table in percentage_checks:
        rows = conn.execute(
            f"SELECT {group_col} AS group_key, SUM({pct_col}) AS total_pct FROM {table} GROUP BY {group_col}"
        ).fetchall()
        for row in rows:
            total = row["total_pct"] or 0
            if abs(total - 1.0) > 0.02:
                add_import_issue(
                    conn,
                    batch_id,
                    "ADVERTENCIA",
                    f"Los pesos de {sheet} para {row['group_key']} suman {total:.4f}; revisar redondeos.",
                    sheet,
                    None,
                    pct_col,
                )
