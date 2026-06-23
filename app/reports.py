from __future__ import annotations

from pathlib import Path
from typing import Iterable

from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .constants import OUTPUT_REPORTS, REPORTS_DIR
from .database import db_session, latest_run_id, quote_identifier


TITLE_FILL = PatternFill("solid", fgColor="1F4D78")
HEADER_FILL = PatternFill("solid", fgColor="E8EEF5")
WHITE_FONT = Font(color="FFFFFF", bold=True)
BOLD_FONT = Font(bold=True)


def _write_rows(ws, rows: list[dict], title: str) -> None:
    ws.append([title])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(1, len(rows[0]) if rows else 4))
    ws["A1"].fill = TITLE_FILL
    ws["A1"].font = WHITE_FONT
    ws["A1"].alignment = Alignment(horizontal="center")
    if not rows:
        ws.append(["Sin datos generados."])
        return
    headers = list(rows[0].keys())
    ws.append(headers)
    for cell in ws[2]:
        cell.fill = HEADER_FILL
        cell.font = BOLD_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in rows:
        ws.append([row.get(header) for header in headers])
    _autosize(ws)


def _autosize(ws) -> None:
    for idx, column in enumerate(ws.columns, start=1):
        width = 10
        for cell in column:
            value = "" if cell.value is None else str(cell.value)
            width = max(width, min(len(value) + 2, 45))
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        ws.column_dimensions[get_column_letter(idx)].width = width


def _table(conn, name: str, run_id: int | None = None) -> list[dict]:
    sql = f"SELECT * FROM {quote_identifier(name)}"
    params = []
    if run_id is not None and _has_run_id(conn, name):
        sql += " WHERE run_id = ?"
        params.append(run_id)
    rows = conn.execute(sql, params).fetchall()
    return [dict(row) for row in rows]


def _query(conn, sql: str, params: tuple = ()) -> list[dict]:
    return [dict(row) for row in conn.execute(sql, params).fetchall()]


def _has_run_id(conn, table: str) -> bool:
    return any(row["name"] == "run_id" for row in conn.execute(f"PRAGMA table_info({quote_identifier(table)})").fetchall())


def _save_workbook(wb: Workbook, path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def _add_formula_sheet(wb: Workbook, formulas: list[tuple[str, str]]) -> None:
    ws = wb.create_sheet("Formulas")
    ws.append(["Modulo", "Formula / criterio"])
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = BOLD_FONT
    for item in formulas:
        ws.append(list(item))
    _autosize(ws)


def generate_proceso_1_report(db_path: str | Path, output_dir: str | Path = REPORTS_DIR) -> Path:
    with db_session(db_path) as conn:
        run_id = latest_run_id(conn)
        wb = Workbook()
        wb.remove(wb.active)
        for sheet, rows in [
            (
                "PD_AC",
                _query(
                    conn,
                    """
                    SELECT run_id, actividad_id, codigo_actividad, nombre_actividad,
                           ts_total, te, fd, tiempo_real_minutos,
                           porcentaje_calculado, porcentaje_usado, total_mod_soles,
                           costo_mod_ac AS costo_pd_ac, formula, fuente, observacion
                    FROM mod_actividad
                    WHERE run_id = ?
                    """,
                    (run_id,),
                ),
            ),
            ("Otros_Recursos_AC", _table(conn, "otros_recursos_actividad", run_id)),
            ("Soporte_AC", _table(conn, "soporte_actividad", run_id)),
            (
                "Total_Actividad",
                _query(
                    conn,
                    """
                    SELECT run_id, actividad_id, codigo_actividad, nombre_actividad,
                           costo_mod_ac AS costo_pd_ac, costo_otros_recursos_ac,
                           costo_conversion_ac, costo_soporte_ac, costo_total_ac,
                           formula, fuente, observacion
                    FROM actividad_total
                    WHERE run_id = ?
                    """,
                    (run_id,),
                ),
            ),
            ("Detalle_EC_AC", _table(conn, "otros_recursos_ac_detalle", run_id)),
        ]:
            ws = wb.create_sheet(sheet)
            _write_rows(ws, rows, f"Proceso 1: EC -> AC - {sheet}")
        _add_formula_sheet(
            wb,
            [
                ("PD -> AC", "costo_PD_AC = porcentaje_AC * total_PD"),
                ("Otros recursos -> AC", "costo_recurso_AC = importe_recurso * porcentaje_inductor"),
                ("Soporte -> AC", "soporte_AC = mantenimiento_AC + control_calidad_AC + ingenieria_AC"),
            ],
        )
        return _save_workbook(wb, Path(output_dir) / "Reporte_Proceso_1_EC_AC.xlsx")


def generate_proceso_2_report(db_path: str | Path, output_dir: str | Path = REPORTS_DIR) -> Path:
    with db_session(db_path) as conn:
        run_id = latest_run_id(conn)
        wb = Workbook()
        wb.remove(wb.active)
        for sheet, rows, title in [
            (
                "Proceso_2_Produccion",
                _query(
                    conn,
                    """
                    SELECT d.run_id, d.actividad_id, d.codigo_actividad, d.nombre_actividad,
                           d.ot_id, d.objeto_costo_id, d.codigo_producto, d.peso_ac_ot,
                           d.costo_conversion_ac, d.costo_conversion_ot,
                           p.costo_produccion_ot, d.formula, d.fuente, d.observacion
                    FROM produccion_ot_detalle d
                    JOIN produccion_ot p
                      ON p.run_id = d.run_id AND p.ot_id = d.ot_id
                    WHERE d.run_id = ?
                    ORDER BY d.ot_id, d.actividad_id
                    """,
                    (run_id,),
                ),
                "Proceso 2 AC-OT (Produccion)",
            ),
            (
                "Proceso_2_Soporte",
                _query(
                    conn,
                    """
                    SELECT d.run_id, d.actividad_id, d.codigo_actividad, d.nombre_actividad,
                           d.ot_id, d.objeto_costo_id, d.codigo_producto, d.peso_ac_ot,
                           d.costo_soporte_ac, d.costo_soporte_ot AS costo_soporte_ot_detalle,
                           s.costo_soporte_ot, d.formula, d.fuente, d.observacion
                    FROM soporte_ot_detalle d
                    JOIN soporte_ot s
                      ON s.run_id = d.run_id AND s.ot_id = d.ot_id
                    WHERE d.run_id = ?
                    ORDER BY d.ot_id, d.actividad_id
                    """,
                    (run_id,),
                ),
                "Proceso 2 AC-OT (Soporte)",
            ),
        ]:
            ws = wb.create_sheet(sheet)
            _write_rows(ws, rows, title)
        _add_formula_sheet(
            wb,
            [
                ("Produccion AC -> OT", "costo_produccion_OT = SUM(costo_conversion_AC * peso_AC_OT_produccion)"),
                ("Soporte AC -> OT", "costo_soporte_OT = SUM(costo_soporte_AC * peso_AC_OT_soporte)"),
            ],
        )
        return _save_workbook(wb, Path(output_dir) / "Reporte_Proceso_2_AC_OT.xlsx")


def generate_cost_sheet_report(db_path: str | Path, output_dir: str | Path = REPORTS_DIR) -> Path:
    with db_session(db_path) as conn:
        run_id = latest_run_id(conn)
        wb = Workbook()
        ws = wb.active
        ws.title = "Hoja_Costos_OT"
        _write_rows(ws, _table(conn, "hoja_costos_ot", run_id), "Hoja de costos por OT")
        _add_formula_sheet(
            wb,
            [("Hoja de costos", "costo_total_OT = CMD + costo_produccion + costo_soporte + CSD")],
        )
        return _save_workbook(wb, Path(output_dir) / "Reporte_Hoja_Costos_OT.xlsx")


def generate_estado_resultados_report(db_path: str | Path, output_dir: str | Path = REPORTS_DIR) -> Path:
    with db_session(db_path) as conn:
        run_id = latest_run_id(conn)
        wb = Workbook()
        ws = wb.active
        ws.title = "Estado_Resultados"
        _write_rows(ws, _table(conn, "estado_resultados_ot", run_id), "Estado de resultados por OT")
        _add_formula_sheet(
            wb,
            [
                (
                    "Estado de resultados",
                    "ventas - materia_prima_directa - costos_conversion = margen_bruto; margen_bruto - soporte = margen_operativo",
                )
            ],
        )
        return _save_workbook(wb, Path(output_dir) / "Reporte_Estado_Resultados.xlsx")


def generate_dashboard_report(db_path: str | Path, output_dir: str | Path = REPORTS_DIR) -> Path:
    with db_session(db_path) as conn:
        run_id = latest_run_id(conn)
        wb = Workbook()
        wb.remove(wb.active)
        for sheet, table in [
            ("Hoja_Costos", "hoja_costos_ot"),
            ("Estado_Resultados", "estado_resultados_ot"),
            ("Costos_Actividad", "actividad_total"),
            ("Alertas", "dashboard_alertas"),
        ]:
            ws = wb.create_sheet(sheet)
            _write_rows(ws, _table(conn, table, run_id), f"Dashboard - {sheet}")
        _add_dashboard_charts(wb)
        return _save_workbook(wb, Path(output_dir) / "Reporte_Dashboard.xlsx")


def _add_dashboard_charts(wb: Workbook) -> None:
    if "Hoja_Costos" in wb.sheetnames:
        ws = wb["Hoja_Costos"]
        if ws.max_row >= 3:
            headers = [cell.value for cell in ws[2]]
            try:
                label_col = headers.index("codigo_ot") + 1
                value_col = headers.index("costo_total_ot") + 1
                chart = BarChart()
                chart.title = "Costo total por OT"
                chart.y_axis.title = "Soles"
                chart.x_axis.title = "Orden de trabajo"
                data = Reference(ws, min_col=value_col, min_row=2, max_row=ws.max_row)
                labels = Reference(ws, min_col=label_col, min_row=3, max_row=ws.max_row)
                chart.add_data(data, titles_from_data=True)
                chart.set_categories(labels)
                ws.add_chart(chart, "V3")
            except ValueError:
                pass
    if "Costos_Actividad" in wb.sheetnames:
        ws = wb["Costos_Actividad"]
        if ws.max_row >= 3:
            headers = [cell.value for cell in ws[2]]
            try:
                label_col = headers.index("nombre_actividad") + 1
                value_col = headers.index("costo_total_ac") + 1
                chart = PieChart()
                chart.title = "Participacion por actividad"
                data = Reference(ws, min_col=value_col, min_row=2, max_row=ws.max_row)
                labels = Reference(ws, min_col=label_col, min_row=3, max_row=ws.max_row)
                chart.add_data(data, titles_from_data=True)
                chart.set_categories(labels)
                ws.add_chart(chart, "V3")
            except ValueError:
                pass


def generate_all_reports(db_path: str | Path, output_dir: str | Path = REPORTS_DIR) -> list[Path]:
    return [
        generate_proceso_1_report(db_path, output_dir),
        generate_proceso_2_report(db_path, output_dir),
        generate_cost_sheet_report(db_path, output_dir),
        generate_estado_resultados_report(db_path, output_dir),
        generate_dashboard_report(db_path, output_dir),
    ]
