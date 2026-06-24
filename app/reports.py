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
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = ws.dimensions
    _format_numeric_columns(ws)
    _autosize(ws)


def _format_numeric_columns(ws) -> None:
    percent_names = {"porcentaje", "peso", "rentabilidad", "margen"}
    currency_names = {"costo", "cmd", "csd", "ventas", "importe", "salario", "total", "materia", "conversion", "soporte"}
    headers = [cell.value for cell in ws[2]] if ws.max_row >= 2 else []
    for col_idx, header in enumerate(headers, start=1):
        header_text = str(header or "").lower()
        if any(token in header_text for token in percent_names):
            number_format = "0.00%"
        elif any(token in header_text for token in currency_names):
            number_format = '#,##0.00'
        else:
            continue
        for row_idx in range(3, ws.max_row + 1):
            ws.cell(row=row_idx, column=col_idx).number_format = number_format


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


def _add_summary_sheet(wb: Workbook, title: str, items: list[tuple[str, object]]) -> None:
    ws = wb.create_sheet("Resumen", 0)
    ws.append([title])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
    ws["A1"].fill = TITLE_FILL
    ws["A1"].font = WHITE_FONT
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.append(["Indicador", "Valor", "Interpretacion", "Fuente"])
    for cell in ws[2]:
        cell.fill = HEADER_FILL
        cell.font = BOLD_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for label, value in items:
        ws.append([label, value, _summary_interpretation(label, value), "Calculos generados desde SQLite"])
    ws.freeze_panes = "A3"
    _format_numeric_columns(ws)
    _autosize(ws)


def _summary_interpretation(label: str, value: object) -> str:
    label_lower = label.lower()
    if "mayor costo" in label_lower:
        return "Elemento prioritario para revision de eficiencia y uso de inductores."
    if "menor rentabilidad" in label_lower:
        return "Objeto de costo que requiere analisis comercial y operativo."
    if "total" in label_lower:
        return "Valor acumulado de la ejecucion del modelo ABC."
    return "Indicador de control para lectura academica del reporte."


def _safe_number(row: dict | None, field: str) -> float:
    if not row:
        return 0.0
    try:
        return float(row.get(field) or 0)
    except (TypeError, ValueError):
        return 0.0


def _sum(rows: list[dict], field: str) -> float:
    return sum(_safe_number(row, field) for row in rows)


def _max_label(rows: list[dict], value_field: str, label_field: str) -> str:
    if not rows:
        return "Sin datos"
    row = max(rows, key=lambda item: _safe_number(item, value_field))
    return f"{row.get(label_field) or row.get('codigo_producto') or row.get('ot_id')} - { _safe_number(row, value_field):,.2f}"


def _min_label(rows: list[dict], value_field: str, label_field: str) -> str:
    if not rows:
        return "Sin datos"
    row = min(rows, key=lambda item: _safe_number(item, value_field))
    value = _safe_number(row, value_field)
    suffix = "%" if "rentabilidad" in value_field else ""
    display_value = value * 100 if suffix else value
    return f"{row.get(label_field) or row.get('codigo_producto') or row.get('ot_id')} - {display_value:,.2f}{suffix}"


def _header_index(ws, header: str) -> int | None:
    if ws.max_row < 2:
        return None
    headers = [cell.value for cell in ws[2]]
    try:
        return headers.index(header) + 1
    except ValueError:
        return None


def _add_bar_chart(ws, label_header: str, value_header: str, title: str, position: str, y_axis: str = "Soles") -> None:
    if ws.max_row < 3:
        return
    label_col = _header_index(ws, label_header)
    value_col = _header_index(ws, value_header)
    if not label_col or not value_col:
        return
    chart = BarChart()
    chart.title = title
    chart.y_axis.title = y_axis
    chart.x_axis.title = label_header
    chart.height = 7
    chart.width = 14
    data = Reference(ws, min_col=value_col, min_row=2, max_row=ws.max_row)
    labels = Reference(ws, min_col=label_col, min_row=3, max_row=ws.max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)
    ws.add_chart(chart, position)


def _add_pie_chart(ws, label_header: str, value_header: str, title: str, position: str) -> None:
    if ws.max_row < 3:
        return
    label_col = _header_index(ws, label_header)
    value_col = _header_index(ws, value_header)
    if not label_col or not value_col:
        return
    chart = PieChart()
    chart.title = title
    chart.height = 7
    chart.width = 11
    data = Reference(ws, min_col=value_col, min_row=2, max_row=ws.max_row)
    labels = Reference(ws, min_col=label_col, min_row=3, max_row=ws.max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)
    ws.add_chart(chart, position)


def _add_stacked_bar_chart(ws, label_header: str, value_headers: list[str], title: str, position: str) -> None:
    if ws.max_row < 3:
        return
    label_col = _header_index(ws, label_header)
    value_cols = [_header_index(ws, header) for header in value_headers]
    if not label_col or any(col is None for col in value_cols):
        return
    chart = BarChart()
    chart.type = "bar"
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.title = title
    chart.y_axis.title = label_header
    chart.x_axis.title = "Soles"
    chart.height = 8
    chart.width = 15
    for value_col in value_cols:
        data = Reference(ws, min_col=value_col, min_row=2, max_row=ws.max_row)
        chart.add_data(data, titles_from_data=True)
    labels = Reference(ws, min_col=label_col, min_row=3, max_row=ws.max_row)
    chart.set_categories(labels)
    ws.add_chart(chart, position)


def generate_proceso_1_report(db_path: str | Path, output_dir: str | Path = REPORTS_DIR) -> Path:
    with db_session(db_path) as conn:
        run_id = latest_run_id(conn)
        wb = Workbook()
        wb.remove(wb.active)
        total_rows: list[dict] = []
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
            if sheet == "Total_Actividad":
                total_rows = rows
                _add_bar_chart(ws, "nombre_actividad", "costo_total_ac", "Costo total por actividad", "N3")
                _add_pie_chart(ws, "nombre_actividad", "costo_total_ac", "Participacion del costo por actividad", "N20")
        _add_formula_sheet(
            wb,
            [
                ("PD -> AC", "costo_PD_AC = porcentaje_AC * total_PD"),
                ("Otros recursos -> AC", "costo_recurso_AC = importe_recurso * porcentaje_inductor"),
                ("Soporte -> AC", "soporte_AC = mantenimiento_AC + control_calidad_AC + ingenieria_AC"),
            ],
        )
        _add_summary_sheet(
            wb,
            "Resumen ejecutivo - Proceso 1 EC -> AC",
            [
                ("Costo total actividades", _sum(total_rows, "costo_total_ac")),
                ("Actividad de mayor costo", _max_label(total_rows, "costo_total_ac", "nombre_actividad")),
                ("Costo total PD", _sum(total_rows, "costo_pd_ac")),
                ("Costo total soporte", _sum(total_rows, "costo_soporte_ac")),
            ],
        )
        return _save_workbook(wb, Path(output_dir) / "Reporte_Proceso_1_EC_AC.xlsx")


def generate_proceso_2_report(db_path: str | Path, output_dir: str | Path = REPORTS_DIR) -> Path:
    with db_session(db_path) as conn:
        run_id = latest_run_id(conn)
        wb = Workbook()
        wb.remove(wb.active)
        produccion_rows: list[dict] = []
        soporte_rows: list[dict] = []
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
            if sheet == "Proceso_2_Produccion":
                produccion_rows = rows
                _add_bar_chart(ws, "codigo_producto", "costo_conversion_ot", "Detalle costo produccion AC -> OT", "Q3")
            if sheet == "Proceso_2_Soporte":
                soporte_rows = rows
                _add_bar_chart(ws, "codigo_producto", "costo_soporte_ot_detalle", "Detalle costo soporte AC -> OT", "Q3")
        _add_formula_sheet(
            wb,
            [
                ("Produccion AC -> OT", "costo_produccion_OT = SUM(costo_conversion_AC * peso_AC_OT_produccion)"),
                ("Soporte AC -> OT", "costo_soporte_OT = SUM(costo_soporte_AC * peso_AC_OT_soporte)"),
            ],
        )
        _add_summary_sheet(
            wb,
            "Resumen ejecutivo - Proceso 2 AC -> OT",
            [
                ("Costo produccion asignado", _sum(produccion_rows, "costo_conversion_ot")),
                ("Costo soporte asignado", _sum(soporte_rows, "costo_soporte_ot_detalle")),
                ("Mayor detalle produccion", _max_label(produccion_rows, "costo_conversion_ot", "codigo_producto")),
                ("Mayor detalle soporte", _max_label(soporte_rows, "costo_soporte_ot_detalle", "codigo_producto")),
            ],
        )
        return _save_workbook(wb, Path(output_dir) / "Reporte_Proceso_2_AC_OT.xlsx")


def generate_cost_sheet_report(db_path: str | Path, output_dir: str | Path = REPORTS_DIR) -> Path:
    with db_session(db_path) as conn:
        run_id = latest_run_id(conn)
        wb = Workbook()
        ws = wb.active
        ws.title = "Hoja_Costos_OT"
        rows = _table(conn, "hoja_costos_ot", run_id)
        _write_rows(ws, rows, "Hoja de costos por OT")
        _add_bar_chart(ws, "codigo_producto", "costo_total_ot", "Costo total por OT", "V3")
        _add_stacked_bar_chart(
            ws,
            "codigo_producto",
            ["cmd_ot", "costo_produccion_ot", "costo_soporte_ot", "csd_ot"],
            "Estructura del costo por OT",
            "V20",
        )
        _add_formula_sheet(
            wb,
            [("Hoja de costos", "costo_total_OT = CMD + costo_produccion + costo_soporte + CSD")],
        )
        _add_summary_sheet(
            wb,
            "Resumen ejecutivo - Hoja de costos OT",
            [
                ("Costo total OT", _sum(rows, "costo_total_ot")),
                ("Mayor costo por OT", _max_label(rows, "costo_total_ot", "codigo_producto")),
                ("Costo unitario promedio", _sum(rows, "costo_unitario_ot") / len(rows) if rows else 0),
                ("Total soporte asignado", _sum(rows, "costo_soporte_ot")),
            ],
        )
        return _save_workbook(wb, Path(output_dir) / "Reporte_Hoja_Costos_OT.xlsx")


def generate_estado_resultados_report(db_path: str | Path, output_dir: str | Path = REPORTS_DIR) -> Path:
    with db_session(db_path) as conn:
        run_id = latest_run_id(conn)
        wb = Workbook()
        ws = wb.active
        ws.title = "Estado_Resultados"
        rows = _table(conn, "estado_resultados_ot", run_id)
        _write_rows(ws, rows, "Estado de resultados por OT")
        _add_bar_chart(ws, "codigo_producto", "margen_operativo", "Margen operativo por OT", "P3")
        _add_bar_chart(ws, "codigo_producto", "rentabilidad_operativa", "Rentabilidad operativa por OT", "P20", y_axis="Ratio")
        _add_formula_sheet(
            wb,
            [
                (
                    "Estado de resultados",
                    "ventas - materia_prima_directa - costos_conversion = margen_bruto; margen_bruto - soporte = margen_operativo",
                )
            ],
        )
        _add_summary_sheet(
            wb,
            "Resumen ejecutivo - Estado de resultados",
            [
                ("Ventas netas totales", _sum(rows, "ventas_netas_soles")),
                ("Margen operativo total", _sum(rows, "margen_operativo")),
                ("Menor rentabilidad", _min_label(rows, "rentabilidad_operativa", "codigo_producto")),
                ("Mayor margen operativo", _max_label(rows, "margen_operativo", "codigo_producto")),
            ],
        )
        return _save_workbook(wb, Path(output_dir) / "Reporte_Estado_Resultados.xlsx")


def generate_dashboard_report(db_path: str | Path, output_dir: str | Path = REPORTS_DIR) -> Path:
    with db_session(db_path) as conn:
        run_id = latest_run_id(conn)
        wb = Workbook()
        wb.remove(wb.active)
        dashboard_rows: dict[str, list[dict]] = {}
        for sheet, table in [
            ("Hoja_Costos", "hoja_costos_ot"),
            ("Estado_Resultados", "estado_resultados_ot"),
            ("Costos_Actividad", "actividad_total"),
            ("Alertas", "dashboard_alertas"),
        ]:
            ws = wb.create_sheet(sheet)
            rows = _table(conn, table, run_id)
            dashboard_rows[sheet] = rows
            _write_rows(ws, rows, f"Dashboard - {sheet}")
        _add_dashboard_charts(wb)
        _add_summary_sheet(
            wb,
            "Resumen ejecutivo - Dashboard",
            [
                ("Costo total OT", _sum(dashboard_rows.get("Hoja_Costos", []), "costo_total_ot")),
                ("Ventas netas totales", _sum(dashboard_rows.get("Estado_Resultados", []), "ventas_netas_soles")),
                ("Actividad de mayor costo", _max_label(dashboard_rows.get("Costos_Actividad", []), "costo_total_ac", "nombre_actividad")),
                ("Menor rentabilidad", _min_label(dashboard_rows.get("Estado_Resultados", []), "rentabilidad_operativa", "codigo_producto")),
            ],
        )
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
        _add_stacked_bar_chart(
            ws,
            "codigo_producto",
            ["cmd_ot", "costo_produccion_ot", "costo_soporte_ot", "csd_ot"],
            "Estructura de costo por OT",
            "V20",
        )
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
    if "Estado_Resultados" in wb.sheetnames:
        ws = wb["Estado_Resultados"]
        _add_bar_chart(ws, "codigo_producto", "rentabilidad_operativa", "Rentabilidad operativa por OT", "P3", y_axis="Ratio")


def generate_all_reports(db_path: str | Path, output_dir: str | Path = REPORTS_DIR) -> list[Path]:
    return [
        generate_proceso_1_report(db_path, output_dir),
        generate_proceso_2_report(db_path, output_dir),
        generate_cost_sheet_report(db_path, output_dir),
        generate_estado_resultados_report(db_path, output_dir),
        generate_dashboard_report(db_path, output_dir),
    ]
