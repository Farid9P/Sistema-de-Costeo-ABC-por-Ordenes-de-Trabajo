from __future__ import annotations

from pathlib import Path
from typing import Iterable

from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .constants import OUTPUT_REPORTS, REPORTS_DIR
from .database import db_session, latest_run_id, quote_identifier
from .utils import codigos_ot, con_codigo_ot, etiqueta_objeto_costo


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
        return "Objeto de costo (OT) que requiere analisis comercial y operativo."
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
    return f"{_etiqueta(row, label_field)} - { _safe_number(row, value_field):,.2f}"


def _min_label(rows: list[dict], value_field: str, label_field: str) -> str:
    if not rows:
        return "Sin datos"
    row = min(rows, key=lambda item: _safe_number(item, value_field))
    value = _safe_number(row, value_field)
    suffix = "%" if "rentabilidad" in value_field else ""
    display_value = value * 100 if suffix else value
    return f"{_etiqueta(row, label_field)} - {display_value:,.2f}{suffix}"


def _etiqueta(row: dict, label_field: str) -> str:
    if label_field == "codigo_ot":
        return etiqueta_objeto_costo(row)
    return str(row.get(label_field) or etiqueta_objeto_costo(row))


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
                           d.ot_id, ot.codigo_ot, d.objeto_costo_id, d.codigo_producto, d.peso_ac_ot,
                           d.costo_conversion_ac, d.costo_conversion_ot,
                           p.costo_produccion_ot, d.formula, d.fuente, d.observacion
                    FROM produccion_ot_detalle d
                    JOIN produccion_ot p
                      ON p.run_id = d.run_id AND p.ot_id = d.ot_id
                    LEFT JOIN ordenes_trabajo ot ON ot.ot_id = d.ot_id
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
                           d.ot_id, ot.codigo_ot, d.objeto_costo_id, d.codigo_producto, d.peso_ac_ot,
                           d.costo_soporte_ac, d.costo_soporte_ot AS costo_soporte_ot_detalle,
                           s.costo_soporte_ot, d.formula, d.fuente, d.observacion
                    FROM soporte_ot_detalle d
                    JOIN soporte_ot s
                      ON s.run_id = d.run_id AND s.ot_id = d.ot_id
                    LEFT JOIN ordenes_trabajo ot ON ot.ot_id = d.ot_id
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
                _add_bar_chart(ws, "codigo_ot", "costo_conversion_ot", "Detalle costo produccion AC -> OT", "Q3")
            if sheet == "Proceso_2_Soporte":
                soporte_rows = rows
                _add_bar_chart(ws, "codigo_ot", "costo_soporte_ot_detalle", "Detalle costo soporte AC -> OT", "Q3")
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
                ("Mayor detalle produccion", _max_label(produccion_rows, "costo_conversion_ot", "codigo_ot")),
                ("Mayor detalle soporte", _max_label(soporte_rows, "costo_soporte_ot_detalle", "codigo_ot")),
            ],
        )
        return _save_workbook(wb, Path(output_dir) / "Reporte_Proceso_2_AC_OT.xlsx")


def generate_cost_sheet_report(
    db_path: str | Path,
    output_dir: str | Path = REPORTS_DIR,
    rows_override: list[dict] | None = None,
) -> Path:
    with db_session(db_path) as conn:
        run_id = latest_run_id(conn)
        wb = Workbook()
        ws = wb.active
        ws.title = "Hoja_Costos_OT"
        rows = rows_override if rows_override is not None else _table(conn, "hoja_costos_ot", run_id)
        _write_rows(ws, rows, "Hoja de costos por OT")
        _add_bar_chart(ws, "codigo_ot", "costo_total_ot", "Costo total por OT", "V3")
        _add_stacked_bar_chart(
            ws,
            "codigo_ot",
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
                ("Mayor costo por OT", _max_label(rows, "costo_total_ot", "codigo_ot")),
                ("Costo unitario promedio", _sum(rows, "costo_unitario_ot") / len(rows) if rows else 0),
                ("Total soporte asignado", _sum(rows, "costo_soporte_ot")),
            ],
        )
        return _save_workbook(wb, Path(output_dir) / "Reporte_Hoja_Costos_OT.xlsx")


def generate_estado_resultados_report(
    db_path: str | Path,
    output_dir: str | Path = REPORTS_DIR,
    rows_override: list[dict] | None = None,
) -> Path:
    with db_session(db_path) as conn:
        run_id = latest_run_id(conn)
        wb = Workbook()
        ws = wb.active
        ws.title = "Estado_Resultados"
        rows = rows_override if rows_override is not None else _table(conn, "estado_resultados_ot", run_id)
        rows = con_codigo_ot(rows, codigos_ot(conn))
        _write_rows(ws, rows, "Estado de resultados por OT")
        _add_bar_chart(ws, "codigo_ot", "margen_operativo", "Margen operativo por OT", "P3")
        _add_bar_chart(ws, "codigo_ot", "rentabilidad_operativa", "Rentabilidad operativa por OT", "P20", y_axis="Ratio")
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
                ("Menor rentabilidad", _min_label(rows, "rentabilidad_operativa", "codigo_ot")),
                ("Mayor margen operativo", _max_label(rows, "margen_operativo", "codigo_ot")),
            ],
        )
        return _save_workbook(wb, Path(output_dir) / "Reporte_Estado_Resultados.xlsx")


def generate_dashboard_report(
    db_path: str | Path,
    output_dir: str | Path = REPORTS_DIR,
    payload_override: dict | None = None,
) -> Path:
    wb = Workbook()
    wb.remove(wb.active)
    dashboard_rows: dict[str, list[dict]] = {}
    if payload_override is not None:
        rows_by_sheet = [
            ("Hoja_Costos", payload_override.get("costos_ot", [])),
            ("Estado_Resultados", payload_override.get("estado_resultados", [])),
            ("Costos_Actividad", payload_override.get("costos_actividad", [])),
            ("Alertas", payload_override.get("alertas", [])),
        ]
    else:
        with db_session(db_path) as conn:
            run_id = latest_run_id(conn)
            rows_by_sheet = [
                ("Hoja_Costos", _table(conn, "hoja_costos_ot", run_id)),
                ("Estado_Resultados", _table(conn, "estado_resultados_ot", run_id)),
                ("Costos_Actividad", _table(conn, "actividad_total", run_id)),
                ("Alertas", _table(conn, "dashboard_alertas", run_id)),
            ]
    try:
        with db_session(db_path) as conn:
            codigos = codigos_ot(conn)
    except Exception:
        codigos = {}
    rows_by_sheet = [(sheet, con_codigo_ot(rows, codigos)) for sheet, rows in rows_by_sheet]
    for sheet, rows in rows_by_sheet:
        ws = wb.create_sheet(sheet)
        dashboard_rows[sheet] = rows
        _write_rows(ws, rows, f"Dashboard - {sheet}")
    _add_dashboard_charts(wb)
    summary_items = [
        ("Costo total OT", _sum(dashboard_rows.get("Hoja_Costos", []), "costo_total_ot")),
        ("Ventas netas totales", _sum(dashboard_rows.get("Estado_Resultados", []), "ventas_netas_soles")),
        ("Actividad de mayor costo", _max_label(dashboard_rows.get("Costos_Actividad", []), "costo_total_ac", "nombre_actividad")),
        ("Menor rentabilidad", _min_label(dashboard_rows.get("Estado_Resultados", []), "rentabilidad_operativa", "codigo_ot")),
    ]
    if payload_override and payload_override.get("modo_periodo") == "PROYECTADO":
        summary_items.append(("Vista de periodo", f"Proyectada con factor {payload_override.get('factor_proyeccion', 1):g}"))
    _add_summary_sheet(wb, "Resumen ejecutivo - Dashboard", summary_items)
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
            "codigo_ot",
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
        _add_bar_chart(ws, "codigo_ot", "rentabilidad_operativa", "Rentabilidad operativa por OT", "P3", y_axis="Ratio")


def generate_all_reports(
    db_path: str | Path,
    output_dir: str | Path = REPORTS_DIR,
    dashboard_payload_override: dict | None = None,
) -> list[Path]:
    return [
        generate_proceso_1_report(db_path, output_dir),
        generate_proceso_2_report(db_path, output_dir),
        generate_cost_sheet_report(
            db_path,
            output_dir,
            dashboard_payload_override.get("costos_ot") if dashboard_payload_override else None,
        ),
        generate_estado_resultados_report(
            db_path,
            output_dir,
            dashboard_payload_override.get("estado_resultados") if dashboard_payload_override else None,
        ),
        generate_dashboard_report(db_path, output_dir, dashboard_payload_override),
        generate_kardex_report(db_path, output_dir),
    ]


# ---------------------------------------------------------------------------
# Reporte de kardex por promedio ponderado (hoja Compras del Excel maestro)
# ---------------------------------------------------------------------------
KX_FONT = "Arial"
KX_INPUT = Font(name=KX_FONT, size=10, color="0000FF")
KX_NORMAL = Font(name=KX_FONT, size=10)
KX_BOLD = Font(name=KX_FONT, size=10, bold=True)
KX_LINK = Font(name=KX_FONT, size=10, color="008000")
KX_TITLE = Font(name=KX_FONT, size=12, bold=True)
KX_GROUP_FILL = PatternFill("solid", fgColor="DCE6F1")
KX_YELLOW = PatternFill("solid", fgColor="FFFF00")
KX_NUM = "#,##0.00;[Red]-#,##0.00;-"
KX_CU = "#,##0.0000;[Red]-#,##0.0000;-"
KX_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _kx_ref(sheet: str, cell: str) -> str:
    return "'" + sheet.replace("'", "''") + "'!" + cell


def _kardex_sheet_name(codigo: str, usados: set[str]) -> str:
    nombre = "".join(ch for ch in codigo if ch not in '[]:*?/\\\\')[:28] or "Item"
    base, n = nombre, 2
    while nombre in usados:
        nombre = f"{base[:25]}_{n}"
        n += 1
    usados.add(nombre)
    return nombre


def _write_kardex_item(ws, item, periodo_texto: str, cmd: dict | None = None) -> dict:
    """Escribe el kardex de un item con formulas (formato de la valorizacion de inventario)."""
    from .kardex_compras import COMPRA, CONSUMO, DEVOLUCION_CONSUMO, DEVOLUCION_PROVEEDOR, SALDO_INICIAL

    ws["A1"] = f"VALORIZACIÓN DE INVENTARIO DE {(item.almacen or 'ALMACÉN').upper()}"
    ws["A1"].font = KX_TITLE
    info = [("Material", f"{item.item_codigo} — {item.item_descripcion}"), ("Almacén", item.almacen),
            ("Periodo", periodo_texto), ("Método", "PROMEDIO"), ("Unidad", item.unidad)]
    for i, (k, v) in enumerate(info, start=2):
        ws.cell(i, 1, k).font = KX_NORMAL
        ws.cell(i, 2, v).font = KX_NORMAL
    ws.cell(5, 3, "Valorización de inventario").font = KX_NORMAL

    ws.cell(8, 1, f"Movimientos del periodo {periodo_texto}").font = KX_NORMAL
    for col, titulo in ((4, "INGRESOS"), (7, "SALIDAS"), (10, "SALDOS"), (13, "DATOS DE COMPRA")):
        ws.merge_cells(start_row=8, start_column=col, end_row=8, end_column=col + 2)
        c = ws.cell(8, col, titulo)
        c.font, c.fill, c.alignment = KX_BOLD, KX_GROUP_FILL, KX_CENTER
    headers = ["Descripción", "Fecha", "Detalle", "Cantidad", "Cto Unit S/", "Cto total S/", "Cantidad", "Cto Unit S/",
               "Cto total S/", "Cantidad", "Cto Unit S/", "Cto total S/", "Precio con IGV S/", "IGV", "Flete S/"]
    for col, titulo in enumerate(headers, start=1):
        c = ws.cell(9, col, titulo)
        c.font, c.fill, c.alignment = KX_BOLD, KX_GROUP_FILL, KX_CENTER

    first = 10
    r = first
    filas_compra: dict[str, int] = {}
    for fila in item.filas:
        prev = r - 1
        ws.cell(r, 1, fila.descripcion).font = KX_NORMAL
        try:
            from datetime import date as _date

            fecha_celda = ws.cell(r, 2, _date.fromisoformat(fila.fecha))
            fecha_celda.number_format = "dd/mm/yyyy"
        except ValueError:
            fecha_celda = ws.cell(r, 2, fila.fecha)
        fecha_celda.font = KX_NORMAL
        detalle = fila.detalle + (f" {fila.documento}" if fila.documento and fila.documento not in fila.detalle else "")
        ws.cell(r, 3, detalle).font = KX_NORMAL
        if fila.tipo == SALDO_INICIAL:
            ws.cell(r, 10, fila.saldo_cantidad).font = KX_INPUT
            ws.cell(r, 12, fila.saldo_total).font = KX_INPUT
            ws.cell(r, 11, f"=IFERROR(L{r}/J{r},0)").font = KX_NORMAL
            for col in (10, 11, 12):
                ws.cell(r, col).fill = KX_YELLOW
        else:
            if fila.tipo in (COMPRA, DEVOLUCION_PROVEEDOR):
                ws.cell(r, 4, fila.ing_cantidad).font = KX_INPUT
                if fila.tipo == DEVOLUCION_PROVEEDOR and fila.documento in filas_compra:
                    if fila.precio_con_igv:
                        ws.cell(r, 13, fila.precio_con_igv).font = KX_INPUT
                        ws.cell(r, 14, fila.igv_pct).font = KX_INPUT
                    ws.cell(r, 5, f"=E{filas_compra[fila.documento]}").font = KX_NORMAL
                elif fila.precio_con_igv:
                    ws.cell(r, 13, fila.precio_con_igv).font = KX_INPUT
                    ws.cell(r, 14, fila.igv_pct).font = KX_INPUT
                    if fila.tipo == COMPRA and fila.flete:
                        ws.cell(r, 15, fila.flete).font = KX_INPUT
                        ws.cell(r, 5, f"=(D{r}*M{r}/(1+N{r})+O{r})/D{r}").font = KX_NORMAL
                    else:
                        ws.cell(r, 5, f"=M{r}/(1+N{r})").font = KX_NORMAL
                    if fila.tipo == COMPRA and fila.documento:
                        filas_compra[fila.documento] = r
                else:
                    ws.cell(r, 5, fila.ing_costo_unitario).font = KX_INPUT
                ws.cell(r, 6, f"=D{r}*E{r}").font = KX_NORMAL
            elif fila.tipo in (CONSUMO, DEVOLUCION_CONSUMO):
                ws.cell(r, 7, fila.sal_cantidad).font = KX_INPUT
                ws.cell(r, 8, f"=K{prev}").font = KX_NORMAL
                ws.cell(r, 9, f"=G{r}*H{r}").font = KX_NORMAL
            ws.cell(r, 10, f"=J{prev}+N(D{r})-N(G{r})").font = KX_NORMAL
            ws.cell(r, 12, f"=L{prev}+N(F{r})-N(I{r})").font = KX_NORMAL
            ws.cell(r, 11, f"=IF(ROUND(J{r},6)=0,0,L{r}/J{r})").font = KX_NORMAL
        r += 1
    last = r - 1
    for col in (10, 11, 12):
        ws.cell(last, col).fill = KX_YELLOW
    total_row = r
    ws.cell(total_row, 3, "TOTALES DEL PERIODO").font = KX_BOLD
    for col, letra in ((4, "D"), (6, "F"), (7, "G"), (9, "I")):
        c = ws.cell(total_row, col, f"=SUM({letra}{first}:{letra}{last})")
        c.font, c.fill = KX_BOLD, KX_YELLOW
    for row in ws.iter_rows(min_row=first, max_row=total_row, min_col=4, max_col=12):
        for c in row:
            c.number_format = KX_CU if c.column in (5, 8, 11) else KX_NUM
    for row in ws.iter_rows(min_row=first, max_row=last, min_col=13, max_col=15):
        row[0].number_format = KX_CU
        row[1].number_format = "0%"
        row[2].number_format = KX_NUM

    # Comprobacion
    c0 = total_row + 2
    ws.cell(c0, 4, "COMPROBACIÓN").font = KX_BOLD
    for col, titulo in ((4, "Concepto"), (5, "Cantidad"), (6, "Costo total S/")):
        c = ws.cell(c0 + 1, col, titulo)
        c.font, c.fill, c.alignment = KX_BOLD, KX_GROUP_FILL, KX_CENTER
    filas = [
        ("Saldo inicial", f"=J{first}", f"=L{first}"),
        ("Compras", f"=D{total_row}", f"=F{total_row}"),
        ("Disponible", f"=E{c0 + 2}+E{c0 + 3}", f"=F{c0 + 2}+F{c0 + 3}"),
        ("Consumos", f"=G{total_row}", f"=I{total_row}"),
        ("Saldo final", f"=E{c0 + 4}-E{c0 + 5}", f"=F{c0 + 4}-F{c0 + 5}"),
        ("Diferencia", f"=ROUND(E{c0 + 6}-J{last},2)", f"=ROUND(F{c0 + 6}-L{last},2)"),
    ]
    for i, (concepto, q, t) in enumerate(filas, start=c0 + 2):
        ws.cell(i, 4, concepto).font = KX_BOLD if concepto == "Saldo final" else KX_NORMAL
        for col, formula in ((5, q), (6, t)):
            c = ws.cell(i, col, formula)
            c.font = KX_BOLD if concepto == "Saldo final" else KX_NORMAL
            c.number_format = KX_NUM
    nota = c0 + 9
    conc = None
    if cmd:
        k0 = c0 + 9
        ws.cell(k0, 4, f"CONCILIACIÓN CON EL COSTO DE MATERIAL DIRECTO — {cmd['codigo_ot']}").font = KX_BOLD
        for col, titulo in ((4, "Concepto"), (5, "Cantidad kg"), (6, "Cto Unit S/"), (7, "Costo total S/")):
            c = ws.cell(k0 + 1, col, titulo)
            c.font, c.fill, c.alignment = KX_BOLD, KX_GROUP_FILL, KX_CENTER
        ws.cell(k0 + 2, 4, "Salidas del kardex").font = KX_NORMAL
        ws.cell(k0 + 2, 5, f"=G{total_row}")
        ws.cell(k0 + 2, 6, f"=IFERROR(G{k0 + 2}/E{k0 + 2},0)")
        ws.cell(k0 + 2, 7, f"=I{total_row}")
        ws.cell(k0 + 3, 4, "CMD de la OT (calculado por el sistema)").font = KX_NORMAL
        ws.cell(k0 + 3, 5, cmd["cmd_cantidad"]).font = KX_INPUT
        ws.cell(k0 + 3, 6, cmd["costo_unitario"]).font = KX_INPUT
        ws.cell(k0 + 3, 7, f"=E{k0 + 3}*F{k0 + 3}")
        ws.cell(k0 + 4, 4, "Diferencia").font = KX_BOLD
        ws.cell(k0 + 4, 5, f"=ROUND(E{k0 + 2}-E{k0 + 3},2)")
        ws.cell(k0 + 4, 7, f"=ROUND(G{k0 + 2}-G{k0 + 3},2)")
        for i in range(k0 + 2, k0 + 5):
            for col in (5, 6, 7):
                c = ws.cell(i, col)
                if isinstance(c.value, str) and c.value.startswith("="):
                    c.font = KX_BOLD if i == k0 + 4 else KX_NORMAL
                c.number_format = KX_CU if col == 6 else KX_NUM
        conc = k0 + 2
        nota = k0 + 6
    ws.cell(nota, 1, "Notas:").font = KX_BOLD
    notas = [
        "Azul = dato de la hoja Compras (o CMD calculado en la conciliación); negro = fórmula.",
        "Cto Unit de compra = (cantidad × precio con IGV / (1 + IGV) + flete) / cantidad.",
        "Salidas a las OT y devoluciones de consumo se valorizan al costo promedio vigente (Cto Unit del saldo anterior).",
        "Devolución a proveedor: ingreso negativo al costo unitario de la compra devuelta.",
        "Las salidas del kardex a cada OT son su costo de material directo (CMD): la hoja de costos toma el CMD de aquí.",
    ]
    for i, texto in enumerate(notas, start=nota + 1):
        ws.cell(i, 1, texto).font = KX_NORMAL
    if item.avisos:
        ws.cell(nota + len(notas) + 2, 1, "Advertencias de la hoja Compras:").font = KX_BOLD
        for i, aviso in enumerate(item.avisos, start=nota + len(notas) + 3):
            ws.cell(i, 1, aviso).font = Font(name=KX_FONT, size=10, color="FF0000")

    for col, ancho in zip("ABCDEFGHIJKLMNO", [16, 12, 58, 13, 12, 14, 11, 12, 14, 11, 12, 14, 16, 8, 11]):
        ws.column_dimensions[col].width = ancho
    ws.freeze_panes = "D10"
    return {"first": first, "last": last, "total": total_row, "comprobacion": c0 + 2, "conciliacion": conc}


def generate_kardex_report(db_path: str | Path, output_dir: str | Path = REPORTS_DIR) -> Path:
    """Reporte_Kardex_Promedio.xlsx: un kardex por item de la hoja Compras + resumen."""
    from .kardex_compras import conciliacion_cmd, kardex_desde_base

    with db_session(db_path) as conn:
        items = kardex_desde_base(conn)
        conciliacion = {c["material_id"]: c for c in conciliacion_cmd(conn, items)}
        for c in conciliacion.values():
            c["costo_unitario"] = c["cmd_soles"] / c["cmd_cantidad"] if c["cmd_cantidad"] else 0
        periodos = {row["periodo_id"]: row["periodo"] for row in conn.execute("SELECT periodo_id, periodo FROM periodo")}
    wb = Workbook()
    resumen = wb.active
    resumen.title = "Resumen"
    resumen["A1"] = "KARDEX DE MATERIALES DIRECTOS POR PROMEDIO PONDERADO — RESUMEN"
    resumen["A1"].font = KX_TITLE
    if not items:
        resumen["A3"] = "No hay datos en la hoja Compras del Excel maestro."
        return _save_workbook(wb, Path(output_dir) / "Reporte_Kardex_Promedio.xlsx")
    headers = ["Código", "Descripción", "Unidad", "Saldo inicial cant.", "Saldo inicial S/", "Compras cant.", "Compras S/",
               "Salidas a OT cant.", "Salidas a OT S/", "Saldo final cant.", "Cto Unit final S/", "Saldo final S/", "Comprobación S/",
               "OT", "CMD S/", "Diferencia vs CMD S/"]
    for col, titulo in enumerate(headers, start=1):
        c = resumen.cell(3, col, titulo)
        c.font, c.fill, c.alignment = KX_BOLD, KX_GROUP_FILL, KX_CENTER
    usados: set[str] = {"Resumen"}
    fila = 4
    for item in items:
        nombre = _kardex_sheet_name(item.item_codigo, usados)
        ws = wb.create_sheet(nombre)
        cmd = conciliacion.get(item.item_codigo)
        pos = _write_kardex_item(ws, item, periodos.get(item.periodo_id, item.periodo_id), cmd)
        f, t, l, c0 = pos["first"], pos["total"], pos["last"], pos["comprobacion"]
        k = pos["conciliacion"]
        valores = [item.item_codigo, item.item_descripcion, item.unidad,
                   "=" + _kx_ref(nombre, f"J{f}"), "=" + _kx_ref(nombre, f"L{f}"),
                   "=" + _kx_ref(nombre, f"D{t}"), "=" + _kx_ref(nombre, f"F{t}"),
                   "=" + _kx_ref(nombre, f"G{t}"), "=" + _kx_ref(nombre, f"I{t}"),
                   "=" + _kx_ref(nombre, f"J{l}"), "=" + _kx_ref(nombre, f"K{l}"), "=" + _kx_ref(nombre, f"L{l}"),
                   "=" + _kx_ref(nombre, f"F{c0 + 4}")]
        if k:
            valores += [cmd["codigo_ot"], "=" + _kx_ref(nombre, f"G{k + 1}"), "=" + _kx_ref(nombre, f"G{k + 2}")]
        for col, valor in enumerate(valores, start=1):
            c = resumen.cell(fila, col, valor)
            c.font = KX_LINK if str(valor).startswith("=") else KX_NORMAL
            if col >= 4 and col != 14:
                c.number_format = KX_CU if col == 11 else KX_NUM
        fila += 1
    resumen.cell(fila, 2, "TOTAL").font = KX_BOLD
    for col in (5, 7, 9, 12, 13, 15, 16):
        letra = get_column_letter(col)
        c = resumen.cell(fila, col, f"=SUM({letra}4:{letra}{fila - 1})")
        c.font, c.number_format, c.fill = KX_BOLD, KX_NUM, KX_YELLOW
    resumen.cell(fila + 2, 1, "Comprobación S/ = saldo inicial + compras − consumos (debe igualar al saldo final del kardex).").font = KX_NORMAL
    resumen.cell(fila + 3, 1, "Salidas a OT S/ = consumo de material directo; Diferencia vs CMD debe ser 0,00 "
                 "(el CMD de la hoja de costos sale de las salidas del kardex).").font = KX_NORMAL
    resumen.cell(fila + 4, 1, "Fuente: hoja Compras del Excel maestro (datos ficticios del almacén de materias primas).").font = KX_NORMAL
    for col, ancho in zip("ABCDEFGHIJKLMNOP", [12, 30, 8, 14, 14, 12, 14, 14, 14, 14, 14, 14, 15, 16, 14, 14]):
        resumen.column_dimensions[col].width = ancho
    resumen.freeze_panes = "A4"
    return _save_workbook(wb, Path(output_dir) / "Reporte_Kardex_Promedio.xlsx")
