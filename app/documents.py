from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor

from .constants import OUTPUT_DOCUMENTS, REPORTS_DIR, SOURCE_ROOT_DIR
from .dashboard import generate_dashboard_image
from .database import db_session, latest_run_id, quote_identifier
from .utils import etiqueta_objeto_costo


HEADING_BLUE = RGBColor(46, 116, 181)
INK_BLUE = RGBColor(31, 77, 120)


def _setup_document(title: str, subtitle: str = "") -> Document:
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    styles = doc.styles
    styles["Normal"].font.name = "Calibri"
    styles["Normal"].font.size = Pt(11)
    for style_name, size, color in [
        ("Heading 1", 16, HEADING_BLUE),
        ("Heading 2", 13, HEADING_BLUE),
        ("Heading 3", 12, INK_BLUE),
    ]:
        style = styles[style_name]
        style.font.name = "Calibri"
        style.font.size = Pt(size)
        style.font.color.rgb = color
        style.font.bold = True
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(title)
    run.bold = True
    run.font.size = Pt(18)
    run.font.color.rgb = HEADING_BLUE
    if subtitle:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run(subtitle).italic = True
    doc.add_paragraph()
    return doc


def _add_key_value_table(doc: Document, rows: list[tuple[str, str]]) -> None:
    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    table.columns[0].width = Inches(1.8)
    table.columns[1].width = Inches(4.7)
    hdr = table.rows[0].cells
    hdr[0].text = "Campo"
    hdr[1].text = "Detalle"
    for label, value in rows:
        cells = table.add_row().cells
        cells[0].text = str(label)
        cells[1].text = str(value)


def _rows(conn, table: str, run_id: int | None = None) -> list[dict]:
    sql = f"SELECT * FROM {quote_identifier(table)}"
    params = []
    if run_id is not None and any(r["name"] == "run_id" for r in conn.execute(f"PRAGMA table_info({quote_identifier(table)})")):
        sql += " WHERE run_id = ?"
        params.append(run_id)
    return [dict(row) for row in conn.execute(sql, params).fetchall()]


def _add_small_table(doc: Document, title: str, rows: list[dict], limit: int = 8) -> None:
    doc.add_heading(title, level=2)
    if not rows:
        doc.add_paragraph("Sin datos disponibles.")
        return
    headers = list(rows[0].keys())[:6]
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    for idx, header in enumerate(headers):
        table.rows[0].cells[idx].text = header
    for row in rows[:limit]:
        cells = table.add_row().cells
        for idx, header in enumerate(headers):
            cells[idx].text = "" if row.get(header) is None else str(row.get(header))


def _safe_number(row: dict | None, field: str) -> float:
    if not row:
        return 0.0
    try:
        return float(row.get(field) or 0)
    except (TypeError, ValueError):
        return 0.0


def _fmt_money(value: float) -> str:
    return f"S/ {value:,.2f}"


def _fmt_pct(value: float) -> str:
    return f"{value * 100:,.2f}%"


def _add_dashboard_picture(
    doc: Document,
    db_path: str | Path,
    output_dir: str | Path,
    title: str = "Dashboard visual",
    payload_override: dict | None = None,
) -> None:
    doc.add_heading(title, level=2)
    path = generate_dashboard_image(db_path, output_dir, payload_override)
    if path and path.exists():
        doc.add_picture(str(path), width=Inches(6.4))
        last = doc.paragraphs[-1]
        last.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("Figura generada desde los resultados calculados del Modelo ABC por Ordenes de Trabajo.")
    else:
        doc.add_paragraph("No se pudo generar la imagen del dashboard para esta ejecucion.")


def _add_executive_metrics(doc: Document, costs: list[dict], results: list[dict], activities: list[dict]) -> None:
    total_cost = sum(_safe_number(row, "costo_total_ot") for row in costs)
    total_sales = sum(_safe_number(row, "ventas_netas_soles") for row in results)
    avg_margin = (
        sum(_safe_number(row, "rentabilidad_operativa") for row in results) / len(results)
        if results
        else 0.0
    )
    top_activity = max(activities, key=lambda row: _safe_number(row, "costo_total_ac"), default={})
    _add_key_value_table(
        doc,
        [
            ("Costo total OT", _fmt_money(total_cost)),
            ("Ventas netas", _fmt_money(total_sales)),
            ("Rentabilidad promedio", _fmt_pct(avg_margin)),
            ("Actividad de mayor costo", str(top_activity.get("nombre_actividad") or "Sin datos")),
        ],
    )


def _save(doc: Document, output_dir: Path, filename: str) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    path = output_dir / filename
    doc.save(path)
    return path


def generate_manual_descriptivo(db_path: str | Path, output_dir: str | Path = REPORTS_DIR) -> Path:
    doc = _setup_document("Manual descriptivo", "Sistema de Costos ABC por Ordenes de Trabajo - Frenosa")
    doc.add_heading("Objetivo", level=1)
    doc.add_paragraph(
        "El aplicativo permite importar un Excel maestro, validar su estructura, cargar datos a SQLite, "
        "ejecutar el Modelo ABC por Ordenes de Trabajo y generar reportes, dashboard y documentos academicos."
    )
    doc.add_heading("Arquitectura funcional", level=1)
    _add_key_value_table(
        doc,
        [
            ("Entrada", "Excel maestro compatible con las hojas obligatorias definidas en AGENTS.md."),
            ("Base de datos", "SQLite local con tablas de datos importados y resultados generados."),
            ("Modelo disponible", "Modelo ABC por Ordenes de Trabajo: Proceso 1 EC -> AC, Proceso 2 AC-OT (Produccion) y Proceso 2 AC-OT (Soporte)."),
            ("Modelos bloqueados", "Modelos 2 y 3 quedan visibles, pero no ejecutables por datos faltantes."),
            ("Salidas", "Reportes Excel, documentos Word, dashboard y pruebas funcionales."),
        ],
    )
    doc.add_heading("Salidas generadas", level=1)
    doc.add_paragraph(
        "Los reportes Excel incluyen hojas de resumen, formulas, datos detallados y graficos para revisar costos por OT, "
        "estructura de costos, actividades relevantes y rentabilidad. Los documentos Word resumen la metodologia, "
        "los hallazgos principales y las propuestas de mejora sustentadas en los calculos."
    )
    return _save(doc, Path(output_dir), "Manual_descriptivo.docx")


def generate_manual_usuario(db_path: str | Path, output_dir: str | Path = REPORTS_DIR) -> Path:
    doc = _setup_document("Manual de usuario", "Uso operativo del aplicativo")
    doc.add_heading("Flujo de trabajo", level=1)
    steps = [
        "Abrir el aplicativo desde Python o desde el ejecutable compilado (empieza sin archivos cargados).",
        "Seleccionar uno o varios Excel maestros (un mes cada uno); el sistema identifica el mes por la hoja Periodo.",
        "Validar e importar los datos a SQLite local: cada mes se importa y se calcula.",
        "Elegir el tipo de analisis (mensual, bimestral, trimestral, cuatrimestral, semestral o anual, segun los meses "
        "consecutivos cargados) y el periodo.",
        "Seleccionar Modelo 1 - Modelo ABC por Ordenes de Trabajo y ejecutar el calculo.",
        "Revisar Proceso 1, Proceso 2 AC-OT (Produccion), Proceso 2 AC-OT (Soporte), hoja de costos, estado de resultados, dashboard y evolucion.",
        "Simular: clic en una celda con lapiz, nuevo valor, 'Ejecutar nuevos calculos'; 'Reestablecer valores iniciales' descarta los cambios.",
        "Generar reportes Excel y documentos Word.",
    ]
    for item in steps:
        doc.add_paragraph(item, style="List Number")
    doc.add_heading("Varios meses y simulacion", level=1)
    for item in [
        "Con varios meses, las pestañas muestran el consolidado del periodo: importes sumados; costo unitario y rentabilidad "
        "sobre los totales; las OT de cada mes se agrupan por producto. 'Ver calculos' muestra la fila mes por mes.",
        "La simulacion permite cambiar datos importados (precios, cantidades, salarios, horas, importes), los costos de cada "
        "actividad, los pesos AC-OT, la cantidad producida, el CMD/CSD y las ventas de cada OT, y comparar con los valores iniciales.",
        "El CMD de cada OT sale del kardex de materiales directos (hoja Compras, promedio ponderado).",
    ]:
        doc.add_paragraph(item, style="List Bullet")
    doc.add_heading("Modelos de costeo", level=1)
    doc.add_paragraph("Modelo 1 esta disponible cuando el Excel contiene actividades, inductores EC-AC y AC-OT.")
    doc.add_paragraph("Modelo 2 y Modelo 3 se muestran como opciones teoricas bloqueadas por ausencia de datos.")
    doc.add_heading("Lectura del dashboard", level=1)
    for item in [
        "Revisar primero las tarjetas KPI para ubicar costo total, ventas y rentabilidad.",
        "Usar el grafico de estructura de costo para identificar si domina CMD, produccion, soporte o CSD.",
        "Usar el ranking de actividades para priorizar acciones de eficiencia.",
        "Leer las alertas y recomendaciones antes de exportar informes finales.",
    ]:
        doc.add_paragraph(item, style="List Bullet")
    return _save(doc, Path(output_dir), "Manual_usuario.docx")


def generate_informe_pruebas(db_path: str | Path, test_results: list[dict] | None = None, output_dir: str | Path = REPORTS_DIR) -> Path:
    doc = _setup_document("Informe de pruebas", "Validacion funcional del sistema")
    doc.add_heading("Resumen", level=1)
    if test_results:
        passed = sum(1 for item in test_results if item.get("ok"))
        doc.add_paragraph(f"Pruebas ejecutadas: {len(test_results)}. Pruebas conformes: {passed}.")
        _add_small_table(doc, "Detalle de pruebas", test_results, limit=50)
    else:
        doc.add_paragraph("El informe se genera desde el modulo de pruebas del aplicativo.")
    doc.add_heading("Verificacion de salidas", level=1)
    for item in [
        "Los reportes Excel deben abrir con hojas de resumen, formulas y graficos visibles.",
        "Los documentos Word deben incluir tablas de indicadores y, cuando existan datos calculados, una imagen del dashboard.",
        "El dashboard PNG debe reflejar costos, actividades, rentabilidad y alertas.",
    ]:
        doc.add_paragraph(item, style="List Bullet")
    return _save(doc, Path(output_dir), "Informe_pruebas.docx")


def generate_informe_mejoras(
    db_path: str | Path,
    output_dir: str | Path = REPORTS_DIR,
    payload_override: dict | None = None,
) -> Path:
    doc = _setup_document("Informe de propuestas de mejora", "Impacto positivo en la rentabilidad")
    if payload_override is not None:
        activities = payload_override.get("costos_actividad", [])
        results = payload_override.get("estado_resultados", [])
        costs = payload_override.get("costos_ot", [])
    else:
        with db_session(db_path) as conn:
            run_id = latest_run_id(conn)
            activities = _rows(conn, "actividad_total", run_id)
            results = _rows(conn, "estado_resultados_ot", run_id)
            costs = _rows(conn, "hoja_costos_ot", run_id)
            codigos = {r.get("ot_id"): r.get("codigo_ot") for r in costs}
            for r in results:
                r.setdefault("codigo_ot", codigos.get(r.get("ot_id")))
    top_activity = max(activities, key=lambda r: r.get("costo_total_ac") or 0, default=None)
    low_margin = min(results, key=lambda r: r.get("rentabilidad_operativa") or 0, default=None)
    doc.add_heading("Resumen ejecutivo", level=1)
    if payload_override and payload_override.get("modo_periodo") == "PROYECTADO":
        doc.add_paragraph(
            f"Vista estimada/proyectada para periodo manual con factor {payload_override.get('factor_proyeccion', 1):g}. "
            "Se usa como referencia la estructura del Excel maestro importado sin alterar los calculos originales."
        )
    _add_executive_metrics(doc, costs, results, activities)
    _add_dashboard_picture(doc, db_path, output_dir, payload_override=payload_override)
    doc.add_heading("Hallazgos ABC", level=1)
    if top_activity:
        doc.add_paragraph(
            f"La actividad con mayor costo total es {top_activity.get('nombre_actividad')}, "
            f"con S/ {top_activity.get('costo_total_ac', 0):,.2f}."
        )
    if low_margin:
        doc.add_paragraph(
            f"El objeto de costo (OT) con menor rentabilidad operativa es {etiqueta_objeto_costo(low_margin)}, "
            f"con {low_margin.get('rentabilidad_operativa', 0) * 100:,.2f}%."
        )
    _add_small_table(doc, "Hoja de costos OT", costs, limit=6)
    _add_small_table(doc, "Estado de resultados OT", results, limit=6)
    doc.add_heading("Propuestas", level=1)
    for item in [
        "Revisar actividades con mayor consumo de MOD y otros recursos indirectos.",
        "Analizar mantenimiento de equipos asociados a actividades de alto costo.",
        "Mejorar la precision de inductores EC-AC y AC-OT para decisiones comerciales.",
        "Evaluar precios y mezcla de las OT (objetos de costo) con menor margen operativo.",
        "Separar actividades que agregan valor y actividades de soporte para focalizar mejoras.",
    ]:
        doc.add_paragraph(item, style="List Bullet")
    return _save(doc, Path(output_dir), "Informe_propuestas_mejora.docx")


def generate_analisis_funcional(db_path: str | Path, output_dir: str | Path = REPORTS_DIR) -> Path:
    doc = _setup_document("Analisis funcional y diseno del sistema", "Monografia 01 GE605U")
    doc.add_heading("Alcance", level=1)
    doc.add_paragraph(
        "Sistema informatico desktop offline para costos basados en actividades por ordenes de trabajo, "
        "con importacion Excel, SQLite, calculos ABC, reportes, dashboard, documentos y pruebas."
    )
    doc.add_heading("Modulos", level=1)
    for module in [
        "Pantalla inicial e importacion Excel.",
        "Modelos de costeo con disponibilidad validada.",
        "Maestros: empresa, periodo, OT (objetos de costo), recursos, actividades e inductores.",
        "Proceso 1: EC -> AC.",
        "Proceso 2 AC-OT (Produccion).",
        "Proceso 2 AC-OT (Soporte).",
        "Hoja de costos, estado de resultados y rentabilidad.",
        "Dashboard, reportes Excel, documentos Word y pruebas.",
        "Carga de varios meses y analisis mensual, bimestral, trimestral, cuatrimestral, semestral o anual.",
        "Simulacion de valores (datos y resultados) con recalculo y comparacion con los valores iniciales.",
    ]:
        doc.add_paragraph(module, style="List Bullet")
    doc.add_heading("Diseno de base de datos", level=1)
    _add_key_value_table(
        doc,
        [
            ("Datos importados", "Tablas espejo del Excel maestro con batch_id."),
            ("Calculos generados", "Tablas por proceso: MOD, EC-AC, soporte, AC-OT, hoja de costos y estado de resultados."),
            ("Trazabilidad", "source_trace, import_batches, import_errors y calculation_runs."),
        ],
    )
    doc.add_heading("Reportes y graficos", level=1)
    doc.add_paragraph(
        "El modulo de reportes consolida salidas en Excel con hojas de resumen, graficos de barras, graficos apilados "
        "y ranking de actividades. El dashboard PNG y los documentos Word usan los mismos resultados calculados para "
        "mantener trazabilidad entre pantalla, reportes y documentos."
    )
    _add_dashboard_picture(doc, db_path, output_dir, "Vista grafica de referencia")
    doc.add_heading("Casos de uso", level=1)
    for case in [
        "Importar Excel maestro.",
        "Validar modelos disponibles.",
        "Ejecutar Modelo ABC por Ordenes de Trabajo.",
        "Consultar hoja de costos y dashboard.",
        "Exportar reportes y documentos.",
        "Ejecutar pruebas funcionales.",
    ]:
        doc.add_paragraph(case, style="List Number")
    return _save(doc, Path(output_dir), "Analisis_funcional_diseno_sistema.docx")


def generate_codigo_principal(db_path: str | Path, output_dir: str | Path = REPORTS_DIR) -> Path:
    doc = _setup_document("Codigos de programacion principales", "Extractos de modulos del aplicativo")
    modules = [
        "app/importers.py",
        "app/costing_mod.py",
        "app/costing_abc.py",
        "app/analisis_periodos.py",
        "app/reports.py",
        "app/documents.py",
        "app/tests_runner.py",
    ]
    for rel in modules:
        path = SOURCE_ROOT_DIR / rel
        doc.add_heading(rel, level=1)
        if path.exists():
            text = path.read_text(encoding="utf-8", errors="ignore")
            excerpt = text[:5000]
            paragraph = doc.add_paragraph()
            run = paragraph.add_run(excerpt)
            run.font.name = "Consolas"
            run.font.size = Pt(8)
        else:
            doc.add_paragraph("Archivo no encontrado.")
    return _save(doc, Path(output_dir), "Codigos_programacion_principales.docx")


def generate_all_documents(
    db_path: str | Path,
    output_dir: str | Path = REPORTS_DIR,
    test_results: list[dict] | None = None,
    dashboard_payload_override: dict | None = None,
) -> list[Path]:
    return [
        generate_manual_descriptivo(db_path, output_dir),
        generate_manual_usuario(db_path, output_dir),
        generate_informe_pruebas(db_path, test_results, output_dir),
        generate_informe_mejoras(db_path, output_dir, dashboard_payload_override),
        generate_analisis_funcional(db_path, output_dir),
        generate_codigo_principal(db_path, output_dir),
    ]
