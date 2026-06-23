from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from ..constants import (
    APP_TITLE,
    DEFAULT_DB_PATH,
    REPORTS_DIR,
)
from ..costing_abc import run_model_1_calculation
from ..dashboard import dashboard_payload, generate_dashboard_image
from ..database import db_session, initialize_database, latest_run_id
from ..database import quote_identifier
from ..documents import generate_all_documents
from ..importers import import_excel_to_sqlite, validate_workbook
from ..model_availability import CostingModelStatus, evaluate_models
from ..reports import generate_all_reports


IMPORTED_TABLES = [
    "empresas",
    "periodos",
    "modelos_costeo",
    "objetos_costo",
    "ordenes_trabajo",
    "materiales_directos",
    "Personal_directos",
    "servicios_directos",
    "costo_otros_recursos(CI)",
    "actividades",
    "inductores",
    "ec_ac_inductores",
    "Soporte_Mantenimiento(CI)",
    "Soporte_Control_Calidad(CI)",
    "Soporte_Ingenieria(CI)",
    "soporte_cc_ing_inductores",
    "ac_ot_pesos_produccion",
    "ac_ot_pesos_soporte",
    "ventas",
    "reglas_calculo",
    "fuentes",
    "diccionario",
]


class CostosABCApp(tk.Tk):
    def __init__(self, db_path: str | Path = DEFAULT_DB_PATH):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1280x780")
        self.minsize(1080, 680)
        self.db_path = Path(db_path)
        self.excel_path = tk.StringVar(value="")
        self.status_text = tk.StringVar(value="Seleccione, valide e importe un Excel maestro para iniciar.")
        self.model_var = tk.StringVar(value="")
        self.table_var = tk.StringVar(value=IMPORTED_TABLES[0])
        self.model_statuses: list[CostingModelStatus] = []
        self.trees: dict[str, ttk.Treeview] = {}
        self.module_tabs: list[tuple[ttk.Frame, str]] = []
        self.is_excel_validated = False
        self.is_excel_imported = False
        self.validated_excel_path: str | None = None
        self.imported_excel_path: str | None = None
        self._configure_style()
        self._build_layout()
        initialize_database(self.db_path)
        self.reset_model_selector()

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TButton", padding=6)
        style.configure("Title.TLabel", font=("Segoe UI", 16, "bold"), foreground="#1F4D78")
        style.configure("Subtitle.TLabel", font=("Segoe UI", 10), foreground="#334155")
        style.configure("Section.TLabelframe.Label", font=("Segoe UI", 10, "bold"))
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))
        style.configure("Treeview", rowheight=24)

    def _build_layout(self) -> None:
        root = ttk.Frame(self, padding=12)
        root.pack(fill=tk.BOTH, expand=True)

        header = ttk.Frame(root)
        header.pack(fill=tk.X)
        ttk.Label(header, text="Sistema Informatico de Costos ABC - Frenosa", style="Title.TLabel").pack(anchor=tk.W)
        ttk.Label(
            header,
            text="Modelo ABC por Ordenes de Trabajo | Datos importados, calculos generados y reportes exportados.",
            style="Subtitle.TLabel",
        ).pack(anchor=tk.W, pady=(2, 8))

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.home_tab = self._build_home_tab()
        self._build_data_tab()
        self._build_table_tab("Proceso 1 EC-AC", "proceso1")
        self._build_table_tab("Proceso 2 AC-OT (Produccion)", "proceso2_produccion")
        self._build_table_tab("Proceso 2 AC-OT (Soporte)", "proceso2_soporte")
        self._build_table_tab("Hoja de costos OT", "hoja")
        self._build_table_tab("Estado de resultados", "estado")
        self._build_dashboard_tab()
        self._build_outputs_tab()
        self.hide_module_tabs()

        status = ttk.Label(root, textvariable=self.status_text, anchor=tk.W)
        status.pack(fill=tk.X, pady=(8, 0))

    def _build_home_tab(self) -> ttk.Frame:
        tab = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(tab, text="Pantalla inicial")

        import_box = ttk.LabelFrame(tab, text="Importar Excel maestro", padding=10, style="Section.TLabelframe")
        import_box.pack(fill=tk.X)
        ttk.Label(import_box, text="Archivo Excel maestro:").grid(row=0, column=0, sticky=tk.W)
        entry = ttk.Entry(import_box, textvariable=self.excel_path)
        entry.grid(row=0, column=1, sticky=tk.EW, padx=8)
        import_box.columnconfigure(1, weight=1)
        ttk.Button(import_box, text="Seleccionar", command=self.choose_excel).grid(row=0, column=2, padx=4)
        ttk.Button(import_box, text="Validar estructura", command=self.validate_excel).grid(row=0, column=3, padx=4)
        ttk.Button(import_box, text="Importar a SQLite", command=self.import_excel).grid(row=0, column=4, padx=4)

        model_box = ttk.LabelFrame(tab, text="Seleccion de modelo de costeo", padding=10, style="Section.TLabelframe")
        model_box.pack(fill=tk.X, pady=12)
        self.model_combo = ttk.Combobox(model_box, state="readonly", textvariable=self.model_var, width=72)
        self.model_combo.grid(row=0, column=0, sticky=tk.W)
        self.model_combo.bind("<<ComboboxSelected>>", lambda _event: self.show_selected_model_reason())
        self.execute_button = ttk.Button(model_box, text="Ejecutar calculo", command=self.execute_calculation)
        self.execute_button.grid(row=0, column=1, padx=8)
        self.model_reason = tk.Text(model_box, height=5, wrap=tk.WORD)
        self.model_reason.grid(row=1, column=0, columnspan=2, sticky=tk.EW, pady=(8, 0))
        model_box.columnconfigure(0, weight=1)

        note = ttk.LabelFrame(tab, text="Criterio academico de modelos", padding=10, style="Section.TLabelframe")
        note.pack(fill=tk.BOTH, expand=True, pady=12)
        self.criteria_text = tk.Text(note, wrap=tk.WORD, height=10)
        self.criteria_text.pack(fill=tk.BOTH, expand=True)
        self.set_criteria_text(
            "Seleccione un Excel maestro, valide su estructura e importelo a SQLite.\n\n"
            "La disponibilidad de Modelo 1, Modelo 2 y Modelo 3 se evaluara recien despues de la importacion."
        )
        return tab

    def _build_data_tab(self) -> None:
        tab = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(tab, text="Datos importados")
        self.module_tabs.append((tab, "Datos importados"))
        top = ttk.Frame(tab)
        top.pack(fill=tk.X)
        ttk.Button(top, text="Volver a pantalla inicial", command=self.show_home_screen).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(top, text="Tabla:").pack(side=tk.LEFT)
        combo = ttk.Combobox(top, textvariable=self.table_var, values=IMPORTED_TABLES, state="readonly", width=42)
        combo.pack(side=tk.LEFT, padx=8)
        combo.bind("<<ComboboxSelected>>", lambda _event: self.load_selected_imported_table())
        ttk.Button(top, text="Actualizar", command=self.load_selected_imported_table).pack(side=tk.LEFT)
        tree = self._make_tree(tab)
        self.trees["datos"] = tree

    def _build_table_tab(self, title: str, key: str) -> None:
        tab = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(tab, text=title)
        self.module_tabs.append((tab, title))
        ttk.Label(tab, text=title, style="Subtitle.TLabel").pack(anchor=tk.W)
        self.trees[key] = self._make_tree(tab)

    def _build_dashboard_tab(self) -> None:
        tab = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(tab, text="Dashboard")
        self.module_tabs.append((tab, "Dashboard"))
        self.dashboard_text = tk.Text(tab, wrap=tk.WORD)
        self.dashboard_text.pack(fill=tk.BOTH, expand=True)

    def _build_outputs_tab(self) -> None:
        tab = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(tab, text="Reportes y documentos")
        self.module_tabs.append((tab, "Reportes y documentos"))
        ttk.Label(tab, text=f"Reportes y documentos se guardan en: {REPORTS_DIR}", style="Subtitle.TLabel").pack(anchor=tk.W)
        actions = ttk.Frame(tab)
        actions.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(actions, text="Generar reportes Excel", command=self.generate_reports).pack(side=tk.LEFT, padx=4)
        ttk.Button(actions, text="Generar documentos Word", command=self.generate_documents).pack(side=tk.LEFT, padx=4)
        ttk.Button(actions, text="Generar dashboard PNG", command=self.generate_dashboard_png).pack(side=tk.LEFT, padx=4)
        ttk.Button(actions, text="Ejecutar informe de pruebas", command=self.run_tests).pack(side=tk.LEFT, padx=4)
        self.trees["salidas"] = self._make_tree(tab)

    def _make_tree(self, parent) -> ttk.Treeview:
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        tree = ttk.Treeview(frame, show="headings")
        yscroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
        xscroll = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        return tree

    def choose_excel(self) -> None:
        path = filedialog.askopenfilename(
            title="Seleccionar Excel maestro",
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")],
        )
        if path:
            self.excel_path.set(path)
            self.is_excel_validated = False
            self.is_excel_imported = False
            self.validated_excel_path = None
            self.imported_excel_path = None
            self.reset_model_selector()
            self.show_home_screen()
            self.status_text.set("Excel seleccionado. Valide la estructura antes de importar.")

    def validate_excel(self) -> None:
        if not self.excel_path.get().strip():
            messagebox.showwarning("Validacion", "Primero seleccione un Excel maestro.")
            self.status_text.set("No hay Excel seleccionado.")
            return
        result = validate_workbook(self.excel_path.get())
        if result.ok:
            self.is_excel_validated = True
            self.validated_excel_path = self.excel_path.get()
            messagebox.showinfo("Validacion", "La estructura del Excel es valida.")
        else:
            self.is_excel_validated = False
            self.validated_excel_path = None
            messagebox.showerror("Validacion", "\n".join(issue.message for issue in result.errors[:8]))
        self.status_text.set(f"Validacion: {len(result.errors)} errores, {len(result.warnings)} advertencias.")

    def import_excel(self) -> None:
        if not self.excel_path.get().strip():
            messagebox.showwarning("Importacion", "Primero seleccione un Excel maestro.")
            self.status_text.set("No hay Excel seleccionado.")
            return
        if not self.is_excel_validated or self.validated_excel_path != self.excel_path.get():
            messagebox.showwarning(
                "Importacion",
                "Primero valide la estructura del Excel seleccionado. Luego importelo a SQLite.",
            )
            self.status_text.set("Debe validar el Excel actual antes de importar.")
            return
        try:
            result = import_excel_to_sqlite(self.excel_path.get(), self.db_path)
            if result.ok:
                self.is_excel_imported = True
                self.imported_excel_path = self.excel_path.get()
                self.refresh_model_statuses()
                self.load_selected_imported_table()
                messagebox.showinfo("Importacion", "Excel importado correctamente a SQLite local.")
            else:
                self.is_excel_imported = False
                self.imported_excel_path = None
                self.reset_model_selector()
                messagebox.showerror("Importacion", "\n".join(issue.message for issue in result.errors[:8]))
            self.status_text.set(f"Importacion batch {result.batch_id}: {result.imported_rows}")
        except Exception as exc:
            self.is_excel_imported = False
            self.imported_excel_path = None
            self.reset_model_selector()
            messagebox.showerror("Importacion", str(exc))
            self.status_text.set(f"Error de importacion: {exc}")

    def refresh_model_statuses(self) -> None:
        if not self.is_excel_imported:
            self.reset_model_selector()
            return
        self.model_statuses = evaluate_models(self.db_path)
        values = [
            f"{item.code} | {item.label}: {item.status} - {item.visual_name}"
            for item in self.model_statuses
        ]
        self.model_combo.configure(values=values, state="readonly")
        self.execute_button.configure(state=tk.NORMAL)
        if values:
            self.model_combo.current(0)
            self.model_var.set(values[0])
        self.show_selected_model_reason()
        self.set_criteria_text(
            "\n\n".join(
                f"{item.label}: {item.status}\n"
                f"Nombre visual: {item.visual_name}\n"
                f"Flujo: {item.flow}\n"
                f"Motivo: {item.reason}"
                for item in self.model_statuses
            )
        )

    def selected_model(self) -> CostingModelStatus:
        if not self.model_statuses:
            raise RuntimeError("Primero valide e importe un Excel maestro para evaluar modelos.")
        selected = self.model_var.get().split("|", 1)[0].strip()
        for item in self.model_statuses:
            if item.code == selected:
                return item
        return self.model_statuses[0]

    def show_selected_model_reason(self) -> None:
        if not hasattr(self, "model_reason"):
            return
        item = self.selected_model() if self.model_statuses else None
        self.model_reason.configure(state=tk.NORMAL)
        self.model_reason.delete("1.0", tk.END)
        if item:
            self.model_reason.insert(
                "1.0",
                f"{item.label}: {item.status}\nNombre visual: {item.visual_name}\nFlujo: {item.flow}\n\n{item.reason}",
            )
        else:
            self.model_reason.insert(
                "1.0",
                "Los modelos se mostraran despues de validar e importar el Excel maestro.",
            )
        self.model_reason.configure(state=tk.DISABLED)

    def execute_calculation(self) -> None:
        if not self.is_excel_validated or self.validated_excel_path != self.excel_path.get():
            messagebox.showwarning("Calculo ABC", "Primero valide la estructura del Excel maestro seleccionado.")
            self.status_text.set("No se puede ejecutar: falta validacion del Excel actual.")
            return
        if not self.is_excel_imported or self.imported_excel_path != self.excel_path.get():
            messagebox.showwarning("Calculo ABC", "Primero importe el Excel maestro validado a SQLite.")
            self.status_text.set("No se puede ejecutar: falta importacion del Excel actual.")
            return
        item = self.selected_model()
        if not item.available:
            messagebox.showwarning(item.label, item.reason)
            self.status_text.set(item.reason)
            return
        try:
            run_id = run_model_1_calculation(self.db_path)
            self.refresh_all_views()
            self.show_modules_screen()
            messagebox.showinfo("Calculo ABC", f"Calculo completado. Ejecucion: {run_id}")
            self.status_text.set(f"Calculo Modelo ABC por Ordenes de Trabajo completado. Run {run_id}.")
        except Exception as exc:
            messagebox.showerror("Calculo ABC", str(exc))
            self.status_text.set(f"Error de calculo: {exc}")

    def load_selected_imported_table(self) -> None:
        table = self.table_var.get()
        try:
            with db_session(self.db_path) as conn:
                rows = [dict(row) for row in conn.execute(f"SELECT * FROM {quote_identifier(table)} LIMIT 500").fetchall()]
            self.populate_tree("datos", rows)
        except Exception as exc:
            self.status_text.set(f"No se pudo cargar {table}: {exc}")

    def refresh_all_views(self) -> None:
        self.load_selected_imported_table()
        with db_session(self.db_path) as conn:
            run_id = latest_run_id(conn)
            queries = {
                "proceso1": """
                    SELECT run_id, actividad_id, codigo_actividad, nombre_actividad,
                           costo_mod_ac AS costo_pd_ac, costo_otros_recursos_ac,
                           costo_conversion_ac, costo_soporte_ac, costo_total_ac,
                           formula, fuente, observacion
                    FROM actividad_total
                    WHERE run_id = ?
                """,
                "proceso2_produccion": """
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
                "proceso2_soporte": """
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
                "hoja": "SELECT * FROM hoja_costos_ot WHERE run_id = ?",
                "estado": "SELECT * FROM estado_resultados_ot WHERE run_id = ?",
            }
            for key, sql in queries.items():
                rows = []
                if run_id:
                    rows = [dict(row) for row in conn.execute(sql, (run_id,)).fetchall()]
                self.populate_tree(key, rows)
        self.refresh_dashboard_text()
        self.refresh_outputs()

    def reset_model_selector(self) -> None:
        self.model_statuses = []
        self.model_var.set("")
        self.model_combo.configure(values=[], state=tk.DISABLED)
        if hasattr(self, "execute_button"):
            self.execute_button.configure(state=tk.DISABLED)
        self.show_selected_model_reason()
        self.set_criteria_text(
            "Seleccione un Excel maestro, valide su estructura e importelo a SQLite.\n\n"
            "La disponibilidad de Modelo 1, Modelo 2 y Modelo 3 se evaluara recien despues de la importacion."
        )

    def set_criteria_text(self, text: str) -> None:
        if not hasattr(self, "criteria_text"):
            return
        self.criteria_text.configure(state=tk.NORMAL)
        self.criteria_text.delete("1.0", tk.END)
        self.criteria_text.insert("1.0", text)
        self.criteria_text.configure(state=tk.DISABLED)

    def hide_module_tabs(self) -> None:
        for tab, _title in self.module_tabs:
            self.notebook.hide(tab)

    def show_modules_screen(self) -> None:
        for tab, title in self.module_tabs:
            self.notebook.add(tab, text=title)
        self.notebook.hide(self.home_tab)
        if self.module_tabs:
            self.notebook.select(self.module_tabs[0][0])

    def show_home_screen(self) -> None:
        self.notebook.add(self.home_tab, text="Pantalla inicial")
        self.hide_module_tabs()
        self.notebook.select(self.home_tab)

    def populate_tree(self, key: str, rows: list[dict]) -> None:
        tree = self.trees.get(key)
        if tree is None:
            return
        tree.delete(*tree.get_children())
        if not rows:
            tree.configure(columns=["mensaje"])
            tree.heading("mensaje", text="Mensaje")
            tree.column("mensaje", width=420)
            tree.insert("", tk.END, values=["Sin datos para mostrar."])
            return
        columns = list(rows[0].keys())
        tree.configure(columns=columns)
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=150, anchor=tk.W)
        for row in rows:
            tree.insert("", tk.END, values=["" if row.get(col) is None else row.get(col) for col in columns])

    def refresh_dashboard_text(self) -> None:
        payload = dashboard_payload(self.db_path)
        self.dashboard_text.configure(state=tk.NORMAL)
        self.dashboard_text.delete("1.0", tk.END)
        if not payload:
            self.dashboard_text.insert("1.0", "Ejecute el calculo del Modelo ABC por Ordenes de Trabajo para ver el dashboard.")
            self.dashboard_text.configure(state=tk.DISABLED)
            return
        lines = ["DASHBOARD DE COSTOS\n"]
        for row in payload.get("costos_ot", []):
            lines.append(
                f"OT {row.get('codigo_ot')}: costo total S/ {row.get('costo_total_ot', 0):,.2f}; "
                f"costo unitario S/ {row.get('costo_unitario_ot', 0):,.4f}"
            )
        lines.append("\nRanking de actividades:")
        for row in sorted(payload.get("costos_actividad", []), key=lambda x: x.get("costo_total_ac") or 0, reverse=True):
            lines.append(f"- {row.get('nombre_actividad')}: S/ {row.get('costo_total_ac', 0):,.2f}")
        lines.append("\nMargenes por OT:")
        for row in payload.get("estado_resultados", []):
            lines.append(
                f"- {row.get('codigo_producto')}: margen bruto S/ {row.get('margen_bruto', 0):,.2f}; "
                f"margen operativo S/ {row.get('margen_operativo', 0):,.2f}"
            )
        lines.append("\nAlertas:")
        for row in payload.get("alertas", []):
            lines.append(f"- [{row.get('severidad')}] {row.get('mensaje')}")
        self.dashboard_text.insert("1.0", "\n".join(lines))
        self.dashboard_text.configure(state=tk.DISABLED)

    def refresh_outputs(self) -> None:
        rows = []
        for path in sorted(REPORTS_DIR.glob("*")) if REPORTS_DIR.exists() else []:
            rows.append({"archivo": path.name, "ruta": str(path), "tipo": path.suffix})
        self.populate_tree("salidas", rows)

    def generate_reports(self) -> None:
        try:
            paths = generate_all_reports(self.db_path, REPORTS_DIR)
            self.refresh_outputs()
            messagebox.showinfo("Reportes", "Reportes generados:\n" + "\n".join(str(path) for path in paths))
        except Exception as exc:
            messagebox.showerror("Reportes", str(exc))

    def generate_documents(self) -> None:
        try:
            paths = generate_all_documents(self.db_path, REPORTS_DIR)
            self.refresh_outputs()
            messagebox.showinfo("Documentos", "Documentos generados:\n" + "\n".join(str(path) for path in paths))
        except Exception as exc:
            messagebox.showerror("Documentos", str(exc))

    def generate_dashboard_png(self) -> None:
        path = generate_dashboard_image(self.db_path, REPORTS_DIR)
        self.refresh_outputs()
        if path:
            messagebox.showinfo("Dashboard", f"Dashboard PNG generado: {path}")
        else:
            messagebox.showwarning(
                "Dashboard",
                "No se genero PNG porque matplotlib no esta disponible o no hay datos. El dashboard Excel e interno siguen disponibles.",
            )

    def run_tests(self) -> None:
        try:
            from ..tests_runner import run_functional_tests

            result = run_functional_tests(self.excel_path.get(), self.db_path, REPORTS_DIR, REPORTS_DIR)
            self.refresh_all_views()
            summary = f"Pruebas ejecutadas: {len(result)}\nConformes: {sum(1 for item in result if item['ok'])}"
            messagebox.showinfo("Pruebas", summary)
        except Exception as exc:
            messagebox.showerror("Pruebas", str(exc))


def run_app() -> None:
    app = CostosABCApp()
    app.mainloop()
