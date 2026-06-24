from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

from ..constants import (
    APP_TITLE,
    DEFAULT_DB_PATH,
    IMAGES_MODELS_DIR,
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
    "periodo",
    "modelos_costeo",
    "objetos_costo",
    "ordenes_trabajo",
    "materiales_directos",
    "Personal_directos",
    "servicios_directos",
    "Costo_Indirectos",
    "actividades",
    "inductores",
    "ec_ac_inductores",
    "soporte_cc_ing_inductores",
    "ac_ot_pesos_produccion",
    "ac_ot_pesos_soporte",
    "ventas",
    "reglas_calculo",
    "fuentes",
    "diccionario",
]

INDIRECT_COST_TABLES = [
    "costo_otros_recursos(CI)",
    "Soporte_Mantenimiento(CI)",
    "Soporte_Control_Calidad(CI)",
    "Soporte_Ingenieria(CI)",
]

DETAIL_MODULES = {"proceso1", "proceso2_produccion", "proceso2_soporte", "hoja", "estado"}


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
        self.indirect_cost_table_var = tk.StringVar(value=INDIRECT_COST_TABLES[0])
        self.model_statuses: list[CostingModelStatus] = []
        self.trees: dict[str, ttk.Treeview] = {}
        self.tree_row_data: dict[str, dict[str, dict]] = {}
        self.module_tabs: list[tuple[ttk.Frame, str]] = []
        self.dashboard_chart_canvas = None
        self.dashboard_kpis: dict[str, tk.Label] = {}
        self.dashboard_header_status: tk.Label | None = None
        self.dashboard_period_label: tk.Label | None = None
        self.dashboard_message: ttk.Label | None = None
        self.dashboard_chart_frame = None
        self.dashboard_insights_text: tk.Text | None = None
        self.criteria_image: tk.PhotoImage | None = None
        self.criteria_source_image: tk.PhotoImage | None = None
        self.criteria_image_model_code: str | None = None
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

        note = ttk.LabelFrame(tab, text="Grafico del modelo de costeo", padding=10, style="Section.TLabelframe")
        note.pack(fill=tk.BOTH, expand=True, pady=12)
        self.criteria_canvas = tk.Canvas(note, bg="#FFFFFF", highlightthickness=1, highlightbackground="#CBD5E1")
        self.criteria_canvas.pack(fill=tk.BOTH, expand=True)
        self.criteria_canvas.bind("<Configure>", lambda _event: self.render_model_diagram())
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
        combo.bind("<<ComboboxSelected>>", lambda _event: self.on_imported_table_selected())
        self.indirect_cost_frame = ttk.Frame(top)
        ttk.Label(self.indirect_cost_frame, text="Tipo CI:").pack(side=tk.LEFT)
        indirect_combo = ttk.Combobox(
            self.indirect_cost_frame,
            textvariable=self.indirect_cost_table_var,
            values=INDIRECT_COST_TABLES,
            state="readonly",
            width=34,
        )
        indirect_combo.pack(side=tk.LEFT, padx=(6, 0))
        indirect_combo.bind("<<ComboboxSelected>>", lambda _event: self.load_selected_imported_table())
        self.refresh_table_button = ttk.Button(top, text="Actualizar", command=self.load_selected_imported_table)
        self.refresh_table_button.pack(side=tk.LEFT)
        ttk.Button(top, text="Ver Dashboard", command=self.show_dashboard_tab).pack(side=tk.LEFT, padx=(8, 0))
        self.update_indirect_cost_selector()
        tree = self._make_tree(tab)
        self.trees["datos"] = tree

    def _build_table_tab(self, title: str, key: str) -> None:
        tab = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(tab, text=title)
        self.module_tabs.append((tab, title))
        ttk.Label(tab, text=title, style="Subtitle.TLabel").pack(anchor=tk.W)
        self.trees[key] = self._make_tree(tab)
        if key in DETAIL_MODULES:
            self.trees[key].bind("<ButtonRelease-1>", lambda event, module_key=key: self.open_calculation_detail_from_event(module_key, event))

    def _build_dashboard_tab(self) -> None:
        tab = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(tab, text="Dashboard")
        self.module_tabs.append((tab, "Dashboard"))
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(3, weight=1)

        header = tk.Frame(tab, bg="#0F172A", padx=16, pady=12)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        tk.Label(
            header,
            text="Dashboard gerencial de costos",
            bg="#0F172A",
            fg="#F8FAFC",
            font=("Segoe UI", 15, "bold"),
        ).grid(row=0, column=0, sticky="w")
        self.dashboard_period_label = tk.Label(
            header,
            text="Frenosa | Modelo ABC por Ordenes de Trabajo",
            bg="#0F172A",
            fg="#CBD5E1",
            font=("Segoe UI", 9),
        )
        self.dashboard_period_label.grid(row=1, column=0, sticky="w", pady=(2, 0))
        self.dashboard_header_status = tk.Label(
            header,
            text="Pendiente de calculo",
            bg="#1E293B",
            fg="#E2E8F0",
            font=("Segoe UI", 9, "bold"),
            padx=12,
            pady=5,
        )
        self.dashboard_header_status.grid(row=0, column=1, rowspan=2, sticky="e")

        kpi_frame = tk.Frame(tab, bg="#F8FAFC")
        kpi_frame.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        for index, (key, title, accent, caption) in enumerate(
            [
                ("total_cost", "Costo total OT", "#2563EB", "Costo acumulado"),
                ("avg_unit", "Costo unitario prom.", "#0F766E", "Promedio por unidad"),
                ("sales", "Ventas netas", "#D97706", "Ingreso total"),
                ("margin", "Rentabilidad prom.", "#7C3AED", "Margen operativo"),
            ]
        ):
            self.create_dashboard_kpi(kpi_frame, index, key, title, accent, caption)
            kpi_frame.columnconfigure(index, weight=1)

        self.dashboard_message = ttk.Label(tab, text="", style="Subtitle.TLabel")
        self.dashboard_message.grid(row=2, column=0, sticky="w", pady=(8, 4))
        self.dashboard_chart_frame = tk.Frame(tab, bg="#FFFFFF", highlightthickness=1, highlightbackground="#CBD5E1")
        self.dashboard_chart_frame.grid(row=3, column=0, sticky="nsew")
        insights = ttk.LabelFrame(tab, text="Alertas y recomendaciones", padding=8, style="Section.TLabelframe")
        insights.grid(row=4, column=0, sticky="ew", pady=(8, 0))
        self.dashboard_insights_text = tk.Text(
            insights,
            wrap=tk.WORD,
            height=5,
            bg="#FFFFFF",
            fg="#0F172A",
            font=("Segoe UI", 9),
            relief=tk.FLAT,
        )
        insights_scroll = ttk.Scrollbar(insights, orient=tk.VERTICAL, command=self.dashboard_insights_text.yview)
        self.dashboard_insights_text.configure(yscrollcommand=insights_scroll.set)
        self.dashboard_insights_text.grid(row=0, column=0, sticky="ew")
        insights_scroll.grid(row=0, column=1, sticky="ns")
        insights.columnconfigure(0, weight=1)

    def create_dashboard_kpi(self, parent, index: int, key: str, title: str, accent: str, caption: str) -> None:
        card = tk.Frame(parent, bg="#FFFFFF", highlightthickness=1, highlightbackground="#CBD5E1")
        card.grid(row=0, column=index, sticky="ew", padx=(0, 8))
        accent_bar = tk.Frame(card, bg=accent, width=5)
        accent_bar.pack(side=tk.LEFT, fill=tk.Y)
        content = tk.Frame(card, bg="#FFFFFF", padx=10, pady=8)
        content.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tk.Label(
            content,
            text=title,
            bg="#FFFFFF",
            fg="#475569",
            font=("Segoe UI", 8, "bold"),
        ).pack(anchor=tk.W)
        value = tk.Label(
            content,
            text="S/ 0.00",
            bg="#FFFFFF",
            fg="#0F172A",
            font=("Segoe UI", 14, "bold"),
        )
        value.pack(anchor=tk.W, pady=(2, 0))
        tk.Label(
            content,
            text=caption,
            bg="#FFFFFF",
            fg="#64748B",
            font=("Segoe UI", 8),
        ).pack(anchor=tk.W, pady=(2, 0))
        self.dashboard_kpis[key] = value

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
        self.model_statuses = [
            item for item in evaluate_models(self.db_path) if item.code != "MODELO_3"
        ]
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
        self.update_model_diagram(item.code if item else None)

    def update_model_diagram(self, model_code: str | None = None) -> None:
        if not hasattr(self, "criteria_canvas"):
            return
        image_name = {
            "MODELO_1": "modelo1.png",
            "MODELO_2": "modelo2.png",
        }.get(model_code or "")
        if not image_name:
            self.criteria_image_model_code = None
            self.criteria_source_image = None
            self.criteria_image = None
            self.criteria_canvas.delete("all")
            return
        path = IMAGES_MODELS_DIR / image_name
        if not path.exists():
            self.criteria_image_model_code = None
            self.criteria_source_image = None
            self.criteria_image = None
            self.criteria_canvas.delete("all")
            return
        self.criteria_image_model_code = model_code
        self.criteria_source_image = tk.PhotoImage(file=str(path))
        self.render_model_diagram()

    def render_model_diagram(self) -> None:
        if not hasattr(self, "criteria_canvas"):
            return
        self.criteria_canvas.delete("all")
        if self.criteria_source_image is None:
            return
        canvas_width = max(self.criteria_canvas.winfo_width(), 1)
        canvas_height = max(self.criteria_canvas.winfo_height(), 1)
        image_width = self.criteria_source_image.width()
        image_height = self.criteria_source_image.height()
        factor = max(
            1,
            (image_width + max(canvas_width - 24, 1) - 1) // max(canvas_width - 24, 1),
            (image_height + max(canvas_height - 24, 1) - 1) // max(canvas_height - 24, 1),
        )
        self.criteria_image = self.criteria_source_image.subsample(factor, factor)
        x = max((canvas_width - self.criteria_image.width()) // 2, 12)
        y = max((canvas_height - self.criteria_image.height()) // 2, 12)
        self.criteria_canvas.create_image(x, y, anchor=tk.NW, image=self.criteria_image)
        self.criteria_canvas.configure(scrollregion=(0, 0, canvas_width, canvas_height))

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

    def on_imported_table_selected(self) -> None:
        self.update_indirect_cost_selector()
        self.load_selected_imported_table()

    def update_indirect_cost_selector(self) -> None:
        if not hasattr(self, "indirect_cost_frame"):
            return
        if self.table_var.get() == "Costo_Indirectos":
            if not self.indirect_cost_frame.winfo_ismapped():
                pack_options = {"side": tk.LEFT, "padx": (0, 8)}
                if hasattr(self, "refresh_table_button"):
                    pack_options["before"] = self.refresh_table_button
                self.indirect_cost_frame.pack(**pack_options)
        else:
            self.indirect_cost_frame.pack_forget()

    def selected_imported_table_name(self) -> str:
        if self.table_var.get() == "Costo_Indirectos":
            return self.indirect_cost_table_var.get()
        return self.table_var.get()

    def load_selected_imported_table(self) -> None:
        table = self.selected_imported_table_name()
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
        self.update_model_diagram()

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

    def show_dashboard_tab(self) -> None:
        for tab, title in self.module_tabs:
            if title == "Dashboard":
                self.notebook.select(tab)
                return

    def populate_tree(self, key: str, rows: list[dict]) -> None:
        tree = self.trees.get(key)
        if tree is None:
            return
        self.tree_row_data[key] = {}
        tree.delete(*tree.get_children())
        if not rows:
            tree.configure(columns=["mensaje"])
            tree.heading("mensaje", text="Mensaje")
            tree.column("mensaje", width=420)
            tree.insert("", tk.END, values=["Sin datos para mostrar."])
            return
        columns = list(rows[0].keys())
        if key in DETAIL_MODULES:
            columns = ["ver_calculos"] + columns
        tree.configure(columns=columns)
        for col in columns:
            tree.heading(col, text="Ver calculos" if col == "ver_calculos" else col)
            width = 110 if col == "ver_calculos" else 150
            tree.column(col, width=width, anchor=tk.CENTER if col == "ver_calculos" else tk.W)
        for row in rows:
            values = []
            for col in columns:
                if col == "ver_calculos":
                    values.append("Ver calculos")
                else:
                    values.append("" if row.get(col) is None else row.get(col))
            item_id = tree.insert("", tk.END, values=values)
            self.tree_row_data[key][item_id] = dict(row)

    def open_calculation_detail_from_event(self, key: str, event) -> None:
        tree = self.trees.get(key)
        if tree is None:
            return
        if tree.identify("region", event.x, event.y) != "cell":
            return
        if tree.identify_column(event.x) != "#1":
            return
        item_id = tree.identify_row(event.y)
        row = self.tree_row_data.get(key, {}).get(item_id)
        if row:
            self.show_calculation_detail(key, row)

    def show_calculation_detail(self, key: str, row: dict) -> None:
        title = {
            "proceso1": "Detalle de calculos - Proceso 1 EC-AC",
            "proceso2_produccion": "Detalle de calculos - Proceso 2 AC-OT (Produccion)",
            "proceso2_soporte": "Detalle de calculos - Proceso 2 AC-OT (Soporte)",
            "hoja": "Detalle de calculos - Hoja de costos OT",
            "estado": "Detalle de calculos - Estado de resultados",
        }.get(key, "Detalle de calculos")
        window = tk.Toplevel(self)
        window.title(title)
        window.geometry("980x680")
        window.minsize(820, 520)

        frame = ttk.Frame(window, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        text = tk.Text(frame, wrap=tk.WORD)
        yscroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=text.yview)
        text.configure(yscrollcommand=yscroll.set)
        text.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        ttk.Button(frame, text="Cerrar", command=window.destroy).grid(row=1, column=0, sticky=tk.E, pady=(8, 0))

        try:
            lines = self.build_calculation_detail_lines(key, row)
        except Exception as exc:
            lines = [f"No se pudo construir el detalle de calculos: {exc}"]
        text.insert("1.0", "\n".join(lines))
        text.configure(state=tk.DISABLED)

    def build_calculation_detail_lines(self, key: str, row: dict) -> list[str]:
        if key == "proceso1":
            return self.build_proceso1_detail(row)
        if key == "proceso2_produccion":
            return self.build_proceso2_produccion_detail(row)
        if key == "proceso2_soporte":
            return self.build_proceso2_soporte_detail(row)
        if key == "hoja":
            return self.build_hoja_costos_detail(row)
        if key == "estado":
            return self.build_estado_resultados_detail(row)
        return ["Modulo sin detalle configurado."]

    def build_proceso1_detail(self, row: dict) -> list[str]:
        run_id = row.get("run_id")
        actividad_id = row.get("actividad_id")
        lines = [f"Actividad: {row.get('codigo_actividad')} - {row.get('nombre_actividad')}", f"Ejecucion: {run_id}", ""]
        with db_session(self.db_path) as conn:
            actividad_total = self.fetch_one_dict(
                conn,
                "SELECT * FROM actividad_total WHERE run_id = ? AND actividad_id = ?",
                (run_id, actividad_id),
            )
            mod_rows = self.fetch_all_dicts(
                conn,
                "SELECT * FROM mod_actividad WHERE run_id = ? AND actividad_id = ?",
                (run_id, actividad_id),
            )
            pd_rows = self.fetch_all_dicts(
                conn,
                'SELECT * FROM "Personal_directos" WHERE actividad_id = ?',
                (actividad_id,),
            )
            otros_detalle = self.fetch_all_dicts(
                conn,
                "SELECT * FROM otros_recursos_ac_detalle WHERE run_id = ? AND actividad_id = ? ORDER BY elemento_costo_id",
                (run_id, actividad_id),
            )
            otros_total = self.fetch_one_dict(
                conn,
                "SELECT * FROM otros_recursos_actividad WHERE run_id = ? AND actividad_id = ?",
                (run_id, actividad_id),
            )
            soporte_total = self.fetch_one_dict(
                conn,
                "SELECT * FROM soporte_actividad WHERE run_id = ? AND actividad_id = ?",
                (run_id, actividad_id),
            )
            mantenimiento_calc = self.fetch_one_dict(
                conn,
                "SELECT * FROM soporte_mantenimiento_ac WHERE run_id = ? AND actividad_id = ?",
                (run_id, actividad_id),
            )
            mantenimiento_input = self.fetch_all_dicts(
                conn,
                'SELECT * FROM "Soporte_Mantenimiento(CI)" WHERE actividad_id_destino = ? ORDER BY soporte_item_id',
                (actividad_id,),
            )
            cc_calc = self.fetch_one_dict(
                conn,
                "SELECT * FROM soporte_control_calidad_ac WHERE run_id = ? AND actividad_id = ?",
                (run_id, actividad_id),
            )
            cc_inductores = self.fetch_all_dicts(
                conn,
                "SELECT * FROM soporte_cc_ing_inductores WHERE actividad_id = ? AND lower(proceso_soporte) LIKE '%calidad%'",
                (actividad_id,),
            )
            cc_inputs = self.fetch_all_dicts(
                conn,
                'SELECT * FROM "Soporte_Control_Calidad(CI)" ORDER BY soporte_item_id',
            )
            ing_calc = self.fetch_one_dict(
                conn,
                "SELECT * FROM soporte_ingenieria_ac WHERE run_id = ? AND actividad_id = ?",
                (run_id, actividad_id),
            )
            ing_inductores = self.fetch_all_dicts(
                conn,
                "SELECT * FROM soporte_cc_ing_inductores WHERE actividad_id = ? AND lower(proceso_soporte) LIKE '%ingen%'",
                (actividad_id,),
            )
            ing_inputs = self.fetch_all_dicts(
                conn,
                'SELECT * FROM "Soporte_Ingenieria(CI)" ORDER BY soporte_item_id',
            )

        self.add_section(lines, "Resultado de la fila", actividad_total)
        self.add_section(lines, "Datos de entrada: Personal_directos", pd_rows)
        self.add_section(lines, "Calculo PD/MOD hacia actividad", mod_rows)
        self.add_section(lines, "Datos y calculos: otros recursos EC-AC", otros_detalle)
        self.add_section(lines, "Total otros recursos por actividad", otros_total)
        self.add_section(lines, "Soporte total por actividad", soporte_total)
        self.add_section(lines, "Mantenimiento: datos fuente", mantenimiento_input)
        self.add_section(lines, "Mantenimiento: calculo agrupado", mantenimiento_calc)
        self.add_section(lines, "Control de calidad: datos fuente usados para el total", cc_inputs)
        self.add_section(lines, "Control de calidad: inductores usados", cc_inductores)
        self.add_section(lines, "Control de calidad: calculo hacia actividad", cc_calc)
        self.add_section(lines, "Ingenieria: datos fuente usados para el total", ing_inputs)
        self.add_section(lines, "Ingenieria: inductores usados", ing_inductores)
        self.add_section(lines, "Ingenieria: calculo hacia actividad", ing_calc)

        mod = self.as_number(actividad_total, "costo_mod_ac")
        otros = self.as_number(actividad_total, "costo_otros_recursos_ac")
        conversion = self.as_number(actividad_total, "costo_conversion_ac")
        soporte = self.as_number(actividad_total, "costo_soporte_ac")
        total = self.as_number(actividad_total, "costo_total_ac")
        mant = self.as_number(soporte_total, "costo_mantenimiento_ac")
        cc = self.as_number(soporte_total, "costo_control_calidad_ac")
        ing = self.as_number(soporte_total, "costo_ingenieria_ac")
        lines.extend(
            [
                "ECUACIONES NUMERICAS",
                f"costo_otros_recursos_ac = SUM(costo_recurso_ac) = {self.format_number(sum(self.as_number(item, 'costo_recurso_ac') for item in otros_detalle))}",
                f"costo_conversion_ac = costo_pd_ac + costo_otros_recursos_ac = {self.format_number(mod)} + {self.format_number(otros)} = {self.format_number(conversion)}",
                f"costo_soporte_ac = mantenimiento_ac + control_calidad_ac + ingenieria_ac = {self.format_number(mant)} + {self.format_number(cc)} + {self.format_number(ing)} = {self.format_number(soporte)}",
                f"costo_total_ac = costo_conversion_ac + costo_soporte_ac = {self.format_number(conversion)} + {self.format_number(soporte)} = {self.format_number(total)}",
                "",
            ]
        )
        return lines

    def build_proceso2_produccion_detail(self, row: dict) -> list[str]:
        run_id = row.get("run_id")
        actividad_id = row.get("actividad_id")
        ot_id = row.get("ot_id")
        lines = [
            f"OT: {ot_id} | Actividad: {row.get('codigo_actividad')} - {row.get('nombre_actividad')}",
            f"Ejecucion: {run_id}",
            "",
        ]
        with db_session(self.db_path) as conn:
            detalle = self.fetch_one_dict(
                conn,
                "SELECT * FROM produccion_ot_detalle WHERE run_id = ? AND actividad_id = ? AND ot_id = ?",
                (run_id, actividad_id, ot_id),
            )
            peso_input = self.fetch_one_dict(
                conn,
                "SELECT * FROM ac_ot_pesos_produccion WHERE actividad_id = ? AND ot_id = ?",
                (actividad_id, ot_id),
            )
            peso_inputs_ot = self.fetch_all_dicts(
                conn,
                "SELECT * FROM ac_ot_pesos_produccion WHERE ot_id = ? ORDER BY actividad_id",
                (ot_id,),
            )
            actividad_total = self.fetch_one_dict(
                conn,
                "SELECT * FROM actividad_total WHERE run_id = ? AND actividad_id = ?",
                (run_id, actividad_id),
            )
            total_ot = self.fetch_one_dict(
                conn,
                "SELECT * FROM produccion_ot WHERE run_id = ? AND ot_id = ?",
                (run_id, ot_id),
            )
            sumatoria = self.fetch_all_dicts(
                conn,
                "SELECT actividad_id, codigo_actividad, nombre_actividad, peso_ac_ot, costo_conversion_ac, costo_conversion_ot FROM produccion_ot_detalle WHERE run_id = ? AND ot_id = ? ORDER BY actividad_id",
                (run_id, ot_id),
            )
        self.add_section(lines, "Resultado de la fila", detalle)
        self.add_section(lines, "Datos de entrada: AC_OT_Pesos_Produccion", peso_input)
        self.add_section(lines, "Datos de entrada: todos los pesos de produccion de la OT", peso_inputs_ot)
        self.add_section(lines, "Costo de actividad de origen", actividad_total)
        self.add_section(lines, "Sumatoria para costo_produccion_ot", sumatoria)
        self.add_section(lines, "Resultado final por OT", total_ot)

        costo_ac = self.as_number(detalle, "costo_conversion_ac")
        peso = self.as_number(detalle, "peso_ac_ot")
        costo_detalle = self.as_number(detalle, "costo_conversion_ot")
        total = self.as_number(total_ot, "costo_produccion_ot")
        lines.extend(
            [
                "ECUACIONES NUMERICAS",
                f"costo_conversion_ot = costo_conversion_ac * peso_ac_ot = {self.format_number(costo_ac)} * {self.format_number(peso)} = {self.format_number(costo_detalle)}",
                f"costo_produccion_ot = SUM(costo_conversion_ot de la OT) = {self.format_number(sum(self.as_number(item, 'costo_conversion_ot') for item in sumatoria))} = {self.format_number(total)}",
                "",
            ]
        )
        return lines

    def build_proceso2_soporte_detail(self, row: dict) -> list[str]:
        run_id = row.get("run_id")
        actividad_id = row.get("actividad_id")
        ot_id = row.get("ot_id")
        lines = [
            f"OT: {ot_id} | Actividad: {row.get('codigo_actividad')} - {row.get('nombre_actividad')}",
            f"Ejecucion: {run_id}",
            "",
        ]
        with db_session(self.db_path) as conn:
            detalle = self.fetch_one_dict(
                conn,
                "SELECT * FROM soporte_ot_detalle WHERE run_id = ? AND actividad_id = ? AND ot_id = ?",
                (run_id, actividad_id, ot_id),
            )
            peso_input = self.fetch_one_dict(
                conn,
                "SELECT * FROM ac_ot_pesos_soporte WHERE actividad_id = ? AND ot_id = ?",
                (actividad_id, ot_id),
            )
            peso_inputs_ot = self.fetch_all_dicts(
                conn,
                "SELECT * FROM ac_ot_pesos_soporte WHERE ot_id = ? ORDER BY actividad_id",
                (ot_id,),
            )
            soporte_actividad = self.fetch_one_dict(
                conn,
                "SELECT * FROM soporte_actividad WHERE run_id = ? AND actividad_id = ?",
                (run_id, actividad_id),
            )
            total_ot = self.fetch_one_dict(
                conn,
                "SELECT * FROM soporte_ot WHERE run_id = ? AND ot_id = ?",
                (run_id, ot_id),
            )
            sumatoria = self.fetch_all_dicts(
                conn,
                "SELECT actividad_id, codigo_actividad, nombre_actividad, peso_ac_ot, costo_soporte_ac, costo_soporte_ot FROM soporte_ot_detalle WHERE run_id = ? AND ot_id = ? ORDER BY actividad_id",
                (run_id, ot_id),
            )
        self.add_section(lines, "Resultado de la fila", detalle)
        self.add_section(lines, "Datos de entrada: AC_OT_Pesos_Soporte", peso_input)
        self.add_section(lines, "Datos de entrada: todos los pesos de soporte de la OT", peso_inputs_ot)
        self.add_section(lines, "Costo de soporte de actividad de origen", soporte_actividad)
        self.add_section(lines, "Sumatoria para costo_soporte_ot", sumatoria)
        self.add_section(lines, "Resultado final por OT", total_ot)

        costo_ac = self.as_number(detalle, "costo_soporte_ac")
        peso = self.as_number(detalle, "peso_ac_ot")
        costo_detalle = self.as_number(detalle, "costo_soporte_ot")
        total = self.as_number(total_ot, "costo_soporte_ot")
        lines.extend(
            [
                "ECUACIONES NUMERICAS",
                f"costo_soporte_ot_detalle = costo_soporte_ac * peso_ac_ot = {self.format_number(costo_ac)} * {self.format_number(peso)} = {self.format_number(costo_detalle)}",
                f"costo_soporte_ot = SUM(costo_soporte_ot_detalle de la OT) = {self.format_number(sum(self.as_number(item, 'costo_soporte_ot') for item in sumatoria))} = {self.format_number(total)}",
                "",
            ]
        )
        return lines

    def fetch_one_dict(self, conn, sql: str, params: tuple = ()) -> dict:
        row = conn.execute(sql, params).fetchone()
        return dict(row) if row else {}

    def build_hoja_costos_detail(self, row: dict) -> list[str]:
        lines = [
            f"OT: {row.get('codigo_ot') or row.get('ot_id')} | Producto: {row.get('codigo_producto')}",
            f"Ejecucion: {row.get('run_id')}",
            "",
        ]
        self.add_section(lines, "Datos de entrada y resultados de la fila Hoja de costos OT", row)

        cmd = self.as_number(row, "cmd_ot")
        produccion = self.as_number(row, "costo_produccion_ot")
        soporte = self.as_number(row, "costo_soporte_ot")
        csd = self.as_number(row, "csd_ot")
        total = self.as_number(row, "costo_total_ot")
        cantidad = self.as_number(row, "cantidad_producida")
        unitario = self.as_number(row, "costo_unitario_ot")
        calculated_total = cmd + produccion + soporte + csd
        calculated_unit = calculated_total / cantidad if cantidad else 0.0
        lines.extend(
            [
                "ECUACIONES NUMERICAS",
                f"costo_total_ot = cmd_ot + costo_produccion_ot + costo_soporte_ot + csd_ot = {self.format_number(cmd)} + {self.format_number(produccion)} + {self.format_number(soporte)} + {self.format_number(csd)} = {self.format_number(calculated_total)}",
                f"Resultado guardado costo_total_ot = {self.format_number(total)}",
                f"costo_unitario_ot = costo_total_ot / cantidad_producida = {self.format_number(total)} / {self.format_number(cantidad)} = {self.format_number(calculated_unit)}",
                f"Resultado guardado costo_unitario_ot = {self.format_number(unitario)}",
                "",
            ]
        )
        return lines

    def build_estado_resultados_detail(self, row: dict) -> list[str]:
        lines = [
            f"OT: {row.get('ot_id')} | Producto: {row.get('codigo_producto')}",
            f"Ejecucion: {row.get('run_id')}",
            "",
        ]
        self.add_section(lines, "Datos de entrada y resultados de la fila Estado de resultados", row)

        ventas = self.as_number(row, "ventas_netas_soles")
        materia = self.as_number(row, "materia_prima_directa")
        conversion = self.as_number(row, "costos_conversion")
        soporte = self.as_number(row, "procesos_soporte")
        margen_bruto = self.as_number(row, "margen_bruto")
        margen_operativo = self.as_number(row, "margen_operativo")
        rentabilidad = self.as_number(row, "rentabilidad_operativa")
        calculated_margen_bruto = ventas - materia - conversion
        calculated_margen_operativo = calculated_margen_bruto - soporte
        calculated_rentabilidad = calculated_margen_operativo / ventas if ventas else 0.0
        lines.extend(
            [
                "ECUACIONES NUMERICAS",
                f"margen_bruto = ventas_netas_soles - materia_prima_directa - costos_conversion = {self.format_number(ventas)} - {self.format_number(materia)} - {self.format_number(conversion)} = {self.format_number(calculated_margen_bruto)}",
                f"Resultado guardado margen_bruto = {self.format_number(margen_bruto)}",
                f"margen_operativo = margen_bruto - procesos_soporte = {self.format_number(margen_bruto)} - {self.format_number(soporte)} = {self.format_number(calculated_margen_operativo)}",
                f"Resultado guardado margen_operativo = {self.format_number(margen_operativo)}",
                f"rentabilidad_operativa = margen_operativo / ventas_netas_soles = {self.format_number(margen_operativo)} / {self.format_number(ventas)} = {self.format_number(calculated_rentabilidad)}",
                f"Resultado guardado rentabilidad_operativa = {self.format_number(rentabilidad)}",
                "",
            ]
        )
        return lines

    def fetch_all_dicts(self, conn, sql: str, params: tuple = ()) -> list[dict]:
        return [dict(item) for item in conn.execute(sql, params).fetchall()]

    def add_section(self, lines: list[str], title: str, data) -> None:
        lines.append(title.upper())
        if not data:
            lines.append("Sin registros.")
            lines.append("")
            return
        rows = data if isinstance(data, list) else [data]
        for index, row in enumerate(rows, start=1):
            if len(rows) > 1:
                lines.append(f"Registro {index}:")
            for field, value in row.items():
                lines.append(f"  {field}: {self.format_detail_value(value)}")
            lines.append("")

    def as_number(self, row: dict, field: str) -> float:
        if not row:
            return 0.0
        try:
            return float(row.get(field) or 0)
        except (TypeError, ValueError):
            return 0.0

    def format_number(self, value: float) -> str:
        text = f"{value:,.6f}".rstrip("0").rstrip(".")
        return text if text else "0"

    def format_detail_value(self, value) -> str:
        if isinstance(value, float):
            return self.format_number(value)
        if value is None:
            return ""
        return str(value)

    def refresh_dashboard_text(self) -> None:
        payload = dashboard_payload(self.db_path)
        self.clear_dashboard_chart()
        if not payload:
            self.set_dashboard_kpis({})
            self.update_dashboard_insights({})
            if self.dashboard_header_status is not None:
                self.dashboard_header_status.configure(text="Pendiente de calculo", bg="#1E293B")
            self.update_dashboard_period_label([])
            if self.dashboard_message is not None:
                self.dashboard_message.configure(text="Ejecute el calculo del Modelo ABC por Ordenes de Trabajo para ver el dashboard.")
            return
        costs = payload.get("costos_ot", [])
        if self.dashboard_header_status is not None:
            self.dashboard_header_status.configure(text=f"Actualizado | {len(costs)} OT", bg="#0F766E")
        self.update_dashboard_period_label(payload.get("periodo", []))
        if self.dashboard_message is not None:
            self.dashboard_message.configure(text="Resumen ejecutivo de costos, actividades, estructura y rentabilidad")
        self.set_dashboard_kpis(payload)
        self.update_dashboard_insights(payload)
        if self.dashboard_chart_frame is None:
            return
        figure = self.build_dashboard_figure(payload)
        self.dashboard_chart_canvas = FigureCanvasTkAgg(figure, master=self.dashboard_chart_frame)
        self.dashboard_chart_canvas.draw()
        self.dashboard_chart_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def clear_dashboard_chart(self) -> None:
        if self.dashboard_chart_canvas is not None:
            self.dashboard_chart_canvas.get_tk_widget().destroy()
            self.dashboard_chart_canvas = None

    def update_dashboard_period_label(self, periods: list[dict]) -> None:
        if self.dashboard_period_label is None:
            return
        base = "Frenosa | Modelo ABC por Ordenes de Trabajo"
        period_text = self.format_dashboard_period(periods)
        text = f"{base} | {period_text}" if period_text else f"{base} | Periodo pendiente"
        self.dashboard_period_label.configure(text=text)

    def format_dashboard_period(self, periods: list[dict]) -> str:
        if not periods:
            return ""
        labels = []
        for row in periods:
            name = row.get("periodo") or row.get("periodo_id") or "Periodo"
            period_type = row.get("tipo_periodo") or ""
            start = row.get("fecha_inicio") or ""
            end = row.get("fecha_fin") or ""
            date_range = f"{start} a {end}" if start and end else start or end
            pieces = [str(name)]
            if period_type:
                pieces.append(str(period_type))
            if date_range:
                pieces.append(str(date_range))
            labels.append(" | ".join(pieces))
        return "Periodo: " + "; ".join(labels)

    def set_dashboard_kpis(self, payload: dict) -> None:
        costs = payload.get("costos_ot", []) if payload else []
        estado = payload.get("estado_resultados", []) if payload else []
        total_cost = sum(self.as_number(row, "costo_total_ot") for row in costs)
        avg_unit = sum(self.as_number(row, "costo_unitario_ot") for row in costs) / len(costs) if costs else 0.0
        sales = sum(self.as_number(row, "ventas_netas_soles") for row in estado)
        avg_margin = (
            sum(self.as_number(row, "rentabilidad_operativa") for row in estado) / len(estado)
            if estado
            else 0.0
        )
        values = {
            "total_cost": f"S/ {total_cost:,.2f}",
            "avg_unit": f"S/ {avg_unit:,.4f}",
            "sales": f"S/ {sales:,.2f}",
            "margin": f"{avg_margin * 100:,.2f}%",
        }
        for key, value in values.items():
            label = self.dashboard_kpis.get(key)
            if label is not None:
                label.configure(text=value)

    def update_dashboard_insights(self, payload: dict) -> None:
        if self.dashboard_insights_text is None:
            return
        self.dashboard_insights_text.configure(state=tk.NORMAL)
        self.dashboard_insights_text.delete("1.0", tk.END)
        self.dashboard_insights_text.insert("1.0", "\n".join(self.build_dashboard_insights_lines(payload)))
        self.dashboard_insights_text.configure(state=tk.DISABLED)

    def build_dashboard_insights_lines(self, payload: dict) -> list[str]:
        if not payload:
            return ["Sin datos calculados para alertas y recomendaciones."]
        costs = payload.get("costos_ot", [])
        activities = payload.get("costos_actividad", [])
        estado = payload.get("estado_resultados", [])
        alerts = payload.get("alertas", [])
        lines = ["ALERTAS"]
        if alerts:
            for alert in alerts:
                value = self.as_number(alert, "valor_referencia")
                value_text = f" | valor: {self.format_number(value)}" if value else ""
                observation = f" ({alert.get('observacion')})" if alert.get("observacion") else ""
                lines.append(f"- [{alert.get('severidad')}] {alert.get('mensaje')}{value_text}{observation}")
        else:
            lines.append("- Sin alertas generadas para la ejecucion actual.")

        recommendations = self.dashboard_recommendations(costs, activities, estado)
        lines.append("")
        lines.append("RECOMENDACIONES")
        if recommendations:
            lines.extend(f"- {item}" for item in recommendations)
        else:
            lines.append("- Ejecute el calculo con datos completos para obtener recomendaciones.")
        return lines

    def dashboard_recommendations(self, costs: list[dict], activities: list[dict], estado: list[dict]) -> list[str]:
        recommendations: list[str] = []
        if activities:
            total_activity_cost = sum(self.as_number(row, "costo_total_ac") for row in activities)
            top_activity = max(activities, key=lambda row: self.as_number(row, "costo_total_ac"))
            top_value = self.as_number(top_activity, "costo_total_ac")
            share = top_value / total_activity_cost if total_activity_cost else 0.0
            recommendations.append(
                f"Priorizar revision de la actividad {top_activity.get('nombre_actividad')}: concentra {share * 100:,.2f}% del costo total de actividades."
            )
        if estado:
            lowest_margin = min(estado, key=lambda row: self.as_number(row, "rentabilidad_operativa"))
            margin = self.as_number(lowest_margin, "rentabilidad_operativa")
            recommendations.append(
                f"Analizar la OT/producto {lowest_margin.get('codigo_producto') or lowest_margin.get('ot_id')}: tiene la menor rentabilidad operativa ({margin * 100:,.2f}%)."
            )
        if costs:
            highest_cost = max(costs, key=lambda row: self.as_number(row, "costo_total_ot"))
            total = self.as_number(highest_cost, "costo_total_ot")
            components = {
                "CMD": self.as_number(highest_cost, "cmd_ot"),
                "Produccion": self.as_number(highest_cost, "costo_produccion_ot"),
                "Soporte": self.as_number(highest_cost, "costo_soporte_ot"),
                "CSD": self.as_number(highest_cost, "csd_ot"),
            }
            dominant_name, dominant_value = max(components.items(), key=lambda item: item[1])
            dominant_share = dominant_value / total if total else 0.0
            recommendations.append(
                f"Revisar la estructura de la OT/producto {highest_cost.get('codigo_producto') or highest_cost.get('codigo_ot')}: el componente dominante es {dominant_name} ({dominant_share * 100:,.2f}% del costo total)."
            )
            support_candidates = [
                row for row in costs if self.as_number(row, "costo_total_ot") and self.as_number(row, "costo_soporte_ot")
            ]
            if support_candidates:
                highest_support = max(
                    support_candidates,
                    key=lambda row: self.as_number(row, "costo_soporte_ot") / self.as_number(row, "costo_total_ot"),
                )
                support_share = self.as_number(highest_support, "costo_soporte_ot") / self.as_number(highest_support, "costo_total_ot")
                if support_share >= 0.20:
                    recommendations.append(
                        f"Evaluar procesos de soporte en {highest_support.get('codigo_producto') or highest_support.get('codigo_ot')}: representan {support_share * 100:,.2f}% del costo total."
                    )
            if all(self.as_number(row, "csd_ot") == 0 for row in costs):
                recommendations.append(
                    "Mantener documentado CSD = 0 mientras el Excel maestro no registre servicios directos; no inventar servicios no sustentados."
                )
        return recommendations

    def build_dashboard_figure(self, payload: dict) -> Figure:
        costs = payload.get("costos_ot", [])
        activities = payload.get("costos_actividad", [])
        estado = payload.get("estado_resultados", [])
        figure = Figure(figsize=(11.2, 4.8), dpi=100, facecolor="#FFFFFF", constrained_layout=True)
        axes = figure.subplots(2, 2)
        ax_cost, ax_mix, ax_activity, ax_margin = axes.flatten()

        self.style_dashboard_axis(ax_cost, "Costo total por OT")
        labels = [str(row.get("codigo_producto") or row.get("codigo_ot") or row.get("ot_id")) for row in costs]
        total_costs = [self.as_number(row, "costo_total_ot") for row in costs]
        if total_costs:
            bars = ax_cost.bar(labels, total_costs, color="#1D4ED8", width=0.62)
            self.annotate_bars(ax_cost, bars, currency=True)
            ax_cost.tick_params(axis="x", rotation=20)
            ax_cost.set_ylabel("Soles")
        else:
            self.empty_axis(ax_cost)

        self.style_dashboard_axis(ax_mix, "Estructura de costo por OT")
        if costs:
            x_positions = range(len(labels))
            components = [
                ("CMD", [self.as_number(row, "cmd_ot") for row in costs], "#0F766E"),
                ("Produccion", [self.as_number(row, "costo_produccion_ot") for row in costs], "#D97706"),
                ("Soporte", [self.as_number(row, "costo_soporte_ot") for row in costs], "#7C3AED"),
                ("CSD", [self.as_number(row, "csd_ot") for row in costs], "#64748B"),
            ]
            bottoms = [0.0 for _ in costs]
            for name, values, color in components:
                ax_mix.bar(x_positions, values, bottom=bottoms, label=name, color=color)
                bottoms = [base + value for base, value in zip(bottoms, values)]
            ax_mix.set_xticks(list(x_positions), labels, rotation=20)
            ax_mix.set_ylabel("Soles")
            ax_mix.legend(loc="upper left", fontsize=8, ncols=2)
        else:
            self.empty_axis(ax_mix)

        self.style_dashboard_axis(ax_activity, "Ranking de actividades")
        top_activities = sorted(activities, key=lambda row: self.as_number(row, "costo_total_ac"), reverse=True)[:7]
        if top_activities:
            names = [str(row.get("nombre_actividad") or row.get("codigo_actividad")) for row in top_activities]
            values = [self.as_number(row, "costo_total_ac") for row in top_activities]
            bars = ax_activity.barh(names, values, color="#059669")
            ax_activity.invert_yaxis()
            self.annotate_horizontal_bars(ax_activity, bars)
            ax_activity.set_xlabel("Soles")
        else:
            self.empty_axis(ax_activity)

        self.style_dashboard_axis(ax_margin, "Rentabilidad operativa por OT")
        margin_labels = [str(row.get("codigo_producto") or row.get("ot_id")) for row in estado]
        margins = [self.as_number(row, "rentabilidad_operativa") * 100 for row in estado]
        if margins:
            colors = ["#DC2626" if value < 0 else "#0891B2" for value in margins]
            bars = ax_margin.bar(margin_labels, margins, color=colors)
            self.annotate_bars(ax_margin, bars, suffix="%")
            ax_margin.axhline(0, color="#334155", linewidth=0.8)
            ax_margin.tick_params(axis="x", rotation=20)
            ax_margin.set_ylabel("%")
        else:
            self.empty_axis(ax_margin)
        return figure

    def style_dashboard_axis(self, axis, title: str) -> None:
        axis.set_title(title, fontsize=10.5, fontweight="bold", color="#0F172A", loc="left")
        axis.set_facecolor("#FFFFFF")
        axis.grid(axis="y", color="#E2E8F0", linewidth=0.8)
        axis.set_axisbelow(True)
        axis.tick_params(colors="#334155", labelsize=8)
        axis.yaxis.label.set_color("#334155")
        axis.xaxis.label.set_color("#334155")
        axis.spines["top"].set_visible(False)
        axis.spines["right"].set_visible(False)
        axis.spines["left"].set_color("#CBD5E1")
        axis.spines["bottom"].set_color("#CBD5E1")

    def empty_axis(self, axis) -> None:
        axis.text(0.5, 0.5, "Sin datos", ha="center", va="center", color="#64748B", transform=axis.transAxes)
        axis.set_xticks([])
        axis.set_yticks([])

    def annotate_bars(self, axis, bars, currency: bool = False, suffix: str = "") -> None:
        for bar in bars:
            value = bar.get_height()
            label = f"S/ {value:,.0f}" if currency else f"{value:,.1f}{suffix}"
            axis.annotate(
                label,
                xy=(bar.get_x() + bar.get_width() / 2, value),
                xytext=(0, 4),
                textcoords="offset points",
                ha="center",
                va="bottom",
                fontsize=8,
                color="#334155",
            )

    def annotate_horizontal_bars(self, axis, bars) -> None:
        for bar in bars:
            value = bar.get_width()
            axis.annotate(
                f"S/ {value:,.0f}",
                xy=(value, bar.get_y() + bar.get_height() / 2),
                xytext=(4, 0),
                textcoords="offset points",
                ha="left",
                va="center",
                fontsize=8,
                color="#334155",
            )

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
