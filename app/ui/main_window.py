from __future__ import annotations

import calendar
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

from ..constants import (
    APP_TITLE,
    DEFAULT_DB_PATH,
    IMAGES_MODELS_DIR,
    MODEL_1_VISUAL_NAME,
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
    "Periodo",
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
        self.after(0, self.maximize_window)
        self.db_path = Path(db_path)
        self.excel_path = tk.StringVar(value="")
        self.status_text = tk.StringVar(value="Seleccione, valide e importe un Excel maestro para iniciar.")
        self.model_var = tk.StringVar(value="")
        self.period_var = tk.StringVar(value="")
        self.table_var = tk.StringVar(value=IMPORTED_TABLES[0])
        self.indirect_cost_table_var = tk.StringVar(value=INDIRECT_COST_TABLES[0])
        self.model_statuses: list[CostingModelStatus] = []
        self.period_options: list[dict] = []
        self.active_period: dict | None = None
        self.trees: dict[str, ttk.Treeview] = {}
        self.tree_row_data: dict[str, dict[str, dict]] = {}
        self.module_tabs: list[tuple[ttk.Frame, str]] = []
        self.dashboard_chart_canvas = None
        self.dashboard_scroll_canvas: tk.Canvas | None = None
        self.dashboard_scroll_content: ttk.Frame | None = None
        self.dashboard_kpis: dict[str, tk.Label] = {}
        self.dashboard_header_status: tk.Label | None = None
        self.dashboard_period_label: tk.Label | None = None
        self.dashboard_message: ttk.Label | None = None
        self.dashboard_chart_frame = None
        self.dashboard_insights_text: tk.Text | None = None
        self.proposal_text: tk.Text | None = None
        self.criteria_image: tk.PhotoImage | None = None
        self.criteria_source_image: tk.PhotoImage | None = None
        self.criteria_image_path: Path | None = None
        self.criteria_image_model_code: str | None = None
        self.is_excel_validated = False
        self.is_excel_imported = False
        self.validated_excel_path: str | None = None
        self.imported_excel_path: str | None = None
        self._configure_style()
        self._build_layout()
        initialize_database(self.db_path)
        self.reset_model_selector()

    def maximize_window(self) -> None:
        try:
            self.state("zoomed")
        except tk.TclError:
            try:
                self.attributes("-zoomed", True)
            except tk.TclError:
                self.geometry(f"{self.winfo_screenwidth()}x{self.winfo_screenheight()}+0+0")

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TButton", padding=6)
        style.configure("Title.TLabel", font=("Segoe UI", 16, "bold"), foreground="#1F4D78")
        style.configure("Subtitle.TLabel", font=("Segoe UI", 10), foreground="#334155")
        style.configure("Section.TLabelframe.Label", font=("Segoe UI", 10, "bold"))
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))
        style.configure("Treeview", rowheight=24)
        style.configure(
            "Dashboard.Vertical.TScrollbar",
            background="#2563EB",
            bordercolor="#93C5FD",
            arrowcolor="#0F172A",
            darkcolor="#1D4ED8",
            lightcolor="#60A5FA",
            troughcolor="#DBEAFE",
            width=18,
            arrowsize=14,
        )
        style.map(
            "Dashboard.Vertical.TScrollbar",
            background=[("active", "#1D4ED8"), ("pressed", "#0F766E")],
        )

    def _build_layout(self) -> None:
        root = ttk.Frame(self, padding=12)
        root.pack(fill=tk.BOTH, expand=True)

        self.header_frame = ttk.Frame(root)
        self.header_frame.pack(fill=tk.X)
        self.header_frame.columnconfigure(0, weight=1)
        ttk.Label(self.header_frame, text="Sistema Informatico de Costos ABC - Frenosa", style="Title.TLabel").grid(
            row=0,
            column=0,
            sticky=tk.W,
        )
        ttk.Label(
            self.header_frame,
            text="Modelo ABC por Ordenes de Trabajo | Datos importados, calculos generados y reportes exportados.",
            style="Subtitle.TLabel",
        ).grid(row=1, column=0, sticky=tk.W, pady=(2, 8))
        self.back_home_button = ttk.Button(
            self.header_frame,
            text="Volver a pantalla inicial",
            command=self.show_home_screen,
        )
        self.back_home_button.grid(row=0, column=1, rowspan=2, sticky=tk.E, padx=(12, 0))
        self.back_home_button.grid_remove()

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
        self._build_improvement_proposal_tab()
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

        self.period_box = ttk.LabelFrame(tab, text="Seleccion de periodo", padding=10, style="Section.TLabelframe")
        self.period_box.pack(fill=tk.X, pady=(12, 0))
        ttk.Label(self.period_box, text="Periodo activo:").grid(row=0, column=0, sticky=tk.W)
        self.period_combo = ttk.Combobox(self.period_box, state="disabled", textvariable=self.period_var, width=72)
        self.period_combo.grid(row=0, column=1, sticky=tk.EW, padx=8)
        self.period_combo.bind("<<ComboboxSelected>>", lambda _event: self.select_active_period())
        self.create_period_button = ttk.Button(
            self.period_box,
            text="Crear nuevo periodo",
            command=self.open_create_period_dialog,
            state=tk.DISABLED,
        )
        self.create_period_button.grid(row=0, column=2, padx=4)
        self.period_status = ttk.Label(
            self.period_box,
            text="Importe el Excel maestro para seleccionar o crear un periodo.",
            style="Subtitle.TLabel",
        )
        self.period_status.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(8, 0))
        self.period_box.columnconfigure(1, weight=1)

        self.model_box = ttk.LabelFrame(tab, text="Seleccion de modelo de costeo", padding=10, style="Section.TLabelframe")
        self.model_box.pack(fill=tk.X, pady=12)
        self.model_combo = ttk.Combobox(self.model_box, state="readonly", textvariable=self.model_var, width=72)
        self.model_combo.grid(row=0, column=0, sticky=tk.W)
        self.model_combo.bind("<<ComboboxSelected>>", lambda _event: self.show_selected_model_reason())
        self.execute_button = ttk.Button(self.model_box, text="Ejecutar calculo", command=self.execute_calculation)
        self.execute_button.grid(row=0, column=1, padx=8)
        self.model_reason = tk.Text(self.model_box, height=5, wrap=tk.WORD)
        self.model_reason.grid(row=1, column=0, columnspan=2, sticky=tk.EW, pady=(8, 0))
        self.model_box.columnconfigure(0, weight=1)

        self.criteria_box = ttk.LabelFrame(tab, text="Grafico del modelo de costeo", padding=10, style="Section.TLabelframe")
        self.criteria_box.pack(fill=tk.BOTH, expand=True, pady=12)
        self.criteria_canvas = tk.Canvas(
            self.criteria_box,
            bg="#FFFFFF",
            height=430,
            highlightthickness=1,
            highlightbackground="#CBD5E1",
        )
        self.criteria_canvas.pack(fill=tk.BOTH, expand=True)
        self.criteria_canvas.bind("<Configure>", lambda _event: self.render_model_diagram())
        self.hide_model_selection()
        return tab

    def _build_data_tab(self) -> None:
        tab = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(tab, text="Datos importados")
        self.module_tabs.append((tab, "Datos importados"))
        top = ttk.Frame(tab)
        top.pack(fill=tk.X)
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
        tab = ttk.Frame(self.notebook, padding=0)
        self.notebook.add(tab, text="Dashboard")
        self.module_tabs.append((tab, "Dashboard"))
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(0, weight=1)

        dashboard_canvas = tk.Canvas(tab, bg="#F8FAFC", highlightthickness=0)
        dashboard_scroll = ttk.Scrollbar(
            tab,
            orient=tk.VERTICAL,
            command=dashboard_canvas.yview,
            style="Dashboard.Vertical.TScrollbar",
        )
        dashboard_canvas.configure(yscrollcommand=dashboard_scroll.set)
        dashboard_canvas.grid(row=0, column=0, sticky="nsew")
        dashboard_scroll.grid(row=0, column=1, sticky="ns")
        self.dashboard_scroll_canvas = dashboard_canvas

        content = ttk.Frame(dashboard_canvas, padding=12)
        content_window = dashboard_canvas.create_window((0, 0), window=content, anchor="nw")
        content.columnconfigure(0, weight=1)
        content.rowconfigure(3, weight=1, minsize=620)
        self.dashboard_scroll_content = content

        def update_dashboard_scrollregion(_event=None) -> None:
            self.refresh_dashboard_scrollregion()

        def resize_dashboard_content(event) -> None:
            dashboard_canvas.itemconfigure(content_window, width=event.width)
            self.refresh_dashboard_scrollregion()

        content.bind("<Configure>", update_dashboard_scrollregion)
        dashboard_canvas.bind("<Configure>", resize_dashboard_content)
        dashboard_canvas.bind("<MouseWheel>", self.scroll_dashboard, add="+")

        header = tk.Frame(content, bg="#0F172A", padx=16, pady=12)
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

        kpi_frame = tk.Frame(content, bg="#F8FAFC")
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

        self.dashboard_message = ttk.Label(content, text="", style="Subtitle.TLabel")
        self.dashboard_message.grid(row=2, column=0, sticky="w", pady=(8, 4))
        self.dashboard_chart_frame = tk.Frame(content, bg="#FFFFFF", highlightthickness=1, highlightbackground="#CBD5E1")
        self.dashboard_chart_frame.grid(row=3, column=0, sticky="nsew", pady=(2, 0))
        insights = ttk.LabelFrame(content, text="Alertas y recomendaciones", padding=8, style="Section.TLabelframe")
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
        self.dashboard_insights_text.grid(row=0, column=0, sticky="ew")
        insights.columnconfigure(0, weight=1)
        self.bind_dashboard_scroll(content)

    def scroll_dashboard(self, event) -> str:
        if self.dashboard_scroll_canvas is None:
            return "break"
        delta = getattr(event, "delta", 0)
        if delta:
            steps = int(-1 * (delta / 120))
            if steps == 0:
                steps = -1 if delta > 0 else 1
            self.dashboard_scroll_canvas.yview_scroll(steps, "units")
        return "break"

    def bind_dashboard_scroll(self, widget) -> None:
        widget.bind("<MouseWheel>", self.scroll_dashboard, add="+")
        for child in widget.winfo_children():
            self.bind_dashboard_scroll(child)

    def refresh_dashboard_scrollregion(self) -> None:
        if self.dashboard_scroll_canvas is None:
            return
        self.dashboard_scroll_canvas.configure(scrollregion=self.dashboard_scroll_canvas.bbox("all"))

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
        self.notebook.add(tab, text="Reportes")
        self.module_tabs.append((tab, "Reportes"))
        ttk.Label(tab, text=f"Reportes se guardan en: {REPORTS_DIR}", style="Subtitle.TLabel").pack(anchor=tk.W)
        actions = ttk.Frame(tab)
        actions.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(actions, text="Generar reportes Excel", command=self.generate_reports).pack(side=tk.LEFT, padx=4)
        ttk.Button(actions, text="Generar documentos Word", command=self.generate_documents).pack(side=tk.LEFT, padx=4)
        ttk.Button(actions, text="Generar dashboard PNG", command=self.generate_dashboard_png).pack(side=tk.LEFT, padx=4)
        ttk.Button(actions, text="Ejecutar informe de pruebas", command=self.run_tests).pack(side=tk.LEFT, padx=4)
        self.trees["salidas"] = self._make_tree(tab)

    def _build_improvement_proposal_tab(self) -> None:
        tab = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(tab, text="Propuesta de mejoras")
        self.module_tabs.append((tab, "Propuesta de mejoras"))
        tab.rowconfigure(2, weight=1)
        tab.columnconfigure(0, weight=1)

        ttk.Label(tab, text="Informe editable de propuestas de mejora", style="Subtitle.TLabel").grid(row=0, column=0, sticky=tk.W)
        actions = ttk.Frame(tab)
        actions.grid(row=1, column=0, sticky=tk.EW, pady=(8, 0))
        ttk.Button(actions, text="Generar propuesta base", command=self.generate_improvement_proposal).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(actions, text="Guardar propuesta editada", command=self.save_improvement_proposal).pack(side=tk.LEFT, padx=6)

        editor = ttk.Frame(tab)
        editor.grid(row=2, column=0, sticky="nsew", pady=(8, 0))
        editor.rowconfigure(0, weight=1)
        editor.columnconfigure(0, weight=1)
        self.proposal_text = tk.Text(
            editor,
            wrap=tk.WORD,
            bg="#FFFFFF",
            fg="#0F172A",
            font=("Segoe UI", 10),
            undo=True,
        )
        scroll = ttk.Scrollbar(editor, orient=tk.VERTICAL, command=self.proposal_text.yview)
        self.proposal_text.configure(yscrollcommand=scroll.set)
        self.proposal_text.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")
        self.reset_improvement_proposal()

    def reset_improvement_proposal(self) -> None:
        if self.proposal_text is None:
            return
        self.proposal_text.delete("1.0", tk.END)
        self.proposal_text.insert("1.0", "Ejecute el calculo ABC y genere la propuesta base.")
        self.proposal_text.edit_reset()

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
            self.reset_period_selector()
            self.reset_model_selector()
            self.show_home_screen()
            self.status_text.set("Excel seleccionado. Valide la estructura antes de importar.")

    def validate_excel(self) -> None:
        if not self.excel_path.get().strip():
            messagebox.showwarning("Validacion", "Primero seleccione un Excel maestro.")
            self.status_text.set("No hay Excel seleccionado.")
            return
        try:
            result = validate_workbook(self.excel_path.get())
        except Exception as exc:
            self.is_excel_validated = False
            self.validated_excel_path = None
            messagebox.showerror("Validacion", f"No se pudo validar el Excel maestro:\n{exc}")
            self.status_text.set(f"Error de validacion: {exc}")
            return
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
                self.reset_model_selector()
                self.hide_model_selection()
                self.load_period_options()
                self.load_selected_imported_table()
                messagebox.showinfo("Importacion", "Excel importado correctamente a SQLite local.")
            else:
                self.is_excel_imported = False
                self.imported_excel_path = None
                self.reset_period_selector()
                self.hide_model_selection()
                self.reset_model_selector()
                messagebox.showerror("Importacion", "\n".join(issue.message for issue in result.errors[:8]))
            self.status_text.set(f"Importacion batch {result.batch_id}: {result.imported_rows}")
        except Exception as exc:
            self.is_excel_imported = False
            self.imported_excel_path = None
            self.reset_period_selector()
            self.hide_model_selection()
            self.reset_model_selector()
            messagebox.showerror("Importacion", str(exc))
            self.status_text.set(f"Error de importacion: {exc}")

    def refresh_model_statuses(self) -> None:
        if not self.is_excel_imported:
            self.reset_model_selector()
            return
        if self.active_period is None:
            self.reset_model_selector()
            self.hide_model_selection()
            return
        if self.active_period.get("origen_periodo") == "MANUAL":
            self.model_statuses = [self.manual_period_model_1_status()]
        else:
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
        self.show_model_selection()
        self.show_selected_model_reason()

    def reset_period_selector(self) -> None:
        self.period_options = []
        self.active_period = None
        self.period_var.set("")
        if hasattr(self, "period_combo"):
            self.period_combo.configure(values=[], state=tk.DISABLED)
        if hasattr(self, "create_period_button"):
            self.create_period_button.configure(state=tk.DISABLED)
        if hasattr(self, "period_status"):
            self.period_status.configure(text="Importe el Excel maestro para seleccionar o crear un periodo.")
        self.reset_improvement_proposal()

    def load_period_options(self) -> None:
        with db_session(self.db_path) as conn:
            rows = conn.execute(
                """
                SELECT
                    periodo_id,
                    empresa_id,
                    periodo,
                    tipo_periodo,
                    fecha_inicio,
                    fecha_fin,
                    descripcion,
                    fuente,
                    COALESCE(origen_periodo, 'IMPORTADO') AS origen_periodo
                FROM periodo
                ORDER BY fecha_inicio, periodo_id
                """
            ).fetchall()
        self.period_options = []
        for row in rows:
            item = dict(row)
            item["display"] = self.format_period_option(item)
            self.period_options.append(item)
        values = [item["display"] for item in self.period_options]
        self.period_combo.configure(values=values, state="readonly" if values else tk.DISABLED)
        self.create_period_button.configure(state=tk.NORMAL)
        self.active_period = None
        self.period_var.set("")
        self.hide_model_selection()
        self.reset_improvement_proposal()
        if values:
            self.period_status.configure(text="Seleccione un periodo existente o cree un nuevo periodo.")
        else:
            self.period_status.configure(text="No hay periodos cargados. Cree un nuevo periodo para continuar.")

    def format_period_option(self, period: dict) -> str:
        code = period.get("periodo_id") or ""
        name = period.get("periodo") or "Periodo sin nombre"
        period_type = period.get("tipo_periodo") or ""
        start = period.get("fecha_inicio") or ""
        end = period.get("fecha_fin") or ""
        origin = "Manual" if period.get("origen_periodo") == "MANUAL" else "Excel"
        date_range = f"{start} a {end}" if start or end else "sin fechas"
        return f"{code} | {name} | {period_type} | {date_range} | {origin}"

    def select_active_period(self) -> None:
        selected = self.period_var.get()
        self.active_period = next((item for item in self.period_options if item["display"] == selected), None)
        if self.active_period is None:
            self.hide_model_selection()
            self.reset_model_selector()
            self.reset_improvement_proposal()
            self.period_status.configure(text="Seleccione un periodo valido para continuar.")
            return
        self.reset_improvement_proposal()
        self.period_status.configure(text=f"Periodo activo: {selected}")
        self.refresh_model_statuses()
        self.status_text.set(f"Periodo activo seleccionado: {selected}")

    def manual_period_model_1_status(self) -> CostingModelStatus:
        return CostingModelStatus(
            code="MODELO_1",
            label="Modelo 1",
            visual_name=MODEL_1_VISUAL_NAME,
            status="Disponible",
            available=True,
            reason=(
                "Disponible para periodo creado manualmente: se considera informacion estimada o proyectada, "
                "tomando como referencia la estructura y datos base del Excel maestro importado. "
                "No se muestran MODELO_2 ni MODELO_3 porque el periodo nuevo no tiene datos suficientes para modelos mas complejos."
            ),
            flow="Proceso 1: EC -> AC; Proceso 2 AC-OT (Produccion); Proceso 2 AC-OT (Soporte)",
        )

    def open_create_period_dialog(self) -> None:
        if not self.is_excel_imported:
            messagebox.showwarning("Periodo", "Primero importe el Excel maestro a SQLite.")
            return
        dialog = tk.Toplevel(self)
        dialog.title("Crear nuevo periodo")
        dialog.transient(self)
        dialog.grab_set()
        dialog.resizable(False, False)
        fields = {
            "periodo_id": tk.StringVar(),
            "periodo": tk.StringVar(),
            "tipo_periodo": tk.StringVar(value="mensual"),
            "fecha_inicio": tk.StringVar(),
            "fecha_fin": tk.StringVar(),
            "descripcion": tk.StringVar(value="Periodo creado manualmente para estimacion o proyeccion."),
        }
        labels = [
            ("Codigo del periodo", "periodo_id"),
            ("Nombre del periodo", "periodo"),
            ("Tipo de periodo", "tipo_periodo"),
            ("Fecha de inicio (YYYY-MM-DD)", "fecha_inicio"),
            ("Fecha de fin (YYYY-MM-DD)", "fecha_fin"),
            ("Descripcion", "descripcion"),
        ]
        body = ttk.Frame(dialog, padding=12)
        body.grid(row=0, column=0, sticky="nsew")
        for row, (label, key) in enumerate(labels):
            ttk.Label(body, text=label + ":").grid(row=row, column=0, sticky=tk.W, pady=4)
            if key == "tipo_periodo":
                widget = ttk.Combobox(
                    body,
                    textvariable=fields[key],
                    values=["mensual", "bimestral", "trimestral", "cuatrimestral", "semestral", "anual"],
                    state="readonly",
                    width=32,
                )
            elif key in {"fecha_inicio", "fecha_fin"}:
                widget = ttk.Frame(body)
                entry = ttk.Entry(widget, textvariable=fields[key], width=32)
                entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
                ttk.Button(
                    widget,
                    text="Calendario",
                    command=lambda field=fields[key]: self.open_calendar_picker(dialog, field),
                ).pack(side=tk.LEFT, padx=(6, 0))
            else:
                widget = ttk.Entry(body, textvariable=fields[key], width=44)
            widget.grid(row=row, column=1, sticky=tk.EW, padx=(8, 0), pady=4)
        body.columnconfigure(1, weight=1)
        actions = ttk.Frame(body)
        actions.grid(row=len(labels), column=0, columnspan=2, sticky=tk.E, pady=(10, 0))
        ttk.Button(actions, text="Cancelar", command=dialog.destroy).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(actions, text="Guardar periodo", command=lambda: self.save_new_period(dialog, fields)).pack(side=tk.RIGHT)

    def open_calendar_picker(self, parent: tk.Toplevel, target_var: tk.StringVar) -> None:
        try:
            current = datetime.strptime(target_var.get(), "%Y-%m-%d").date()
        except ValueError:
            current = datetime.now().date()
        state = {"year": current.year, "month": current.month}
        picker = tk.Toplevel(parent)
        picker.title("Seleccionar fecha")
        picker.transient(parent)
        picker.grab_set()
        picker.resizable(False, False)
        content = ttk.Frame(picker, padding=10)
        content.grid(row=0, column=0)
        month_names = [
            "Enero",
            "Febrero",
            "Marzo",
            "Abril",
            "Mayo",
            "Junio",
            "Julio",
            "Agosto",
            "Septiembre",
            "Octubre",
            "Noviembre",
            "Diciembre",
        ]
        month_var = tk.StringVar(value=month_names[current.month - 1])
        year_var = tk.StringVar(value=str(current.year))
        header = ttk.Frame(content)
        header.grid(row=0, column=0, columnspan=7, sticky=tk.EW)
        ttk.Button(header, text="<", width=3, command=lambda: move_month(-1)).pack(side=tk.LEFT)
        month_combo = ttk.Combobox(header, textvariable=month_var, values=month_names, state="readonly", width=13)
        month_combo.pack(side=tk.LEFT, padx=(8, 4))
        year_spin = tk.Spinbox(header, from_=1900, to=2100, width=6, textvariable=year_var, command=lambda: on_month_year_change())
        year_spin.pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(header, text=">", width=3, command=lambda: move_month(1)).pack(side=tk.LEFT)
        days_frame = ttk.Frame(content)
        days_frame.grid(row=1, column=0, columnspan=7, pady=(8, 0))
        month_combo.bind("<<ComboboxSelected>>", lambda _event: on_month_year_change())
        year_spin.bind("<Return>", lambda _event: on_month_year_change())
        year_spin.bind("<FocusOut>", lambda _event: on_month_year_change())

        def sync_controls() -> None:
            month_var.set(month_names[state["month"] - 1])
            year_var.set(str(state["year"]))

        def on_month_year_change() -> None:
            try:
                state["month"] = month_names.index(month_var.get()) + 1
                state["year"] = int(year_var.get())
            except (ValueError, TypeError):
                state["year"] = current.year
                state["month"] = current.month
            state["year"] = min(max(state["year"], 1900), 2100)
            render_days()

        def move_month(delta: int) -> None:
            month = state["month"] + delta
            year = state["year"]
            if month < 1:
                month = 12
                year -= 1
            elif month > 12:
                month = 1
                year += 1
            state["year"] = year
            state["month"] = month
            render_days()

        def select_day(day: int) -> None:
            target_var.set(f"{state['year']:04d}-{state['month']:02d}-{day:02d}")
            picker.destroy()

        def render_days() -> None:
            for child in days_frame.winfo_children():
                child.destroy()
            sync_controls()
            for col, label in enumerate(["Lu", "Ma", "Mi", "Ju", "Vi", "Sa", "Do"]):
                ttk.Label(days_frame, text=label, width=4, anchor=tk.CENTER).grid(row=0, column=col, padx=1, pady=1)
            for row, week in enumerate(calendar.monthcalendar(state["year"], state["month"]), start=1):
                for col, day in enumerate(week):
                    if day == 0:
                        ttk.Label(days_frame, text="", width=4).grid(row=row, column=col, padx=1, pady=1)
                    else:
                        ttk.Button(days_frame, text=str(day), width=4, command=lambda value=day: select_day(value)).grid(
                            row=row,
                            column=col,
                            padx=1,
                            pady=1,
                        )

        render_days()

    def save_new_period(self, dialog: tk.Toplevel, fields: dict[str, tk.StringVar]) -> None:
        data = {key: value.get().strip() for key, value in fields.items()}
        error = self.validate_new_period(data)
        if error:
            messagebox.showerror("Crear nuevo periodo", error)
            return
        with db_session(self.db_path) as conn:
            conn.execute(
                """
                INSERT INTO periodo
                    (periodo_id, empresa_id, periodo, tipo_periodo, fecha_inicio, fecha_fin,
                     descripcion, fuente, origen_periodo, batch_id)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, NULL)
                """,
                (
                    data["periodo_id"],
                    self.default_company_id(conn),
                    data["periodo"],
                    data["tipo_periodo"],
                    data["fecha_inicio"],
                    data["fecha_fin"],
                    data["descripcion"],
                    "Registro manual del usuario",
                    "MANUAL",
                ),
            )
        dialog.destroy()
        self.load_period_options()
        created = next((item for item in self.period_options if item.get("periodo_id") == data["periodo_id"]), None)
        if created:
            self.period_var.set(created["display"])
            self.select_active_period()
        messagebox.showinfo("Crear nuevo periodo", "Periodo creado correctamente.")

    def validate_new_period(self, data: dict[str, str]) -> str | None:
        required = ["periodo_id", "periodo", "tipo_periodo", "fecha_inicio", "fecha_fin"]
        missing = [field for field in required if not data.get(field)]
        if missing:
            return "Complete los campos obligatorios: " + ", ".join(missing) + "."
        with db_session(self.db_path) as conn:
            duplicate = conn.execute(
                "SELECT 1 FROM periodo WHERE periodo_id = ?",
                (data["periodo_id"],),
            ).fetchone()
        if duplicate:
            return f"El codigo de periodo {data['periodo_id']} ya existe."
        try:
            start = datetime.strptime(data["fecha_inicio"], "%Y-%m-%d").date()
            end = datetime.strptime(data["fecha_fin"], "%Y-%m-%d").date()
        except ValueError:
            return "Las fechas deben tener el formato YYYY-MM-DD."
        if start > end:
            return "La fecha de inicio no puede ser mayor que la fecha de fin."
        duration = (end - start).days + 1
        if not self.is_reasonable_period_duration(data["tipo_periodo"], duration):
            return (
                f"La duracion de {duration} dias no es razonable para un periodo "
                f"{data['tipo_periodo']}."
            )
        return None

    def is_reasonable_period_duration(self, period_type: str, duration_days: int) -> bool:
        ranges = {
            "mensual": (28, 31),
            "bimestral": (58, 62),
            "trimestral": (89, 92),
            "cuatrimestral": (119, 123),
            "semestral": (181, 184),
            "anual": (365, 366),
        }
        minimum, maximum = ranges.get(period_type, (0, 0))
        return minimum <= duration_days <= maximum

    def default_company_id(self, conn) -> str | None:
        row = conn.execute("SELECT empresa_id FROM empresas ORDER BY empresa_id LIMIT 1").fetchone()
        return row["empresa_id"] if row else None

    def selected_model(self) -> CostingModelStatus:
        if not self.model_statuses:
            raise RuntimeError("Primero seleccione un periodo activo para evaluar modelos.")
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
                "Los modelos se mostraran despues de importar el Excel maestro y seleccionar un periodo activo.",
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
            self.criteria_image_path = None
            self.criteria_canvas.delete("all")
            return
        path = IMAGES_MODELS_DIR / image_name
        if not path.exists():
            self.criteria_image_model_code = None
            self.criteria_source_image = None
            self.criteria_image = None
            self.criteria_image_path = None
            self.criteria_canvas.delete("all")
            return
        self.criteria_image_model_code = model_code
        self.criteria_image_path = path
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
        target_width = max(canvas_width - 10, 1)
        target_height = max(canvas_height - 10, 1)
        if self.criteria_image_path is not None:
            try:
                from PIL import Image, ImageTk

                with Image.open(self.criteria_image_path) as source:
                    scale = min(target_width / source.width, target_height / source.height, 1.08)
                    display_width = max(1, int(source.width * scale))
                    display_height = max(1, int(source.height * scale))
                    resized = source.resize((display_width, display_height), Image.LANCZOS)
                self.criteria_image = ImageTk.PhotoImage(resized)
                x = max((canvas_width - display_width) // 2, 5)
                y = max((canvas_height - display_height) // 2, 5)
                self.criteria_canvas.create_image(x, y, anchor=tk.NW, image=self.criteria_image)
                self.criteria_canvas.configure(scrollregion=(0, 0, canvas_width, canvas_height))
                return
            except Exception:
                pass
        factor = max(
            1,
            (image_width + target_width - 1) // target_width,
            (image_height + target_height - 1) // target_height,
        )
        self.criteria_image = self.criteria_source_image.subsample(factor, factor)
        x = max((canvas_width - self.criteria_image.width()) // 2, 5)
        y = max((canvas_height - self.criteria_image.height()) // 2, 5)
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
        if self.active_period is None:
            messagebox.showwarning("Calculo ABC", "Primero seleccione o cree un periodo activo.")
            self.status_text.set("No se puede ejecutar: falta seleccionar el periodo activo.")
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
            projection_note = (
                " Periodo manual: resultados tratados como referencia estimada/proyectada."
                if self.active_period.get("origen_periodo") == "MANUAL"
                else ""
            )
            self.status_text.set(
                f"Calculo Modelo ABC por Ordenes de Trabajo completado. Run {run_id}.{projection_note}"
            )
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
        if self.table_var.get() == "Periodo":
            return "periodo"
        return self.table_var.get()

    def load_selected_imported_table(self) -> None:
        table = self.selected_imported_table_name()
        try:
            with db_session(self.db_path) as conn:
                if table == "periodo":
                    if self.active_period is None:
                        rows = []
                    else:
                        rows = [
                            dict(row)
                            for row in conn.execute(
                                f"SELECT * FROM {quote_identifier(table)} WHERE periodo_id = ? LIMIT 500",
                                (self.active_period.get("periodo_id"),),
                            ).fetchall()
                        ]
                else:
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

    def hide_model_selection(self) -> None:
        if hasattr(self, "model_box"):
            self.model_box.pack_forget()
        if hasattr(self, "criteria_box"):
            self.criteria_box.pack_forget()

    def show_model_selection(self) -> None:
        if hasattr(self, "model_box") and not self.model_box.winfo_ismapped():
            self.model_box.pack(fill=tk.X, pady=12, after=self.period_box)
        if hasattr(self, "criteria_box") and not self.criteria_box.winfo_ismapped():
            self.criteria_box.pack(fill=tk.BOTH, expand=True, pady=12, after=self.model_box)

    def hide_module_tabs(self) -> None:
        for tab, _title in self.module_tabs:
            self.notebook.hide(tab)

    def show_modules_screen(self) -> None:
        for tab, title in self.module_tabs:
            self.notebook.add(tab, text=title)
        if hasattr(self, "back_home_button"):
            self.back_home_button.grid()
        self.notebook.hide(self.home_tab)
        if self.module_tabs:
            self.notebook.select(self.module_tabs[0][0])

    def show_home_screen(self) -> None:
        self.notebook.add(self.home_tab, text="Pantalla inicial")
        if hasattr(self, "back_home_button"):
            self.back_home_button.grid_remove()
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
        payload = self.active_dashboard_payload()
        self.clear_dashboard_chart()
        if not payload:
            self.set_dashboard_kpis({})
            self.update_dashboard_insights({})
            if self.dashboard_header_status is not None:
                self.dashboard_header_status.configure(text="Pendiente de calculo", bg="#1E293B")
            self.update_dashboard_period_label([])
            if self.dashboard_message is not None:
                self.dashboard_message.configure(text="Ejecute el calculo del Modelo ABC por Ordenes de Trabajo para ver el dashboard.")
            self.refresh_dashboard_scrollregion()
            return
        costs = payload.get("costos_ot", [])
        if self.dashboard_header_status is not None:
            status = "Proyectado" if payload.get("modo_periodo") == "PROYECTADO" else "Actualizado"
            self.dashboard_header_status.configure(text=f"{status} | {len(costs)} OT", bg="#0F766E")
        self.update_dashboard_period_label(payload.get("periodo", []))
        if self.dashboard_message is not None:
            message = "Resumen ejecutivo de costos, actividades, estructura y rentabilidad"
            if payload.get("modo_periodo") == "PROYECTADO":
                message = (
                    "Vista estimada/proyectada desde la estructura base del Excel maestro. "
                    "Los calculos originales no se modifican."
                )
            self.dashboard_message.configure(text=message)
        self.set_dashboard_kpis(payload)
        self.update_dashboard_insights(payload)
        if self.dashboard_chart_frame is None:
            return
        figure = self.build_dashboard_figure(payload)
        self.dashboard_chart_canvas = FigureCanvasTkAgg(figure, master=self.dashboard_chart_frame)
        self.dashboard_chart_canvas.draw()
        chart_widget = self.dashboard_chart_canvas.get_tk_widget()
        chart_widget.pack(fill=tk.BOTH, expand=True)
        self.bind_dashboard_scroll(chart_widget)
        self.update_idletasks()
        self.refresh_dashboard_scrollregion()

    def active_dashboard_payload(self) -> dict:
        return self.dashboard_payload_for_active_period(dashboard_payload(self.db_path))

    def dashboard_payload_for_active_period(self, payload: dict) -> dict:
        if not payload:
            return payload
        if not self.active_period or self.active_period.get("origen_periodo") != "MANUAL":
            return payload
        factor = self.active_period_projection_factor()
        projected = {
            key: [dict(row) for row in rows] if isinstance(rows, list) else rows
            for key, rows in payload.items()
        }
        projected["periodo"] = [dict(self.active_period)]
        projected["modo_periodo"] = "PROYECTADO"
        projected["factor_proyeccion"] = factor
        for row in projected.get("costos_ot", []):
            self.scale_fields(
                row,
                [
                    "cantidad_producida",
                    "cmd_ot",
                    "mod_ot",
                    "otros_recursos_ot",
                    "costo_produccion_ot",
                    "costo_soporte_ot",
                    "csd_ot",
                    "costo_total_ot",
                ],
                factor,
            )
        for row in projected.get("costos_actividad", []):
            self.scale_fields(
                row,
                ["costo_mod_ac", "costo_otros_recursos_ac", "costo_conversion_ac", "costo_soporte_ac", "costo_total_ac"],
                factor,
            )
        for row in projected.get("estado_resultados", []):
            self.scale_fields(
                row,
                [
                    "ventas_netas_soles",
                    "materia_prima_directa",
                    "costos_conversion",
                    "margen_bruto",
                    "procesos_soporte",
                    "margen_operativo",
                ],
                factor,
            )
        for row in projected.get("alertas", []):
            self.scale_fields(row, ["valor_referencia"], factor)
            observation = row.get("observacion") or ""
            projection_note = f"Periodo manual proyectado con factor {factor:g}; calculos originales sin modificar."
            row["observacion"] = f"{observation} {projection_note}".strip()
        return projected

    def active_period_projection_factor(self) -> float:
        if not self.active_period:
            return 1.0
        type_factor = {
            "mensual": 1.0,
            "bimestral": 2.0,
            "trimestral": 3.0,
            "cuatrimestral": 4.0,
            "semestral": 6.0,
            "anual": 12.0,
        }.get(str(self.active_period.get("tipo_periodo") or "").lower())
        if type_factor is not None:
            return type_factor
        try:
            start = datetime.strptime(str(self.active_period.get("fecha_inicio")), "%Y-%m-%d").date()
            end = datetime.strptime(str(self.active_period.get("fecha_fin")), "%Y-%m-%d").date()
        except (TypeError, ValueError):
            return 1.0
        duration = max((end - start).days + 1, 1)
        return duration / 31

    def scale_fields(self, row: dict, fields: list[str], factor: float) -> None:
        for field in fields:
            if field in row and row[field] is not None:
                row[field] = self.as_number(row, field) * factor

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
        lines = self.build_dashboard_insights_lines(payload)
        self.dashboard_insights_text.configure(state=tk.NORMAL)
        self.dashboard_insights_text.delete("1.0", tk.END)
        self.dashboard_insights_text.insert("1.0", "\n".join(lines))
        self.dashboard_insights_text.configure(height=min(max(len(lines), 6), 14))
        self.dashboard_insights_text.configure(state=tk.DISABLED)
        self.refresh_dashboard_scrollregion()

    def build_dashboard_insights_lines(self, payload: dict) -> list[str]:
        if not payload:
            return ["Sin datos calculados para alertas y recomendaciones."]
        costs = payload.get("costos_ot", [])
        activities = payload.get("costos_actividad", [])
        estado = payload.get("estado_resultados", [])
        alerts = payload.get("alertas", [])
        lines = ["ALERTAS"]
        if payload.get("modo_periodo") == "PROYECTADO":
            factor = payload.get("factor_proyeccion", 1)
            lines.append(
                f"- Vista estimada/proyectada para periodo manual. Factor aplicado: {factor:g}. Los calculos originales no se modifican."
            )
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

    def generate_improvement_proposal(self) -> None:
        if self.proposal_text is None:
            return
        payload = self.active_dashboard_payload()
        text = self.build_improvement_proposal_text(payload)
        self.proposal_text.delete("1.0", tk.END)
        self.proposal_text.insert("1.0", text)
        self.proposal_text.edit_reset()
        self.status_text.set("Propuesta de mejoras generada para revision y edicion.")

    def save_improvement_proposal(self) -> None:
        if self.proposal_text is None:
            return
        REPORTS_DIR.mkdir(parents=True, exist_ok=True)
        path = filedialog.asksaveasfilename(
            title="Guardar propuesta de mejoras",
            initialdir=str(REPORTS_DIR),
            initialfile="Propuesta_mejoras_editable.txt",
            defaultextension=".txt",
            filetypes=[("Texto", "*.txt"), ("Todos", "*.*")],
        )
        if not path:
            return
        Path(path).write_text(self.proposal_text.get("1.0", tk.END).strip() + "\n", encoding="utf-8")
        messagebox.showinfo("Propuesta de mejoras", f"Propuesta guardada en:\n{path}")
        self.status_text.set(f"Propuesta de mejoras guardada: {path}")

    def build_improvement_proposal_text(self, payload: dict) -> str:
        if not payload:
            return (
                "INFORME DE PROPUESTAS DE MEJORA\n"
                "Rol: Gerencia de Contabilidad de Costos\n\n"
                "No hay una ejecucion de calculo disponible. Ejecute el Modelo ABC por Ordenes de Trabajo "
                "para generar una propuesta sustentada en resultados."
            )
        costs = payload.get("costos_ot", [])
        activities = payload.get("costos_actividad", [])
        estado = payload.get("estado_resultados", [])
        alerts = payload.get("alertas", [])
        period_text = self.format_dashboard_period(payload.get("periodo", [])) or "Periodo no especificado"
        total_cost = sum(self.as_number(row, "costo_total_ot") for row in costs)
        total_sales = sum(self.as_number(row, "ventas_netas_soles") for row in estado)
        avg_margin = (
            sum(self.as_number(row, "rentabilidad_operativa") for row in estado) / len(estado)
            if estado
            else 0.0
        )
        top_activity = max(activities, key=lambda row: self.as_number(row, "costo_total_ac")) if activities else {}
        lowest_margin = min(estado, key=lambda row: self.as_number(row, "rentabilidad_operativa")) if estado else {}
        highest_cost = max(costs, key=lambda row: self.as_number(row, "costo_total_ot")) if costs else {}
        recommendations = self.dashboard_recommendations(costs, activities, estado)

        lines = [
            "INFORME DE PROPUESTAS DE MEJORA",
            "Rol: Gerencia de Contabilidad de Costos",
            "",
            "1. Alcance ejecutivo",
            "Este informe editable presenta acciones de mejora orientadas a elevar la rentabilidad de la empresa evaluada, usando como base los resultados del sistema ABC por ordenes de trabajo.",
            f"Periodo evaluado: {period_text}",
        ]
        if payload.get("modo_periodo") == "PROYECTADO":
            lines.append(
                f"Nota de periodo nuevo: informacion estimada/proyectada con factor {payload.get('factor_proyeccion', 1):g}, "
                "tomando como referencia la estructura del Excel maestro importado sin alterar los calculos originales."
            )
        lines.extend(
            [
                "",
                "2. Diagnostico de costos y rentabilidad",
                f"Costo total acumulado de OT: S/ {total_cost:,.2f}",
                f"Ventas netas registradas: S/ {total_sales:,.2f}",
                f"Rentabilidad operativa promedio: {avg_margin * 100:,.2f}%",
            ]
        )
        if highest_cost:
            lines.append(
                f"OT/producto con mayor costo total: {highest_cost.get('codigo_producto') or highest_cost.get('codigo_ot') or highest_cost.get('ot_id')} "
                f"(S/ {self.as_number(highest_cost, 'costo_total_ot'):,.2f})."
            )
        if top_activity:
            lines.append(
                f"Actividad con mayor costo: {top_activity.get('nombre_actividad') or top_activity.get('codigo_actividad')} "
                f"(S/ {self.as_number(top_activity, 'costo_total_ac'):,.2f})."
            )
        if lowest_margin:
            lines.append(
                f"OT/producto con menor rentabilidad: {lowest_margin.get('codigo_producto') or lowest_margin.get('ot_id')} "
                f"({self.as_number(lowest_margin, 'rentabilidad_operativa') * 100:,.2f}%)."
            )

        lines.extend(["", "3. Alertas de gestion"])
        if alerts:
            for alert in alerts:
                value = self.as_number(alert, "valor_referencia")
                value_text = f" Valor de referencia: {self.format_number(value)}." if value else ""
                lines.append(f"- {alert.get('severidad')}: {alert.get('mensaje')}.{value_text}")
        else:
            lines.append("- No se registran alertas calculadas para la ejecucion actual.")

        lines.extend(["", "4. Propuestas de mejora con impacto positivo en rentabilidad"])
        if recommendations:
            for index, recommendation in enumerate(recommendations, start=1):
                lines.append(f"{index}. {recommendation}")
        else:
            lines.append("1. Completar el ciclo de importacion y calculo para sustentar propuestas cuantitativas.")

        lines.extend(
            [
                "",
                "5. Plan de accion sugerido",
                "1. Validar con responsables operativos las actividades de mayor costo y sus inductores.",
                "2. Priorizar mejoras en actividades o recursos con mayor participacion en el costo total.",
                "3. Revisar precios, mezcla comercial y condiciones productivas de las OT con menor margen.",
                "4. Monitorear el impacto mensual con el mismo modelo ABC para comparar costo, margen y rentabilidad.",
                "",
                "6. Espacio para ajustes del usuario",
                "Edite, agregue o elimine propuestas segun el criterio gerencial, restricciones operativas y objetivos de la empresa.",
            ]
        )
        return "\n".join(lines)

    def build_dashboard_figure(self, payload: dict) -> Figure:
        costs = payload.get("costos_ot", [])
        activities = payload.get("costos_actividad", [])
        estado = payload.get("estado_resultados", [])
        figure = Figure(figsize=(13.4, 6.7), dpi=100, facecolor="#FFFFFF", constrained_layout=True)
        figure.set_constrained_layout_pads(w_pad=0.08, h_pad=0.14, wspace=0.08, hspace=0.12)
        axes = figure.subplots(2, 2, gridspec_kw={"height_ratios": [1, 1.08]})
        ax_cost, ax_mix, ax_activity, ax_margin = axes.flatten()

        self.style_dashboard_axis(ax_cost, "Costo total por OT")
        labels = [self.dashboard_label(row.get("codigo_producto") or row.get("codigo_ot") or row.get("ot_id")) for row in costs]
        total_costs = [self.as_number(row, "costo_total_ot") for row in costs]
        if total_costs:
            bars = ax_cost.bar(labels, total_costs, color="#1D4ED8", width=0.62)
            self.pad_axis_ylim(ax_cost, total_costs)
            self.annotate_bars(ax_cost, bars, currency=True)
            ax_cost.tick_params(axis="x", rotation=0)
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
            self.pad_axis_ylim(ax_mix, bottoms)
            ax_mix.set_xticks(list(x_positions), labels, rotation=0)
            ax_mix.set_ylabel("Soles")
            ax_mix.legend(loc="upper left", fontsize=8, ncols=4, frameon=False)
        else:
            self.empty_axis(ax_mix)

        self.style_dashboard_axis(ax_activity, "Ranking de actividades")
        top_activities = sorted(activities, key=lambda row: self.as_number(row, "costo_total_ac"), reverse=True)[:7]
        if top_activities:
            names = [self.dashboard_label(row.get("nombre_actividad") or row.get("codigo_actividad"), max_chars=28) for row in top_activities]
            values = [self.as_number(row, "costo_total_ac") for row in top_activities]
            bars = ax_activity.barh(names, values, color="#059669")
            ax_activity.invert_yaxis()
            self.annotate_horizontal_bars(ax_activity, bars)
            ax_activity.set_xlabel("Soles")
        else:
            self.empty_axis(ax_activity)

        self.style_dashboard_axis(ax_margin, "Rentabilidad operativa por OT")
        margin_labels = [self.dashboard_label(row.get("codigo_producto") or row.get("ot_id")) for row in estado]
        margins = [self.as_number(row, "rentabilidad_operativa") * 100 for row in estado]
        if margins:
            colors = ["#DC2626" if value < 0 else "#0891B2" for value in margins]
            bars = ax_margin.bar(margin_labels, margins, color=colors)
            self.pad_axis_ylim(ax_margin, margins, include_zero=True)
            self.annotate_bars(ax_margin, bars, suffix="%")
            ax_margin.axhline(0, color="#334155", linewidth=0.8)
            ax_margin.tick_params(axis="x", rotation=0)
            ax_margin.set_ylabel("%")
        else:
            self.empty_axis(ax_margin)
        return figure

    def style_dashboard_axis(self, axis, title: str) -> None:
        axis.set_title(title, fontsize=11, fontweight="bold", color="#0F172A", loc="left", pad=8)
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

    def dashboard_label(self, value, max_chars: int = 20) -> str:
        text = str(value or "")
        if len(text) <= max_chars:
            return text
        midpoint = max_chars // 2
        return text[:midpoint].rstrip() + "\n" + text[midpoint:max_chars].rstrip() + "..."

    def pad_axis_ylim(self, axis, values: list[float], include_zero: bool = True) -> None:
        if not values:
            return
        minimum = min(values)
        maximum = max(values)
        if include_zero:
            minimum = min(minimum, 0)
            maximum = max(maximum, 0)
        if minimum == maximum:
            padding = abs(maximum) * 0.2 or 1
        else:
            padding = (maximum - minimum) * 0.18
        axis.set_ylim(minimum - padding, maximum + padding)

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
            payload = self.active_dashboard_payload()
            paths = generate_all_reports(self.db_path, REPORTS_DIR, payload or None)
            self.refresh_outputs()
            messagebox.showinfo("Reportes", "Reportes generados:\n" + "\n".join(str(path) for path in paths))
        except Exception as exc:
            messagebox.showerror("Reportes", str(exc))

    def generate_documents(self) -> None:
        try:
            payload = self.active_dashboard_payload()
            paths = generate_all_documents(self.db_path, REPORTS_DIR, dashboard_payload_override=payload or None)
            self.refresh_outputs()
            messagebox.showinfo("Documentos", "Documentos generados:\n" + "\n".join(str(path) for path in paths))
        except Exception as exc:
            messagebox.showerror("Documentos", str(exc))

    def generate_dashboard_png(self) -> None:
        path = generate_dashboard_image(self.db_path, REPORTS_DIR, self.active_dashboard_payload() or None)
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
