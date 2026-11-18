"""Pantalla inicial: carga de uno o varios Excel maestros (un mes cada uno), eleccion del periodo de
analisis (mensual, bimestral, trimestral, cuatrimestral, semestral, anual) y ejecucion del calculo."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .. import analisis_periodos as ap
from ..database import db_session
from ..importers import validate_workbook
from ..model_availability import evaluate_models


class FlujoPeriodosMixin:
    # ------------------------------------------------------------------ estado
    def _init_flujo(self) -> None:
        self.archivos_excel: list[str] = []
        self.archivos_periodo: dict[str, dict] = {}
        self.meses_cargados: list[dict] = []
        self.tipo_analisis_var = tk.StringVar(value="")
        self.periodo_analisis_var = tk.StringVar(value="")
        self.meses_info_var = tk.StringVar(value="Aun no hay archivos cargados.")
        self.archivos_info_var = tk.StringVar(value="")
        self.mes_visible_var = tk.StringVar(value="")
        self.analisis: dict | None = None  # {"tipo", "titulo", "meses"}
        self.analisis_payload: dict | None = None
        self.payload_inicial: dict | None = None
        self.resultados_actuales: dict = {}
        self.mes_visible: tuple[int, int] | None = None
        self._ventanas: dict = {}
        self.calculo_ejecutado = False

    def es_multimes(self) -> bool:
        return bool(self.analisis and len(self.analisis["meses"]) > 1)

    # ------------------------------------------------------------------ construccion
    def _build_import_box(self, parent) -> None:
        caja = ttk.LabelFrame(parent, text="Importar Excel maestro (uno o varios meses)", padding=12, style="Section.TLabelframe")
        caja.pack(fill=tk.X)
        ttk.Label(caja, text="Archivos Excel maestros:").grid(row=0, column=0, sticky=tk.W)
        entrada = ttk.Entry(caja, textvariable=self.excel_path, state="readonly")
        entrada.grid(row=0, column=1, sticky=tk.EW, padx=8)
        caja.columnconfigure(1, weight=1)
        ttk.Button(caja, text="Seleccionar", command=self.choose_excel).grid(row=0, column=2, padx=4)
        ttk.Button(caja, text="Validar estructura", command=self.validate_excel).grid(row=0, column=3, padx=4)
        ttk.Button(caja, text="Importar a SQLite", command=self.import_excel).grid(row=0, column=4, padx=4)
        info = ttk.Label(caja, textvariable=self.archivos_info_var, style="Subtitle.TLabel", justify=tk.LEFT, wraplength=1100)
        info.grid(row=1, column=0, columnspan=5, sticky=tk.W, pady=(8, 0))
        info.bind("<Configure>", lambda e: info.configure(wraplength=max(400, e.width - 10)))
        ttk.Label(
            caja,
            text="Puede seleccionar varios archivos a la vez (Ctrl o Shift + clic). Cada archivo es un mes: el sistema lee su "
                 "hoja Periodo, lo valida, lo importa y calcula sus costos.",
            style="Subtitle.TLabel",
        ).grid(row=2, column=0, columnspan=5, sticky=tk.W, pady=(4, 0))

    def _build_period_box(self, parent) -> None:
        self.period_box = ttk.LabelFrame(parent, text="Seleccion del periodo de analisis", padding=12, style="Section.TLabelframe")
        self.period_box.pack(fill=tk.X, pady=(12, 0))
        caja = self.period_box
        ttk.Label(caja, text="Meses cargados:").grid(row=0, column=0, sticky=tk.W)
        meses = ttk.Label(caja, textvariable=self.meses_info_var, style="Subtitle.TLabel", justify=tk.LEFT, wraplength=900)
        meses.grid(row=0, column=1, columnspan=3, sticky=tk.W, padx=8)
        self.quitar_button = ttk.Button(caja, text="Quitar archivos cargados", command=self.quitar_archivos, state=tk.DISABLED)
        self.quitar_button.grid(row=0, column=4, sticky=tk.E)
        ttk.Label(caja, text="Tipo de analisis:").grid(row=1, column=0, sticky=tk.W, pady=(10, 0))
        self.tipo_combo = ttk.Combobox(caja, textvariable=self.tipo_analisis_var, state="disabled", width=20)
        self.tipo_combo.grid(row=1, column=1, sticky=tk.W, padx=8, pady=(10, 0))
        self.tipo_combo.bind("<<ComboboxSelected>>", lambda _e: self.actualizar_opciones_periodo())
        ttk.Label(caja, text="Periodo a analizar:").grid(row=1, column=2, sticky=tk.W, pady=(10, 0))
        self.period_combo = ttk.Combobox(caja, textvariable=self.periodo_analisis_var, state="disabled", width=60)
        self.period_combo.grid(row=1, column=3, sticky=tk.EW, padx=8, pady=(10, 0))
        self.period_combo.bind("<<ComboboxSelected>>", lambda _e: self.seleccionar_analisis())
        self.period_status = ttk.Label(caja, text="Importe uno o varios Excel maestros mensuales para elegir el periodo.",
                                       style="Subtitle.TLabel")
        self.period_status.grid(row=2, column=0, columnspan=5, sticky=tk.W, pady=(10, 0))
        caja.columnconfigure(3, weight=1)

    # ------------------------------------------------------------------ archivos
    def choose_excel(self) -> None:
        rutas = filedialog.askopenfilenames(title="Seleccionar Excel maestros (uno por mes)",
                                            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")])
        if not rutas:
            return
        self.archivos_excel = list(rutas)
        self.archivos_periodo = {}
        self.is_excel_validated = False
        lineas = []
        for ruta in self.archivos_excel:
            try:
                periodo = ap.leer_periodo_excel(ruta)
                self.archivos_periodo[ruta] = periodo
                lineas.append(f"{Path(ruta).name} → {ap.nombre_mes(periodo['anio'], periodo['mes'])} ({periodo.get('periodo_id')})")
            except Exception as exc:
                lineas.append(f"{Path(ruta).name} → no se identifico el periodo: {exc}")
        nombres = ", ".join(Path(r).name for r in self.archivos_excel)
        self.excel_path.set(f"{len(self.archivos_excel)} archivo(s): {nombres}")
        self.archivos_info_var.set("Seleccionados: " + " | ".join(lineas))
        self.status_text.set("Excel seleccionados. Valide la estructura antes de importar.")

    def validate_excel(self) -> None:
        if not self.archivos_excel:
            messagebox.showwarning("Validacion", "Primero seleccione uno o varios Excel maestros.")
            return
        lineas, errores, meses = [], 0, {}
        for ruta in self.archivos_excel:
            nombre = Path(ruta).name
            try:
                periodo = ap.leer_periodo_excel(ruta)
                mes = (periodo["anio"], periodo["mes"])
                resultado = validate_workbook(ruta)
            except Exception as exc:
                errores += 1
                lineas.append(f"✗ {nombre}: {exc}")
                continue
            if mes in meses:
                errores += 1
                lineas.append(f"✗ {nombre}: {ap.nombre_mes(*mes)} ya esta en {Path(meses[mes]).name}; elija un archivo por mes.")
                continue
            meses[mes] = ruta
            if resultado.ok:
                lineas.append(f"✓ {nombre}: {ap.nombre_mes(*mes)} ({periodo.get('periodo_id')}) — estructura valida"
                              + (f", {len(resultado.warnings)} advertencia(s)" if resultado.warnings else ""))
            else:
                errores += 1
                lineas.append(f"✗ {nombre}: " + "; ".join(i.message for i in resultado.errors[:3]))
        self.is_excel_validated = errores == 0
        self.validated_excel_path = self.excel_path.get() if self.is_excel_validated else None
        (messagebox.showinfo if errores == 0 else messagebox.showerror)("Validacion", "\n".join(lineas[:30]))
        self.status_text.set(f"Validacion: {len(self.archivos_excel) - errores} archivo(s) validos, {errores} con errores.")

    def import_excel(self) -> None:
        if not self.archivos_excel:
            messagebox.showwarning("Importacion", "Primero seleccione uno o varios Excel maestros.")
            return
        if not self.is_excel_validated or self.validated_excel_path != self.excel_path.get():
            messagebox.showwarning("Importacion", "Primero valide la estructura de los Excel seleccionados. Luego importelos.")
            return

        def progreso(i, n, nombre):
            self.status_text.set(f"Importando y calculando {i}/{n}: {nombre}...")
            self.update_idletasks()

        self.configure(cursor="watch")
        try:
            hechos, errores = ap.procesar_varios(self.archivos_excel, self.db_path, progreso)
        finally:
            self.configure(cursor="")
        lineas = [f"Importados {len(hechos)} de {len(self.archivos_excel)} archivo(s):"]
        lineas += [f"- {h.etiqueta} ({h.periodo_id}) desde {h.archivo}" + (" — reemplazo el mes ya cargado" if h.reemplazado else "")
                    for h in hechos]
        if errores:
            lineas += ["", "No se importaron:"] + [f"- {n}: {m}" for n, m in errores]
        (messagebox.showwarning if errores else messagebox.showinfo)("Importacion", "\n".join(lineas[:40]))
        self.is_excel_imported = bool(hechos)
        self.imported_excel_path = self.excel_path.get()
        self.status_text.set(lineas[0])
        self.recargar_meses_cargados()

    def quitar_archivos(self) -> None:
        if not self.meses_cargados:
            return
        if not messagebox.askyesno("Archivos cargados", "¿Quitar todos los meses cargados? Los Excel no se borran; "
                                                        "puede volver a importarlos."):
            return
        ap.iniciar_sesion_vacia(self.db_path)
        self.analisis = None
        self.analisis_payload = None
        self.calculo_ejecutado = False
        self.recargar_meses_cargados()
        self.status_text.set("Se quitaron los meses cargados.")

    # ------------------------------------------------------------------ periodo de analisis
    def recargar_meses_cargados(self) -> None:
        self.meses_cargados = ap.periodos_procesados(self.db_path)
        meses = [(p["anio"], p["mes"]) for p in self.meses_cargados]
        self._ventanas = ap.ventanas(meses)
        if self.meses_cargados:
            self.meses_info_var.set(f"{len(meses)} mes(es): " + ", ".join(
                f"{p['etiqueta']} ({p['archivo']})" for p in self.meses_cargados))
            self.quitar_button.configure(state=tk.NORMAL)
        else:
            self.meses_info_var.set("Aun no hay archivos cargados.")
            self.quitar_button.configure(state=tk.DISABLED)
        tipos = list(self._ventanas)
        self.tipo_combo.configure(values=tipos, state="readonly" if tipos else "disabled")
        if self.tipo_analisis_var.get() not in tipos:
            self.tipo_analisis_var.set(tipos[0] if tipos else "")
        self.analisis = None
        self.hide_model_selection()
        self.reset_model_selector()
        self.actualizar_opciones_periodo()

    def actualizar_opciones_periodo(self) -> None:
        opciones = self._ventanas.get(self.tipo_analisis_var.get(), [])
        etiquetas = [self._etiqueta_opcion(e, m) for e, m in opciones]
        self.period_combo.configure(values=etiquetas, state="readonly" if etiquetas else "disabled")
        self.periodo_analisis_var.set("")
        self.analisis = None
        self.hide_model_selection()
        if not self.meses_cargados:
            self.period_status.configure(text="Importe uno o varios Excel maestros mensuales para elegir el periodo.")
        elif etiquetas:
            disponibles = ", ".join(self._ventanas)
            self.period_status.configure(
                text=f"Tipos disponibles segun los meses consecutivos cargados: {disponibles}. Elija el periodo a analizar.")
        if len(etiquetas) == 1:
            self.periodo_analisis_var.set(etiquetas[0])
            self.seleccionar_analisis()

    def _etiqueta_opcion(self, etiqueta: str, meses: list[tuple[int, int]]) -> str:
        ids = [next((p["periodo_id"] for p in self.meses_cargados if (p["anio"], p["mes"]) == m), "") for m in meses]
        return f"{etiqueta} ({' + '.join(ids)})"

    def seleccionar_analisis(self) -> None:
        tipo = self.tipo_analisis_var.get()
        opciones = self._ventanas.get(tipo, [])
        etiquetas = [self._etiqueta_opcion(e, m) for e, m in opciones]
        if self.periodo_analisis_var.get() not in etiquetas:
            self.analisis = None
            self.hide_model_selection()
            return
        titulo, meses = opciones[etiquetas.index(self.periodo_analisis_var.get())]
        if tipo != "Mensual":
            titulo = f"{tipo}: {titulo}"
        self.analisis = {"tipo": tipo, "titulo": titulo, "meses": meses}
        self.calculo_ejecutado = False
        self.fijar_mes_visible(meses[0])
        self.is_excel_imported = True
        self.period_status.configure(text=f"Periodo de analisis: {titulo} — {len(meses)} mes(es). Elija el modelo y presione "
                                          "'Ejecutar calculo'.")
        self.refresh_model_statuses()
        self.status_text.set(f"Periodo de analisis seleccionado: {titulo}")

    def fijar_mes_visible(self, mes: tuple[int, int]) -> None:
        """Carga el mes en las tablas de trabajo (las que muestra 'Datos importados')."""
        self.mes_visible = mes
        ap.preparar_mes_en_base(self.db_path, *mes)
        fila = next((p for p in self.meses_cargados if (p["anio"], p["mes"]) == mes), None)
        self.active_period = None if fila is None else {
            "periodo_id": fila["periodo_id"], "periodo": fila["periodo"], "tipo_periodo": fila["tipo_periodo"],
            "fecha_inicio": fila["fecha_inicio"], "fecha_fin": fila["fecha_fin"], "origen_periodo": "IMPORTADO",
        }
        self.mes_visible_var.set(ap.nombre_mes(*mes))
        if hasattr(self, "mes_datos_combo") and self.analisis:
            self.mes_datos_combo.configure(values=[ap.nombre_mes(*m) for m in self.analisis["meses"]])
            if self.es_multimes():
                self.mes_datos_frame.pack(side=tk.LEFT, padx=(8, 0), before=self.refresh_table_button)
            else:
                self.mes_datos_frame.pack_forget()

    def cambiar_mes_visible(self) -> None:
        if not self.analisis:
            return
        nombres = {ap.nombre_mes(*m): m for m in self.analisis["meses"]}
        mes = nombres.get(self.mes_visible_var.get())
        if mes and mes != self.mes_visible:
            self.fijar_mes_visible(mes)
            self.load_selected_imported_table()

    def refresh_model_statuses(self) -> None:
        if not self.analisis:
            self.reset_model_selector()
            self.hide_model_selection()
            return
        self.model_statuses = [item for item in evaluate_models(self.db_path) if item.code != "MODELO_3"]
        values = [f"{item.code} | {item.label}: {item.status} - {item.visual_name}" for item in self.model_statuses]
        self.model_combo.configure(values=values, state="readonly")
        self.execute_button.configure(state=tk.NORMAL)
        if values:
            self.model_combo.current(0)
            self.model_var.set(values[0])
        self.show_model_selection()
        self.show_selected_model_reason()

    # ------------------------------------------------------------------ ejecucion
    def execute_calculation(self) -> None:
        if not self.analisis:
            messagebox.showwarning("Calculo ABC", "Primero importe los Excel y elija el periodo de analisis.")
            return
        item = self.selected_model()
        if not item.available:
            messagebox.showwarning(item.label, item.reason)
            self.status_text.set(item.reason)
            return
        try:
            self.ejecutar_analisis()
            self.show_modules_screen()
            messagebox.showinfo("Calculo ABC", f"Calculo completado: {self.analisis['titulo']}.")
        except Exception as exc:
            messagebox.showerror("Calculo ABC", str(exc))
            self.status_text.set(f"Error de calculo: {exc}")

    def ejecutar_analisis(self) -> None:
        """Calcula el periodo elegido con los ajustes de simulacion vigentes y refresca todas las vistas."""
        tipo, titulo, meses = self.analisis["tipo"], self.analisis["titulo"], self.analisis["meses"]
        self.configure(cursor="watch")
        self.update_idletasks()
        try:
            if len(meses) == 1:
                run_id = ap.calcular_mes_en_base(self.db_path, *meses[0])
                self.mes_visible = meses[0]
                with db_session(self.db_path) as conn:
                    resultados = {meses[0]: ap.volcar_resultados(conn)}
            else:
                resultados = ap.resultados_meses(self.db_path, meses)
            payload = ap.analizar(self.db_path, meses, tipo, titulo, resultados=resultados)
            if len(meses) > 1:
                visible = self.mes_visible if self.mes_visible in meses else meses[0]
                self.fijar_mes_visible(visible)
                run_id = ap.consolidar_en_base(self.db_path, payload, resultados)
            self.resultados_actuales = resultados
            self.analisis_payload = payload
            self.payload_inicial = ap.analizar(self.db_path, meses, tipo, titulo)
            self.calculo_ejecutado = True
        finally:
            self.configure(cursor="")
        self.cambios_pendientes = 0
        self.refresh_all_views()
        self.actualizar_banner_simulacion()
        n = self.contar_cambios()
        nota = f" Simulacion con {n} cambio(s)." if n else ""
        self.status_text.set(f"Calculo completado: {titulo} (ejecucion {run_id}).{nota}")
