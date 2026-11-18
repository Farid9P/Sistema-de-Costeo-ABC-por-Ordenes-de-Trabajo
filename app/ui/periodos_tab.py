"""Pestaña "Evolucion y comparativo": mes a mes del periodo analizado, comparativo entre
subperiodos del año (con los meses cargados) y kardex consolidado de materiales."""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk

import warnings

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

from .. import analisis_periodos as ap
from .pagina import agregar_pagina

TITULO_PESTANA = "Evolucion"
warnings.filterwarnings("ignore", message="constrained_layout not applied")
DECIMALES_4 = {"costo_unitario_ot", "costo_unitario_medio", "peso_ac_ot", "saldo_final_cto_unit", "rentabilidad_operativa"}
ENTEROS = {"anio", "num_mes", "meses", "cantidad_producida", "unidades"}


def formatear(filas: list[dict]) -> list[dict]:
    salida = []
    for fila in filas:
        nueva = {}
        for clave, valor in fila.items():
            if clave.startswith("_"):
                continue
            if isinstance(valor, float):
                valor = f"{valor:,.0f}" if clave in ENTEROS else f"{valor:,.4f}" if clave in DECIMALES_4 else f"{valor:,.2f}"
            nueva[clave] = valor
        salida.append(nueva)
    return salida


class PestanaEvolucion:
    def __init__(self, app, notebook: ttk.Notebook):
        self.app = app
        self._lienzo = None
        pagina, tab = agregar_pagina(notebook, TITULO_PESTANA, min_alto=1650)
        app.module_tabs.append((pagina, TITULO_PESTANA))
        self.tab = pagina
        self.titulo = ttk.Label(tab, text="Evolucion y comparativo", style="Subtitle.TLabel", font=("Segoe UI", 12, "bold"))
        self.titulo.pack(anchor=tk.W)
        self.info = ttk.Label(tab, text="", style="Subtitle.TLabel", wraplength=1200, justify=tk.LEFT)
        self.info.pack(anchor=tk.W, fill=tk.X, pady=(2, 8))
        self.info.bind("<Configure>", lambda e: self.info.configure(wraplength=max(400, e.width - 10)))
        self.graficos = tk.Frame(tab, bg="#FFFFFF", highlightthickness=1, highlightbackground="#CBD5E1", height=430)
        self.graficos.pack(fill=tk.X)
        self.graficos.pack_propagate(False)
        for clave, titulo, alto in (("evo_mensual", "Mes a mes del periodo analizado", 330),
                                    ("evo_comparativo", "Comparativo por subperiodo del año (meses cargados)", 260),
                                    ("evo_kardex", "Kardex de materiales directos del periodo", 220)):
            caja = ttk.LabelFrame(tab, text=titulo, padding=6, style="Section.TLabelframe", height=alto)
            caja.pack(fill=tk.X, pady=(10, 0))
            caja.pack_propagate(False)
            app.trees[clave] = app._make_tree(caja)

    def refrescar(self) -> None:
        app = self.app
        payload = app.analisis_payload or {}
        if not payload.get("meses_con_datos"):
            for clave in ("evo_mensual", "evo_comparativo", "evo_kardex"):
                app.populate_tree(clave, [])
            self.info.configure(text="Ejecute el calculo de un periodo para ver su evolucion.")
            return
        meses = payload["meses_con_datos"]
        anio = meses[-1][0]
        tipo = app.analisis["tipo"] if app.analisis else "Mensual"
        tipo_comp = tipo if ap.TIPOS_ANALISIS.get(tipo, 12) < 12 else "Trimestral"
        if tipo_comp == "Mensual":
            tipo_comp = "Trimestral"
        meses_anio = [(p["anio"], p["mes"]) for p in app.meses_cargados if p["anio"] == anio]
        actuales = app.resultados_actuales or {}
        resultados = ap.resultados_meses(app.db_path, [m for m in meses_anio if m not in actuales])
        resultados.update(actuales)
        self.comparativo = ap.comparar_bloques(app.db_path, tipo_comp, anio, resultados=resultados)
        self.titulo.configure(text=f"Evolucion y comparativo — {payload['titulo']}")
        self.info.configure(text=(
            f"Meses del periodo: {payload['texto_meses']}. Comparativo {tipo_comp.lower()} {anio} con los meses cargados "
            "(los subperiodos sin todos sus meses figuran como 'Incompleto'). Importes = suma de los meses; costo unitario "
            "y rentabilidad = sobre los totales."))
        app.populate_tree("evo_mensual", formatear(payload.get("mensual", [])))
        app.populate_tree("evo_comparativo", formatear(self.comparativo))
        app.populate_tree("evo_kardex", formatear(payload.get("kardex", [])))
        self._dibujar(payload)

    def _dibujar(self, payload: dict) -> None:
        if self._lienzo is not None:
            self._lienzo.get_tk_widget().destroy()
        app = self.app
        mensual = payload.get("mensual", [])
        figura = Figure(figsize=(13.4, 4.2), dpi=100, facecolor="#FFFFFF", constrained_layout=True)
        ax1, ax2, ax3 = figura.subplots(1, 3)
        etiquetas = [ap.nombre_mes(f["anio"], f["num_mes"], corto=True) for f in mensual]
        x = list(range(len(mensual)))
        app.style_dashboard_axis(ax1, "Ventas, costo total y margen operativo")
        ancho = 0.27
        ax1.bar([i - ancho for i in x], [f["ventas_netas_soles"] for f in mensual], ancho, color="#D97706", label="Ventas")
        ax1.bar(x, [f["costo_total_ot"] for f in mensual], ancho, color="#1D4ED8", label="Costo total")
        ax1.bar([i + ancho for i in x], [f["margen_operativo"] for f in mensual], ancho,
                color=["#DC2626" if f["margen_operativo"] < 0 else "#059669" for f in mensual], label="Margen operativo")
        ax1.axhline(0, color="#334155", linewidth=0.8)
        todos = [f[c] for f in mensual for c in ("ventas_netas_soles", "costo_total_ot", "margen_operativo")]
        if todos:
            minimo, maximo = min(min(todos), 0), max(todos)
            ax1.set_ylim(minimo - (maximo - minimo) * 0.05, maximo + (maximo - minimo) * 0.25)
        ax1.set_xticks(x, etiquetas, rotation=45 if len(x) > 6 else 0)
        ax1.legend(fontsize=8, frameon=False, loc="upper left", ncols=3)
        app.style_dashboard_axis(ax2, "Rentabilidad operativa por mes")
        valores = [f["rentabilidad_operativa"] * 100 for f in mensual]
        self._barras(ax2, etiquetas, valores, "#0891B2")
        app.style_dashboard_axis(ax3, "Rentabilidad por subperiodo")
        nombres = [b["subperiodo"].split(":")[0].replace("Trimestre ", "T").replace("Bimestre ", "B")
                   .replace("Cuatrimestre ", "C").replace("Semestre ", "S") for b in self.comparativo]
        nombres = [n if len(n) <= 6 else n[:3] for n in nombres]
        vals = [b["rentabilidad_operativa"] * 100 if b["estado"] != "Sin datos" else 0 for b in self.comparativo]
        self._barras(ax3, nombres, vals, "#7C3AED", vacios=[b["estado"] == "Sin datos" for b in self.comparativo])
        self._lienzo = FigureCanvasTkAgg(figura, master=self.graficos)
        self._lienzo.draw()
        self._lienzo.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def _barras(self, eje, etiquetas, valores, color, vacios=None) -> None:
        vacios = vacios or [False] * len(valores)
        colores = ["#CBD5E1" if v else ("#DC2626" if x < 0 else color) for x, v in zip(valores, vacios)]
        barras = eje.bar(etiquetas, valores, color=colores)
        self.app.pad_axis_ylim(eje, valores)
        eje.axhline(0, color="#334155", linewidth=0.8)
        eje.set_ylabel("%")
        eje.tick_params(axis="x", rotation=45 if len(etiquetas) > 6 else 0)
        for barra, valor, vacio in zip(barras, valores, vacios):
            if vacio:
                continue
            eje.annotate(f"{valor:,.1f}%", xy=(barra.get_x() + barra.get_width() / 2, valor),
                         xytext=(0, 4 if valor >= 0 else -4), textcoords="offset points", ha="center",
                         va="bottom" if valor >= 0 else "top", fontsize=8, color="#334155")
