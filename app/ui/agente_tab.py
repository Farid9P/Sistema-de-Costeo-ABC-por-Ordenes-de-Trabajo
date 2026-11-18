"""Pestaña "Agente IA Finanzas" de la ventana principal.

El agente es de solo lectura: esta pestaña no tiene botones que guarden,
importen, calculen ni generen archivos.
"""

from __future__ import annotations

import textwrap
import tkinter as tk
from tkinter import ttk

from ..agente_ia import AVISO_SOLO_LECTURA, AgenteFinanzas

TITULO_PESTANA = "Agente IA Finanzas"
EJEMPLO = "¿Cuál es el producto con más ventas?"
TITULOS = {"PREGUNTA INTERPRETADA", "RESPUESTA", "PEDIDO DE CAMBIO NO PERMITIDO", "NO LOGRÉ INTERPRETAR LA PREGUNTA",
           "ALERTAS DEL DASHBOARD"}


class PestanaAgente:
    def __init__(self, app, notebook: ttk.Notebook):
        self.app = app
        self._agente: AgenteFinanzas | None = None
        self.pregunta_var = tk.StringVar(value=EJEMPLO)

        from .pagina import agregar_pagina

        pagina, tab = agregar_pagina(notebook, "Agente IA", min_alto=640)
        app.module_tabs.append((pagina, "Agente IA"))
        self.tab = pagina
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(3, weight=1)

        ttk.Label(tab, text="Agente IA de Finanzas y Contabilidad", style="Subtitle.TLabel",
                  font=("Segoe UI", 12, "bold")).grid(row=0, column=0, sticky=tk.W)
        ttk.Label(tab, text=AVISO_SOLO_LECTURA, style="Subtitle.TLabel", wraplength=1100).grid(
            row=1, column=0, sticky=tk.W, pady=(2, 8))

        barra = ttk.Frame(tab)
        barra.grid(row=2, column=0, sticky=tk.EW)
        barra.columnconfigure(1, weight=1)
        ttk.Label(barra, text="Pregunta:").grid(row=0, column=0, sticky=tk.W, padx=(0, 6))
        entrada = ttk.Entry(barra, textvariable=self.pregunta_var, font=("Segoe UI", 11))
        entrada.grid(row=0, column=1, sticky=tk.EW)
        entrada.bind("<Return>", lambda _e: self.consultar())
        ttk.Button(barra, text="Interpretar y responder", command=self.consultar).grid(row=0, column=2, padx=(6, 0))
        ttk.Button(barra, text="Solo interpretar", command=lambda: self.consultar(responder=False)).grid(row=0, column=3, padx=(6, 0))
        ttk.Button(barra, text="Propuestas de mejora", command=self.propuestas).grid(row=0, column=4, padx=(6, 0))
        ttk.Button(barra, text="Limpiar", command=self.limpiar).grid(row=0, column=5, padx=(6, 0))

        panel = ttk.PanedWindow(tab, orient=tk.HORIZONTAL)
        panel.grid(row=3, column=0, sticky="nsew", pady=(8, 0))

        izquierda = ttk.LabelFrame(panel, text="Catálogo base (doble clic para preguntar)", padding=6,
                                   style="Section.TLabelframe")
        izquierda.rowconfigure(0, weight=1)
        izquierda.columnconfigure(0, weight=1)
        self.arbol = ttk.Treeview(izquierda, show="tree", selectmode="browse")
        scroll_arbol = ttk.Scrollbar(izquierda, orient=tk.VERTICAL, command=self.arbol.yview)
        self.arbol.configure(yscrollcommand=scroll_arbol.set)
        self.arbol.grid(row=0, column=0, sticky="nsew")
        scroll_arbol.grid(row=0, column=1, sticky="ns")
        self.arbol.bind("<Double-1>", self._usar_pregunta_del_catalogo)
        panel.add(izquierda, weight=1)

        derecha = ttk.LabelFrame(panel, text="Interpretación y respuesta", padding=6, style="Section.TLabelframe")
        derecha.rowconfigure(0, weight=1)
        derecha.columnconfigure(0, weight=1)
        self.salida = tk.Text(derecha, wrap=tk.NONE, bg="#FFFFFF", fg="#0F172A", font=("Consolas", 10),
                              state=tk.DISABLED, padx=8, pady=6)
        self.salida.tag_configure("titulo", font=("Segoe UI", 10, "bold"), foreground="#1F4D78")
        self.salida.tag_configure("interpretada", font=("Segoe UI", 11, "bold"), foreground="#0F766E")
        sy = ttk.Scrollbar(derecha, orient=tk.VERTICAL, command=self.salida.yview)
        sx = ttk.Scrollbar(derecha, orient=tk.HORIZONTAL, command=self.salida.xview)
        self.salida.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.salida.grid(row=0, column=0, sticky="nsew")
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")
        panel.add(derecha, weight=3)

        self._cargar_catalogo()
        self._mostrar("Escriba una pregunta en lenguaje natural, por ejemplo:\n"
                      "  - ¿Cuál es el producto con más ventas?\n"
                      "  - ¿Qué producto generó más costos?\n"
                      "  - ¿Cuánto cuesta cada pieza de SW50?\n"
                      "  - Muéstrame el kardex de materiales\n\n"
                      "El agente la reformula con el vocabulario del Modelo ABC por Órdenes de Trabajo, "
                      "la relaciona con el catálogo base y responde con los datos ya calculados.")

    # ------------------------------------------------------------------
    @property
    def agente(self) -> AgenteFinanzas:
        if self._agente is None or str(self._agente.db_path) != str(self.app.db_path):
            self._agente = AgenteFinanzas(self.app.db_path)
        return self._agente

    def _cargar_catalogo(self) -> None:
        try:
            agente = self.agente
        except Exception as exc:
            self.arbol.insert("", tk.END, text=f"No se pudo leer el catálogo: {exc}")
            return
        nodos: dict[str, str] = {}
        for pregunta in agente.catalogo:
            padre = nodos.get(pregunta.categoria)
            if padre is None:
                padre = self.arbol.insert("", tk.END, text=pregunta.categoria, open=False)
                nodos[pregunta.categoria] = padre
            self.arbol.insert(padre, tk.END, text=f"{pregunta.codigo}. {pregunta.texto}", values=(pregunta.texto,))

    def _usar_pregunta_del_catalogo(self, _event=None) -> None:
        item = self.arbol.focus()
        valores = self.arbol.item(item, "values")
        if not valores:
            return
        self.pregunta_var.set(valores[0].replace("{", "").replace("}", ""))
        self.consultar()

    def _mostrar(self, texto: str) -> None:
        self.salida.configure(state=tk.NORMAL)
        self.salida.delete("1.0", tk.END)
        siguiente_es_interpretada = False
        for linea in texto.splitlines():
            if linea.strip() in TITULOS:
                self.salida.insert(tk.END, linea + "\n", "titulo")
                siguiente_es_interpretada = linea.strip() == "PREGUNTA INTERPRETADA"
            elif siguiente_es_interpretada:
                self.salida.insert(tk.END, linea + "\n", "interpretada")
                siguiente_es_interpretada = False
            elif len(linea) > 110 and "  " not in linea.strip():
                sangria = " " * (len(linea) - len(linea.lstrip()) + 2)
                self.salida.insert(tk.END, textwrap.fill(linea, 110, subsequent_indent=sangria) + "\n")
            else:
                self.salida.insert(tk.END, linea + "\n")
        self.salida.configure(state=tk.DISABLED)

    # ------------------------------------------------------------------
    def consultar(self, responder: bool = True) -> None:
        pregunta = self.pregunta_var.get().strip()
        if not pregunta:
            self._mostrar("Escriba una pregunta.")
            return
        try:
            resultado = self.agente.consultar(pregunta, responder=responder)
            self._mostrar(f"Pregunta original: {pregunta}\n\n" + resultado.texto())
            self.app.status_text.set("Agente IA: pregunta interpretada (solo lectura).")
        except Exception as exc:
            self._mostrar(f"No se pudo procesar la pregunta: {exc}")

    def propuestas(self) -> None:
        try:
            self._mostrar(self.agente.propuestas_texto())
            self.app.status_text.set("Agente IA: propuestas de mejora generadas (no se aplicó ningún cambio).")
        except Exception as exc:
            self._mostrar(f"No se pudieron generar las propuestas: {exc}")

    def limpiar(self) -> None:
        self.pregunta_var.set("")
        self._mostrar("")
