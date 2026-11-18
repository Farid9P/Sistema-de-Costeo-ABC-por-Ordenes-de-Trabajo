"""Pagina con barras de desplazamiento generales (vertical y horizontal) para las pestañas.

El contenido ocupa todo el espacio disponible, pero nunca menos de un tamaño minimo:
si la ventana es mas chica, aparecen las barras y el contenido no se aprieta.
La rueda del mouse desplaza la pagina cuando el puntero no esta sobre una tabla o
un cuadro de texto (esos se desplazan solos); Shift + rueda desplaza en horizontal.
"""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk

_PAGINA_ACTIVA: "PaginaDesplazable | None" = None
_RUEDA_INSTALADA = False
_WIDGETS_PROPIOS = (ttk.Treeview, tk.Text, tk.Listbox)


def _desplazar(evento, horizontal: bool = False) -> None:
    pagina = _PAGINA_ACTIVA
    if pagina is None or not pagina.winfo_exists():
        return
    widget = evento.widget
    if isinstance(widget, str):  # widget de una ventana secundaria ya cerrada
        return
    if isinstance(widget, _WIDGETS_PROPIOS) and not horizontal:
        return
    try:
        if widget.winfo_toplevel() is not pagina.winfo_toplevel():
            return
    except (tk.TclError, AttributeError):
        return
    if getattr(evento, "num", None) in (4, 5):
        pasos = -1 if evento.num == 4 else 1
    else:
        delta = getattr(evento, "delta", 0)
        if not delta:
            return
        pasos = -1 if delta > 0 else 1
    if horizontal:
        pagina.canvas.xview_scroll(pasos * 3, "units")
    elif pagina.puede_desplazar_vertical():
        pagina.canvas.yview_scroll(pasos * 3, "units")


def _instalar_rueda(widget) -> None:
    global _RUEDA_INSTALADA
    if _RUEDA_INSTALADA:
        return
    widget.bind_all("<MouseWheel>", lambda e: _desplazar(e), add="+")
    widget.bind_all("<Shift-MouseWheel>", lambda e: _desplazar(e, horizontal=True), add="+")
    widget.bind_all("<Button-4>", lambda e: _desplazar(e), add="+")
    widget.bind_all("<Button-5>", lambda e: _desplazar(e), add="+")
    _RUEDA_INSTALADA = True


class PaginaDesplazable(ttk.Frame):
    def __init__(self, parent, min_ancho: int = 1150, min_alto: int = 680, padding: int = 14):
        super().__init__(parent)
        self.min_ancho, self.min_alto = min_ancho, min_alto
        self.canvas = tk.Canvas(self, highlightthickness=0, bg="#DCDAD5")
        self.v = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.canvas.yview, style="Dashboard.Vertical.TScrollbar")
        self.h = ttk.Scrollbar(self, orient=tk.HORIZONTAL, command=self.canvas.xview, style="Tabla.Horizontal.TScrollbar")
        self.canvas.configure(yscrollcommand=self.v.set, xscrollcommand=self.h.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.v.grid(row=0, column=1, sticky="ns")
        self.h.grid(row=1, column=0, sticky="ew")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.contenido = ttk.Frame(self.canvas, padding=padding)
        self._ventana = self.canvas.create_window((0, 0), window=self.contenido, anchor="nw")
        self.canvas.bind("<Configure>", self._ajustar)
        self.contenido.bind("<Configure>", lambda _e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        for w in (self, self.canvas, self.contenido):
            w.bind("<Enter>", self._activar, add="+")
        _instalar_rueda(self)

    def _activar(self, _evento=None) -> None:
        global _PAGINA_ACTIVA
        _PAGINA_ACTIVA = self

    def _ajustar(self, evento) -> None:
        alto_necesario = self.contenido.winfo_reqheight()
        self.canvas.itemconfigure(
            self._ventana,
            width=max(evento.width, self.min_ancho),
            height=max(evento.height, self.min_alto, alto_necesario),
        )
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def puede_desplazar_vertical(self) -> bool:
        primero, ultimo = self.canvas.yview()
        return primero > 0 or ultimo < 1


def agregar_pagina(notebook: ttk.Notebook, titulo: str, min_ancho: int = 1150, min_alto: int = 680,
                   padding: int = 14) -> tuple[PaginaDesplazable, ttk.Frame]:
    """Crea una pestaña desplazable y devuelve (pestaña, marco de contenido)."""
    pagina = PaginaDesplazable(notebook, min_ancho, min_alto, padding)
    notebook.add(pagina, text=titulo)
    return pagina, pagina.contenido
