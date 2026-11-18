"""Visor de imagen con zoom y barras de desplazamiento (grafico del modelo de costeo)."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import ttk


class VisorImagen(ttk.Frame):
    """Canvas con la imagen ajustada a la ventana, zoom (+ / − / 100 %), barras y arrastre con el mouse.

    Rueda: desplaza en vertical (Shift + rueda: horizontal; Ctrl + rueda: zoom).
    """

    ZOOM_MIN, ZOOM_MAX = 0.1, 4.0

    def __init__(self, parent, height: int = 600, boton_ventana: bool = True, titulo_ventana: str = "Grafico"):
        super().__init__(parent)
        self.ruta: Path | None = None
        self.zoom: float | None = None  # None = ajustar a la ventana
        self.escala_actual = 1.0
        self._foto = None
        self._fuente = None
        self._titulo_ventana = titulo_ventana

        barra = ttk.Frame(self)
        barra.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6))
        ttk.Button(barra, text="Ajustar a la ventana", command=self.ajustar).pack(side=tk.LEFT)
        ttk.Button(barra, text="Ajustar al ancho", command=self.ajustar_ancho).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(barra, text="−", width=3, command=lambda: self.cambiar_zoom(1 / 1.25)).pack(side=tk.LEFT, padx=(12, 0))
        ttk.Button(barra, text="+", width=3, command=lambda: self.cambiar_zoom(1.25)).pack(side=tk.LEFT, padx=(4, 0))
        ttk.Button(barra, text="100 %", command=lambda: self.fijar_zoom(1.0)).pack(side=tk.LEFT, padx=(4, 0))
        self.zoom_var = tk.StringVar(value="")
        ttk.Label(barra, textvariable=self.zoom_var, width=12).pack(side=tk.LEFT, padx=(8, 0))
        if boton_ventana:
            ttk.Button(barra, text="Ver en pantalla completa", command=self.abrir_ventana).pack(side=tk.RIGHT)
        ttk.Label(barra, text="Rueda: desplazar | Ctrl + rueda: zoom | arrastrar: mover",
                  foreground="#64748B").pack(side=tk.RIGHT, padx=(0, 12))

        self.canvas = tk.Canvas(self, bg="#FFFFFF", height=height, highlightthickness=1, highlightbackground="#CBD5E1")
        self.v = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.canvas.yview, style="Tabla.Vertical.TScrollbar")
        self.h = ttk.Scrollbar(self, orient=tk.HORIZONTAL, command=self.canvas.xview, style="Tabla.Horizontal.TScrollbar")
        self.canvas.configure(yscrollcommand=self.v.set, xscrollcommand=self.h.set)
        self.canvas.grid(row=1, column=0, sticky="nsew")
        self.v.grid(row=1, column=1, sticky="ns")
        self.h.grid(row=2, column=0, sticky="ew")
        self.rowconfigure(1, weight=1)
        self.columnconfigure(0, weight=1)

        self.canvas.bind("<Configure>", lambda _e: self.dibujar())
        self.canvas.bind("<ButtonPress-1>", lambda e: self.canvas.scan_mark(e.x, e.y))
        self.canvas.bind("<B1-Motion>", lambda e: self.canvas.scan_dragto(e.x, e.y, gain=1))
        for widget in (self.canvas,):
            widget.bind("<MouseWheel>", self._rueda)
            widget.bind("<Shift-MouseWheel>", self._rueda_horizontal)
            widget.bind("<Control-MouseWheel>", self._rueda_zoom)
            widget.bind("<Button-4>", lambda e: self._desplazar(-1))
            widget.bind("<Button-5>", lambda e: self._desplazar(1))

    # ------------------------------------------------------------------ API
    def mostrar(self, ruta: Path | None) -> None:
        self.ruta = Path(ruta) if ruta else None
        self._fuente = None
        if self.ruta and self.ruta.exists():
            try:
                from PIL import Image

                with Image.open(self.ruta) as img:
                    self._fuente = img.copy()
            except Exception:
                self._fuente = None
        self.dibujar()

    def ajustar(self) -> None:
        self.zoom = None
        self.dibujar()

    def ajustar_ancho(self) -> None:
        if self._fuente is None:
            return
        ancho = max(self.canvas.winfo_width() - 10, 1)
        self.fijar_zoom(ancho / self._fuente.width)

    def fijar_zoom(self, valor: float) -> None:
        self.zoom = min(max(valor, self.ZOOM_MIN), self.ZOOM_MAX)
        self.dibujar()

    def cambiar_zoom(self, factor: float) -> None:
        self.fijar_zoom((self.zoom or self.escala_actual) * factor)

    def tiene_desborde_vertical(self) -> bool:
        primero, ultimo = self.canvas.yview()
        return primero > 0 or ultimo < 1

    # ------------------------------------------------------------------ dibujo
    def dibujar(self) -> None:
        self.canvas.delete("all")
        if self.ruta is None:
            self.zoom_var.set("")
            self.canvas.configure(scrollregion=(0, 0, 1, 1))
            return
        ancho = max(self.canvas.winfo_width(), 1)
        alto = max(self.canvas.winfo_height(), 1)
        if self._fuente is None:
            self._dibujar_sin_pil(ancho, alto)
            return
        from PIL import Image, ImageTk

        fuente = self._fuente
        if self.zoom is None:
            escala = min((ancho - 10) / fuente.width, (alto - 10) / fuente.height, 2.0)
        else:
            escala = self.zoom
        escala = max(escala, 0.05)
        self.escala_actual = escala
        w, h = max(1, int(fuente.width * escala)), max(1, int(fuente.height * escala))
        self._foto = ImageTk.PhotoImage(fuente.resize((w, h), Image.LANCZOS))
        x = max((ancho - w) // 2, 5)
        y = max((alto - h) // 2, 5)
        self.canvas.create_image(x, y, anchor=tk.NW, image=self._foto)
        self.canvas.configure(scrollregion=(0, 0, max(ancho, w + 10), max(alto, h + 10)))
        self.zoom_var.set(f"Zoom {escala * 100:,.0f} %")

    def _dibujar_sin_pil(self, ancho: int, alto: int) -> None:
        try:
            foto = tk.PhotoImage(file=str(self.ruta))
        except tk.TclError:
            return
        if self.zoom is None:
            factor = max(1, -(-foto.width() // max(ancho - 10, 1)), -(-foto.height() // max(alto - 10, 1)))
            foto = foto.subsample(factor, factor)
        self._foto = foto
        self.canvas.create_image(max((ancho - foto.width()) // 2, 5), max((alto - foto.height()) // 2, 5),
                                 anchor=tk.NW, image=foto)
        self.canvas.configure(scrollregion=(0, 0, max(ancho, foto.width() + 10), max(alto, foto.height() + 10)))

    # ------------------------------------------------------------------ rueda
    def _desplazar(self, pasos: int) -> str:
        self.canvas.yview_scroll(pasos * 3, "units")
        return "break"

    def _rueda(self, evento) -> str:
        return self._desplazar(-1 if evento.delta > 0 else 1)

    def _rueda_horizontal(self, evento) -> str:
        self.canvas.xview_scroll(-3 if evento.delta > 0 else 3, "units")
        return "break"

    def _rueda_zoom(self, evento) -> str:
        self.cambiar_zoom(1.1 if evento.delta > 0 else 1 / 1.1)
        return "break"

    # ------------------------------------------------------------------ ventana
    def abrir_ventana(self) -> None:
        if self.ruta is None:
            return
        ventana = tk.Toplevel(self)
        ventana.title(self._titulo_ventana)
        ventana.geometry(f"{ventana.winfo_screenwidth() - 80}x{ventana.winfo_screenheight() - 120}+20+20")
        try:
            ventana.state("zoomed")
        except tk.TclError:
            pass
        visor = VisorImagen(ventana, height=600, boton_ventana=False)
        visor.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        ttk.Button(ventana, text="Cerrar", command=ventana.destroy).pack(pady=(0, 8))
        visor.mostrar(self.ruta)
