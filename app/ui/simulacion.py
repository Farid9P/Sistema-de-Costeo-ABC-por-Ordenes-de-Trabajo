"""Simulacion para la toma de decisiones: celdas editables (lapiz), cuadro para el nuevo valor,
"Ejecutar nuevos calculos" y "Reestablecer valores iniciales".

Que se puede cambiar (se guarda por mes en ``sim_ajustes`` y se aplica al recalcular):

- Datos importados: cualquier valor numerico de las hojas de entrada del mes visible
  (cantidades, precios, salarios, horas, importes, porcentajes...), salvo codigos.
- Proceso 1 EC-AC: PD (MOD), otros recursos y soporte de cada actividad (valor fijado).
- Proceso 2 AC-OT: peso de la actividad en la OT (dato de AC_OT_Pesos_*).
- Hoja de costos: cantidad producida (dato), CMD y CSD de la OT (valor fijado).
- Estado de resultados: ventas netas de la OT (dato de Ventas).

En un analisis de varios meses, un valor del periodo se reparte entre los meses en
proporcion a sus valores actuales (los pesos se aplican igual a todos los meses).
"""

from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk

from .. import analisis_periodos as ap
from ..database import db_session, quote_identifier

TABLAS_NO_EDITABLES = {"periodo", "empresas", "modelos_costeo", "reglas_calculo", "fuentes", "diccionario", "inductores"}
EDITABLES_CALCULO = {
    "proceso1": {"costo_pd_ac": "costo_mod_ac", "costo_otros_recursos_ac": "costo_otros_recursos_ac",
                 "costo_soporte_ac": "costo_soporte_ac"},
    "proceso2_produccion": {"peso_ac_ot": "porcentaje_inductor"},
    "proceso2_soporte": {"peso_ac_ot": "porcentaje_inductor"},
    "hoja": {"cantidad_producida": "cantidad_producida", "cmd_ot": "cmd_ot", "csd_ot": "csd_ot"},
    "estado": {"ventas_netas_soles": "ventas_netas_soles"},
}
NOMBRE_PESTANA = {"datos": "Datos importados", "proceso1": "Proceso 1 EC-AC", "proceso2_produccion": "Proceso 2 AC-OT (Produccion)",
                  "proceso2_soporte": "Proceso 2 AC-OT (Soporte)", "hoja": "Hoja de costos OT", "estado": "Estado de resultados"}


def _es_numero(valor) -> bool:
    return isinstance(valor, (int, float)) and not isinstance(valor, bool)


def _a_numero(texto: str) -> float:
    texto = str(texto).strip().replace(" ", "").replace("S/", "").replace("%", "")
    if "," in texto and "." in texto:
        texto = texto.replace(",", "")  # 1,234.56
    elif texto.count(",") > 1:
        texto = texto.replace(",", "")  # 1,234,567
    elif "," in texto:
        texto = texto.replace(",", ".")  # 1234,56 (coma decimal)
    return float(texto)


class SimulacionMixin:
    def _init_simulacion(self) -> None:
        self.cambios_pendientes = 0
        self._lapices: dict[str, tk.Label] = {}
        self._marcadores: dict = {}

    # ------------------------------------------------------------------ interfaz
    def _build_sim_header(self, header) -> None:
        self.sim_buttons = ttk.Frame(header)
        self.sim_buttons.grid(row=0, column=1, rowspan=2, sticky=tk.E, padx=(12, 0))
        self.recalcular_button = ttk.Button(self.sim_buttons, text="Ejecutar nuevos cálculos", style="Accion.TButton",
                                            command=self.ejecutar_nuevos_calculos)
        self.recalcular_button.pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(self.sim_buttons, text="Reestablecer valores iniciales", command=self.restablecer_valores).pack(side=tk.LEFT)
        self.sim_buttons.grid_remove()

    def _build_sim_banner(self, root, antes_de) -> None:
        self.sim_banner = tk.Frame(root, bg="#FEF3C7", highlightthickness=1, highlightbackground="#F59E0B", padx=10, pady=6)
        self.sim_banner_text = tk.Label(self.sim_banner, bg="#FEF3C7", fg="#78350F", font=("Segoe UI", 9, "bold"),
                                        justify=tk.LEFT, anchor="w")
        self.sim_banner_text.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(self.sim_banner, text="Ver comparacion y cambios", command=self.ver_detalle_simulacion).pack(side=tk.RIGHT)
        self._banner_antes = antes_de

    def mostrar_controles_simulacion(self, visible: bool) -> None:
        if hasattr(self, "sim_buttons"):
            (self.sim_buttons.grid if visible else self.sim_buttons.grid_remove)()
        if not visible and hasattr(self, "sim_banner"):
            self.sim_banner.pack_forget()
        elif visible:
            self.actualizar_banner_simulacion()

    def habilitar_edicion(self, claves: list[str]) -> None:
        for clave in claves:
            arbol = self.trees.get(clave)
            if arbol is None:
                continue
            lapiz = tk.Label(arbol, text="✎", bg="#FDE68A", fg="#92400E", font=("Segoe UI Symbol", 10, "bold"), padx=2)
            self._lapices[clave] = lapiz
            arbol.tag_configure("modificado", background="#FEF9C3")
            arbol.bind("<Motion>", lambda e, c=clave: self._al_mover(c, e), add="+")
            arbol.bind("<Leave>", lambda e, c=clave: self._ocultar_lapiz(c), add="+")
            arbol.bind("<ButtonRelease-1>", lambda e, c=clave: self._al_clic(c, e), add="+")

    def _celda(self, clave: str, evento):
        arbol = self.trees[clave]
        if arbol.identify("region", evento.x, evento.y) != "cell":
            return None
        item = arbol.identify_row(evento.y)
        col_id = arbol.identify_column(evento.x)
        try:
            columna = arbol["columns"][int(col_id.replace("#", "")) - 1]
        except (ValueError, IndexError):
            return None
        fila = self.tree_row_data.get(clave, {}).get(item)
        if not fila or not self.celda_editable(clave, fila, columna):
            return None
        return item, col_id, columna, fila

    def _al_mover(self, clave: str, evento) -> None:
        arbol = self.trees[clave]
        celda = self._celda(clave, evento)
        lapiz = self._lapices[clave]
        if celda is None:
            arbol.configure(cursor="")
            lapiz.place_forget()
            return
        item, col_id, _columna, _fila = celda
        arbol.configure(cursor="pencil")
        caja = arbol.bbox(item, col_id)
        if caja:
            x, y, ancho, alto = caja
            lapiz.place(x=x + ancho - 18, y=y + 1, width=17, height=max(alto - 2, 10))

    def _ocultar_lapiz(self, clave: str) -> None:
        self.trees[clave].configure(cursor="")
        self._lapices[clave].place_forget()

    def _al_clic(self, clave: str, evento) -> None:
        celda = self._celda(clave, evento)
        if celda is not None:
            self._ocultar_lapiz(clave)
            self.abrir_editor(clave, *celda)

    # ------------------------------------------------------------------ reglas
    def celda_editable(self, clave: str, fila: dict, columna: str) -> bool:
        if not getattr(self, "calculo_ejecutado", False) or "mensaje" in fila:
            return False
        if clave == "datos":
            tabla = self.selected_imported_table_name()
            if tabla not in ap.TABLAS_IMPORTADAS or tabla in TABLAS_NO_EDITABLES or "_fila" not in fila:
                return False
            if columna.startswith("_") or columna.endswith("_id") or columna == "batch_id" or columna.startswith("redondear"):
                return False
            return _es_numero(fila.get(columna))
        return columna in EDITABLES_CALCULO.get(clave, {}) and fila.get(columna) is not None

    # ------------------------------------------------------------------ editor
    def abrir_editor(self, clave: str, item: str, col_id: str, columna: str, fila: dict) -> None:
        actual = fila.get(columna)
        ventana = tk.Toplevel(self)
        ventana.title("Simular un nuevo valor")
        ventana.transient(self)
        ventana.resizable(False, False)
        marco = ttk.Frame(ventana, padding=14)
        marco.pack(fill=tk.BOTH, expand=True)
        descripcion = self._describir_fila(clave, fila)
        ttk.Label(marco, text=f"{NOMBRE_PESTANA.get(clave, clave)} — {descripcion}", font=("Segoe UI", 10, "bold"),
                  wraplength=460).grid(row=0, column=0, columnspan=2, sticky=tk.W)
        ttk.Label(marco, text=f"Campo: {columna}").grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(8, 0))
        ttk.Label(marco, text=f"Valor actual: {self._fmt(actual)}").grid(row=2, column=0, columnspan=2, sticky=tk.W)
        nota = self._nota_edicion(clave, fila, columna)
        if nota:
            ttk.Label(marco, text=nota, foreground="#92400E", wraplength=460, justify=tk.LEFT).grid(
                row=3, column=0, columnspan=2, sticky=tk.W, pady=(6, 0))
        ttk.Label(marco, text="Nuevo valor (punto decimal):").grid(row=4, column=0, sticky=tk.W, pady=(10, 0))
        inicial = f"{actual:.6f}".rstrip("0").rstrip(".") if _es_numero(actual) else ("" if actual is None else str(actual))
        valor = tk.StringVar(value=inicial)
        entrada = ttk.Entry(marco, textvariable=valor, width=24)
        entrada.grid(row=4, column=1, sticky=tk.W, pady=(10, 0), padx=(6, 0))
        entrada.select_range(0, tk.END)
        entrada.focus_set()
        mensaje = ttk.Label(marco, text="", foreground="#B91C1C")
        mensaje.grid(row=5, column=0, columnspan=2, sticky=tk.W)

        def aceptar(_e=None):
            try:
                nuevo = _a_numero(valor.get())
            except ValueError:
                mensaje.configure(text="Ingrese un numero (use punto o coma decimal).")
                return
            if nuevo < 0 and columna not in ("margen_operativo",):
                mensaje.configure(text="El valor no puede ser negativo.")
                return
            try:
                self.aplicar_cambio(clave, fila, columna, nuevo)
            except Exception as exc:  # se muestra en el cuadro
                mensaje.configure(text=str(exc))
                return
            arbol = self.trees[clave]
            arbol.set(item, columna, self._fmt(nuevo))
            fila[columna] = nuevo
            arbol.item(item, tags=("modificado",))
            ventana.destroy()

        botones = ttk.Frame(marco)
        botones.grid(row=6, column=0, columnspan=2, sticky=tk.E, pady=(12, 0))
        ttk.Button(botones, text="Aceptar", command=aceptar).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(botones, text="Cancelar", command=ventana.destroy).pack(side=tk.LEFT)
        ventana.bind("<Return>", aceptar)
        ventana.bind("<Escape>", lambda _e: ventana.destroy())
        ventana.update_idletasks()
        x = self.winfo_pointerx() - ventana.winfo_width() // 2
        y = self.winfo_pointery() + 12
        ventana.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        ventana.grab_set()

    @staticmethod
    def _fmt(valor) -> str:
        if _es_numero(valor):
            return f"{valor:,.6f}".rstrip("0").rstrip(".") if abs(valor) < 10 else f"{valor:,.4f}".rstrip("0").rstrip(".")
        return "" if valor is None else str(valor)

    def _describir_fila(self, clave: str, fila: dict) -> str:
        if clave == "datos":
            mes = ap.nombre_mes(*self.mes_visible) if self.mes_visible else ""
            ident = next((f"{k} {v}" for k, v in fila.items() if k.endswith("_id") and v), f"fila {fila.get('_fila')}")
            return f"{self.selected_imported_table_name()} | {ident} | {mes}"
        partes = [str(fila.get(k)) for k in ("codigo_actividad", "nombre_actividad", "codigo_ot", "codigo_producto") if fila.get(k)]
        return " | ".join(partes)

    def _nota_edicion(self, clave: str, fila: dict, columna: str) -> str:
        notas = []
        if self.es_multimes() and clave != "datos":
            notas.append("Periodo de varios meses: " + ("el peso se aplica igual a todos los meses." if columna == "peso_ac_ot"
                         else "el nuevo total se reparte entre los meses en proporcion a sus valores actuales."))
        if clave in ("proceso2_produccion", "proceso2_soporte"):
            filas = [f for f in self.tree_row_data.get(clave, {}).values() if f.get("actividad_id") == fila.get("actividad_id")]
            suma = sum(float(f.get("peso_ac_ot") or 0) for f in filas)
            notas.append(f"Los pesos de la actividad suman {suma:.4f} (100 % = 1). Si sube uno, baje otro para no "
                         "asignar mas (o menos) del costo de la actividad.")
        if clave == "proceso1":
            notas.append("Se fija el valor de la actividad; el costo de conversion, el total y su asignacion a las OT se recalculan.")
        if clave == "hoja" and columna in ("cmd_ot", "csd_ot"):
            notas.append("Se fija el costo de la OT (el kardex no cambia).")
        if clave == "datos" and self.selected_imported_table_name() == "ventas":
            notas.append("Ventas: cantidad × valor de venta unitario = ventas netas (se ajusta el otro valor).")
        return " ".join(notas)

    # ------------------------------------------------------------------ registrar cambios
    def _ots_del_mes(self, mes, fila: dict) -> list[str]:
        if not self.es_multimes():
            return [fila.get("ot_id")]
        datos = ap.datos_mes(self.db_path, *mes, con_ajustes=False)
        producto = fila.get("codigo_producto")
        objetos = {o.get("objeto_costo_id") for o in datos.get("objetos_costo", []) if o.get("codigo_producto") == producto}
        return [o.get("ot_id") for o in datos.get("ordenes_trabajo", []) if o.get("objeto_costo_id") in objetos]

    def _repartir(self, actuales: dict, nuevo: float) -> dict:
        total = sum(actuales.values())
        if total:
            return {k: v * nuevo / total for k, v in actuales.items()}
        return {k: nuevo / len(actuales) for k in actuales} if actuales else {}

    def _registrar_dato(self, mes, tabla: str, indice: int, columna: str, nuevo: float, descripcion: str = "") -> None:
        original = ap.datos_mes(self.db_path, *mes, con_ajustes=False)[tabla][indice].get(columna)
        ap.registrar_ajuste(self.db_path, mes[0], mes[1], "dato", tabla, columna, original, nuevo, fila=indice,
                            descripcion=descripcion)

    def _registrar_resultado(self, mes, tabla: str, clave: str, columna: str, nuevo: float, original: float | None,
                             descripcion: str = "") -> None:
        ap.registrar_ajuste(self.db_path, mes[0], mes[1], "resultado", tabla, columna, original, nuevo, clave=clave,
                            descripcion=descripcion)

    def aplicar_cambio(self, clave: str, fila: dict, columna: str, nuevo: float) -> None:
        meses = self.analisis["meses"]
        descripcion = f"{NOMBRE_PESTANA.get(clave, clave)}: {self._describir_fila(clave, fila)} | {columna}"
        if clave == "datos":
            self._cambio_dato_visible(fila, columna, nuevo, descripcion)
        elif clave == "proceso1":
            col = EDITABLES_CALCULO["proceso1"][columna]
            actuales, iniciales = {}, {}
            base = ap._cargar_meses(self.db_path, meses)
            for mes in meses:
                actual = next((a for a in self.resultados_actuales[mes]["actividad_total"] if a["actividad_id"] == fila["actividad_id"]), {})
                inicial = next((a for a in base[mes]["actividad_total"] if a["actividad_id"] == fila["actividad_id"]), {})
                actuales[mes], iniciales[mes] = float(actual.get(col) or 0), inicial.get(col)
            for mes, valor in self._repartir(actuales, nuevo).items():
                self._registrar_resultado(mes, "actividad_total", fila["actividad_id"], col, valor, iniciales[mes], descripcion)
        elif clave in ("proceso2_produccion", "proceso2_soporte"):
            tabla = "ac_ot_pesos_produccion" if clave == "proceso2_produccion" else "ac_ot_pesos_soporte"
            for mes in meses:
                ots = self._ots_del_mes(mes, fila)
                filas = ap.datos_mes(self.db_path, *mes, con_ajustes=False).get(tabla, [])
                for i, f in enumerate(filas):
                    if f.get("actividad_id") == fila.get("actividad_id") and f.get("ot_id") in ots:
                        self._registrar_dato(mes, tabla, i, "porcentaje_inductor", nuevo, descripcion)
        elif clave == "hoja" and columna == "cantidad_producida":
            self._repartir_dato(meses, fila, "ordenes_trabajo", "cantidad_producida", nuevo, descripcion)
        elif clave == "hoja":
            tabla = "cmd_ot" if columna == "cmd_ot" else "csd_ot"
            base = ap._cargar_meses(self.db_path, meses)
            actuales, iniciales, ots = {}, {}, {}
            for mes in meses:
                for ot in self._ots_del_mes(mes, fila):
                    actual = next((h for h in self.resultados_actuales[mes]["hoja_costos_ot"] if h["ot_id"] == ot), {})
                    inicial = next((h for h in base[mes]["hoja_costos_ot"] if h["ot_id"] == ot), {})
                    actuales[(mes, ot)], iniciales[(mes, ot)] = float(actual.get(columna) or 0), inicial.get(columna)
            for (mes, ot), valor in self._repartir(actuales, nuevo).items():
                self._registrar_resultado(mes, tabla, ot, columna, valor, iniciales[(mes, ot)], descripcion)
        elif clave == "estado":
            self._repartir_dato(meses, fila, "ventas", "ventas_netas_soles", nuevo, descripcion, ajustar_unitario=True)
        self.cambios_pendientes += 1
        self.actualizar_banner_simulacion()
        self.status_text.set(f"Cambio registrado ({descripcion} = {self._fmt(nuevo)}). Presione 'Ejecutar nuevos cálculos'.")

    def _repartir_dato(self, meses, fila, tabla: str, columna: str, nuevo: float, descripcion: str,
                       ajustar_unitario: bool = False) -> None:
        actuales = {}
        for mes in meses:
            datos = ap.datos_mes(self.db_path, *mes)
            ots = self._ots_del_mes(mes, fila)
            for i, f in enumerate(datos.get(tabla, [])):
                if f.get("ot_id") in ots:
                    actuales[(mes, i)] = (float(f.get(columna) or 0), f)
        repartido = self._repartir({k: v[0] for k, v in actuales.items()}, nuevo)
        for (mes, i), valor in repartido.items():
            self._registrar_dato(mes, tabla, i, columna, valor, descripcion)
            if ajustar_unitario:
                cantidad = float(actuales[(mes, i)][1].get("cantidad_vendida") or 0)
                if cantidad:
                    self._registrar_dato(mes, tabla, i, "valor_venta_unitario_soles", valor / cantidad, descripcion)

    def _cambio_dato_visible(self, fila: dict, columna: str, nuevo: float, descripcion: str) -> None:
        tabla = self.selected_imported_table_name()
        mes = self.mes_visible
        indice = int(fila["_fila"]) - 1
        cambios = {columna: nuevo}
        if tabla == "ventas":
            cantidad = float(fila.get("cantidad_vendida") or 0)
            unitario = float(fila.get("valor_venta_unitario_soles") or 0)
            if columna == "cantidad_vendida":
                cambios["ventas_netas_soles"] = nuevo * unitario
            elif columna == "valor_venta_unitario_soles":
                cambios["ventas_netas_soles"] = cantidad * nuevo
            elif columna == "ventas_netas_soles" and cantidad:
                cambios["valor_venta_unitario_soles"] = nuevo / cantidad
        # primero se guardan los ajustes (cada uno en su propia transaccion) y despues se actualiza la tabla visible;
        # escribir con dos conexiones abiertas a la vez bloqueaba la base ("database is locked").
        for col, valor in cambios.items():
            self._registrar_dato(mes, tabla, indice, col, valor, descripcion)
        with db_session(self.db_path) as conn:
            for col, valor in cambios.items():
                conn.execute(f"UPDATE {quote_identifier(tabla)} SET {quote_identifier(col)} = ? WHERE rowid = ?",
                             (valor, indice + 1))
        if len(cambios) > 1:
            self.after(10, self.load_selected_imported_table)

    # ------------------------------------------------------------------ marcas de filas modificadas
    def preparar_marcadores(self) -> None:
        marcas = {"datos": set(), "actividad": set(), "ot": set(), "peso_prod": set(), "peso_sop": set()}
        if not self.analisis:
            self._marcadores = marcas
            return
        cache = {}
        for aj in ap.ajustes(self.db_path, self.analisis["meses"]):
            mes = (aj["anio"], aj["mes"])
            if aj["tipo"] == "resultado":
                if aj["tabla"] == "actividad_total":
                    marcas["actividad"].add(aj["clave"])
                else:
                    marcas["ot"].add(self._clave_ot(mes, aj["clave"], cache))
                continue
            marcas["datos"].add((mes, aj["tabla"], aj["fila"]))
            if mes not in cache:
                cache[mes] = ap.datos_mes(self.db_path, *mes, con_ajustes=False)
            filas = cache[mes].get(aj["tabla"], [])
            f = filas[aj["fila"]] if aj["fila"] is not None and aj["fila"] < len(filas) else {}
            ot = self._clave_ot(mes, f.get("ot_id"), cache)
            if aj["tabla"] in ("ordenes_trabajo", "ventas"):
                marcas["ot"].add(ot)
            elif aj["tabla"] == "ac_ot_pesos_produccion":
                marcas["peso_prod"].add((f.get("actividad_id"), ot))
            elif aj["tabla"] == "ac_ot_pesos_soporte":
                marcas["peso_sop"].add((f.get("actividad_id"), ot))
        self._marcadores = marcas

    def _clave_ot(self, mes, ot_id, cache) -> str:
        if not self.es_multimes():
            return ot_id
        if mes not in cache:
            cache[mes] = ap.datos_mes(self.db_path, *mes, con_ajustes=False)
        datos = cache[mes]
        objeto = next((o.get("objeto_costo_id") for o in datos.get("ordenes_trabajo", []) if o.get("ot_id") == ot_id), None)
        return next((o.get("codigo_producto") for o in datos.get("objetos_costo", []) if o.get("objeto_costo_id") == objeto), ot_id)

    def fila_modificada(self, clave: str, fila: dict) -> bool:
        m = self._marcadores
        if not m:
            return False
        if clave == "datos":
            if "_fila" not in fila or not self.mes_visible:
                return False
            return (self.mes_visible, self.selected_imported_table_name(), int(fila["_fila"]) - 1) in m["datos"]
        if clave == "proceso1":
            return fila.get("actividad_id") in m["actividad"]
        ot = fila.get("ot_id")
        if clave in ("hoja", "estado"):
            return ot in m["ot"]
        if clave == "proceso2_produccion":
            return (fila.get("actividad_id"), ot) in m["peso_prod"]
        if clave == "proceso2_soporte":
            return (fila.get("actividad_id"), ot) in m["peso_sop"]
        return False

    # ------------------------------------------------------------------ acciones
    def ejecutar_nuevos_calculos(self) -> None:
        if not self.analisis or not self.calculo_ejecutado:
            messagebox.showwarning("Simulacion", "Primero ejecute el calculo del periodo desde la pantalla inicial.")
            return
        try:
            self.ejecutar_analisis()
        except Exception as exc:
            messagebox.showerror("Simulacion", f"No se pudo recalcular:\n{exc}")
            return
        messagebox.showinfo("Simulacion", "Calculos actualizados con los valores modificados.\n\n" + self.texto_comparacion())

    def restablecer_valores(self) -> None:
        if not self.analisis:
            return
        n = len(ap.ajustes(self.db_path))
        if not n:
            messagebox.showinfo("Simulacion", "No hay valores modificados: ya se muestran los valores iniciales.")
            return
        if not messagebox.askyesno("Simulacion", f"¿Volver a los valores iniciales de los Excel? Se descartan {n} cambio(s)."):
            return
        ap.borrar_ajustes(self.db_path)
        self.ejecutar_analisis()
        messagebox.showinfo("Simulacion", "Se reestablecieron los valores iniciales y se recalculo el periodo.")

    def contar_cambios(self) -> int:
        """Cambios hechos por el usuario (un cambio puede tocar varios meses o columnas relacionadas)."""
        registros = ap.ajustes(self.db_path, self.analisis["meses"]) if self.analisis else []
        return len({a.get("descripcion") or f"{a['tabla']}|{a['fila']}|{a['clave']}|{a['columna']}" for a in registros})

    # ------------------------------------------------------------------ comparacion
    def comparacion(self) -> list[tuple[str, float, float, str]]:
        ini = (self.payload_inicial or {}).get("kpis") or {}
        sim = (self.analisis_payload or {}).get("kpis") or {}
        filas = [("Ventas netas", "ventas", "S/"), ("Costo total OT", "costo_total", "S/"), ("Margen bruto", "margen_bruto", "S/"),
                 ("Margen operativo", "margen_operativo", "S/"), ("Rentabilidad operativa", "rentabilidad_global", "%"),
                 ("Costo unitario medio", "costo_unitario_medio", "S/u")]
        return [(n, ini.get(k, 0.0), sim.get(k, 0.0), u) for n, k, u in filas]

    def texto_comparacion(self) -> str:
        lineas = []
        for nombre, ini, sim, unidad in self.comparacion():
            if unidad == "%":
                lineas.append(f"{nombre}: {ini * 100:,.2f}% → {sim * 100:,.2f}% ({(sim - ini) * 100:+,.2f} pts)")
            else:
                lineas.append(f"{nombre}: S/ {ini:,.2f} → S/ {sim:,.2f} ({sim - ini:+,.2f})")
        return "\n".join(lineas)

    def actualizar_banner_simulacion(self) -> None:
        if not hasattr(self, "sim_banner"):
            return
        n = self.contar_cambios()
        if not n or not self.calculo_ejecutado:
            self.sim_banner.pack_forget()
            return
        pendiente = f" | {self.cambios_pendientes} cambio(s) sin calcular: presione 'Ejecutar nuevos cálculos'" if self.cambios_pendientes else ""
        comp = {c[0]: c for c in self.comparacion()}
        _, mo_i, mo_s, _ = comp["Margen operativo"]
        _, r_i, r_s, _ = comp["Rentabilidad operativa"]
        self.sim_banner_text.configure(
            text=f"SIMULACION ({n} cambio(s)) — Margen operativo: S/ {mo_i:,.2f} → S/ {mo_s:,.2f} "
                 f"({mo_s - mo_i:+,.2f}) | Rentabilidad: {r_i * 100:,.2f}% → {r_s * 100:,.2f}%{pendiente}"
        )
        if not self.sim_banner.winfo_ismapped():
            self.sim_banner.pack(fill=tk.X, pady=(0, 6), before=self._banner_antes)

    def ver_detalle_simulacion(self) -> None:
        ventana = tk.Toplevel(self)
        ventana.title("Simulacion: valores iniciales vs. nuevos")
        ventana.geometry("1100x620")
        marco = ttk.Frame(ventana, padding=10)
        marco.pack(fill=tk.BOTH, expand=True)
        titulo = self.analisis["titulo"] if self.analisis else ""
        ttk.Label(marco, text=f"Resultados del periodo {titulo}", font=("Segoe UI", 11, "bold")).pack(anchor=tk.W)
        comp = ttk.Treeview(marco, columns=("concepto", "inicial", "simulado", "diferencia"), show="headings", height=7)
        for col, texto, ancho in (("concepto", "Concepto", 260), ("inicial", "Valores iniciales", 200),
                                  ("simulado", "Con cambios", 200), ("diferencia", "Diferencia", 200)):
            comp.heading(col, text=texto)
            comp.column(col, width=ancho, anchor=tk.W if col == "concepto" else tk.E)
        for nombre, ini, sim, unidad in self.comparacion():
            if unidad == "%":
                comp.insert("", tk.END, values=(nombre, f"{ini * 100:,.2f}%", f"{sim * 100:,.2f}%", f"{(sim - ini) * 100:+,.2f} pts"))
            else:
                comp.insert("", tk.END, values=(nombre, f"{ini:,.2f}", f"{sim:,.2f}", f"{sim - ini:+,.2f}"))
        ini_ot = {f["codigo_producto"]: f for f in (self.payload_inicial or {}).get("estado", [])}
        for f in (self.analisis_payload or {}).get("estado", []):
            i = ini_ot.get(f["codigo_producto"], {})
            comp.insert("", tk.END, values=(f"Margen operativo {f['codigo_ot']}", f"{i.get('margen_operativo', 0):,.2f}",
                                            f"{f['margen_operativo']:,.2f}", f"{f['margen_operativo'] - i.get('margen_operativo', 0):+,.2f}"))
        comp.configure(height=min(14, 6 + len(ini_ot)))
        comp.pack(fill=tk.X, pady=(6, 10))
        ttk.Label(marco, text="Valores modificados", font=("Segoe UI", 11, "bold")).pack(anchor=tk.W)
        cols = ("mes", "que", "campo", "inicial", "nuevo")
        lista = ttk.Treeview(marco, columns=cols, show="headings", height=10)
        for col, texto, ancho in (("mes", "Mes", 110), ("que", "Donde", 520), ("campo", "Campo", 190),
                                  ("inicial", "Valor inicial", 120), ("nuevo", "Valor nuevo", 120)):
            lista.heading(col, text=texto)
            lista.column(col, width=ancho, anchor=tk.E if col in ("inicial", "nuevo") else tk.W)
        for aj in ap.ajustes(self.db_path, self.analisis["meses"] if self.analisis else None):
            lista.insert("", tk.END, values=(ap.nombre_mes(aj["anio"], aj["mes"]), aj.get("descripcion") or aj["tabla"],
                                             aj["columna"], self._fmt(aj["valor_original"]), self._fmt(aj["valor_nuevo"])))
        lista.pack(fill=tk.BOTH, expand=True)
        ttk.Button(marco, text="Cerrar", command=ventana.destroy).pack(anchor=tk.E, pady=(8, 0))
