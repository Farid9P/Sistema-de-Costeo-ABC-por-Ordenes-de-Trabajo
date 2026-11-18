"""Respuestas del agente a cada pregunta del catalogo base.

Todas las respuestas se arman LEYENDO la base SQLite (datos importados y
calculos ya generados por el sistema). El agente no ejecuta calculos que
escriban en la base, no importa Excel, no genera archivos ni corre pruebas.
"""

from __future__ import annotations

from pathlib import Path

from ..constants import (
    DEFAULT_EXCEL_PATH,
    OUTPUT_DOCUMENTS,
    OUTPUT_REPORTS,
    REPORTS_DIR,
    SHEET_TO_TABLE,
)
from .formato import Respuesta, cifra, num, pct, soles, tabla
from .interprete import Interpretacion
from .kardex import kardex_desde_base, salidas_por_material
from .lectura import LectorSoloLectura

HOJA_DE_TABLA = {tabla_bd: hoja for hoja, tabla_bd in SHEET_TO_TABLE.items()}

# Pregunta -> (tabla de datos importados) para preguntas de trazabilidad (5 y 6)
TEMAS_IMPORTADOS = [
    ("VENTAS", "ventas"), ("KARDEX", "compras"), ("CMD", "compras"),
    ("TIEMPO_STD", "Personal_directos"), ("FD", "Personal_directos"), ("MOD", "Personal_directos"),
    ("MANTENIMIENTO", "Soporte_Mantenimiento(CI)"), ("CALIDAD", "Soporte_Control_Calidad(CI)"),
    ("INGENIERIA", "Soporte_Ingenieria(CI)"), ("CSD", "servicios_directos"),
    ("ENT_RECURSO", "costo_otros_recursos(CI)"), ("RECURSO", "costo_otros_recursos(CI)"),
    ("PESO", "ac_ot_pesos_produccion"), ("ENT_ACTIVIDAD", "actividades"), ("ACTIVIDAD", "actividades"),
    ("UNIDADES", "ordenes_trabajo"), ("EMPRESA", "empresas"), ("PERIODO", "periodo"),
    ("MODELO", "modelos_costeo"), ("ENT_OT", "ordenes_trabajo"), ("OBJETO", "ordenes_trabajo"),
]
TEMAS_CALCULADOS = [
    ("RENTABILIDAD", "estado_resultados_ot"), ("MARGEN_OPERATIVO", "estado_resultados_ot"),
    ("MARGEN_BRUTO", "estado_resultados_ot"), ("MARGEN", "estado_resultados_ot"),
    ("UNITARIO", "hoja_costos_ot"), ("COSTO_VENTAS", "hoja_costos_ot"), ("TOTAL", "hoja_costos_ot"),
]

CALCULO_REQUERIDO = {str(n) for n in range(9, 47)} | {"E1", "E2"}
CALCULO_REQUERIDO -= {"10", "13", "14", "18", "44", "46"}


class Consultor:
    def __init__(self, lector: LectorSoloLectura, interpretacion: Interpretacion):
        self.L = lector
        self.i = interpretacion
        self.run = lector.ultimo_run() if lector.disponible else None
        self.dir = interpretacion.direccion
        self.ot_sel = [e.id for e in interpretacion.entidades_de("OT")]
        self.act_sel = [e.id for e in interpretacion.entidades_de("ACTIVIDAD")]
        self.rec_sel = [e.id for e in interpretacion.entidades_de("RECURSO")]
        self._ots_cache: dict | None = None

    # ------------------------------------------------------------------
    # despacho
    # ------------------------------------------------------------------
    def responder(self) -> Respuesta:
        pregunta = self.i.pregunta
        if pregunta is None:
            return Respuesta("Sin interpretación", "No pude relacionar la pregunta con el catálogo.", sin_datos=True)
        if not self.L.disponible:
            return Respuesta("Sin base de datos", "Aún no existe la base SQLite: importe el Excel maestro desde la pantalla inicial.",
                             sin_datos=True)
        if pregunta.codigo in CALCULO_REQUERIDO and self.run is None:
            return Respuesta(
                "Sin calculo",
                "Todavía no hay un cálculo completado del Modelo ABC por Órdenes de Trabajo. "
                "Ejecútelo desde la pantalla inicial; el agente solo lee resultados y no ejecuta cálculos.",
                sin_datos=True,
            )
        metodo = getattr(self, f"r{pregunta.codigo}")
        respuesta: Respuesta = metodo()
        if not respuesta.periodo:
            respuesta.periodo = self.periodo_texto()
        return respuesta

    # ------------------------------------------------------------------
    # utilidades
    # ------------------------------------------------------------------
    def ots(self) -> dict[str, dict]:
        if self._ots_cache is None:
            objetos = {r["objeto_costo_id"]: r for r in self.L.tabla("objetos_costo")}
            salida = {}
            for ot in self.L.tabla("ordenes_trabajo", order="ot_id"):
                fila = dict(ot)
                obj = objetos.get(ot.get("objeto_costo_id"), {})
                fila["codigo_producto"] = obj.get("codigo_producto") or ot.get("objeto_costo_id")
                fila["descripcion_producto"] = obj.get("descripcion_producto") or ot.get("descripcion_ot")
                salida[ot["ot_id"]] = fila
            self._ots_cache = salida
        return self._ots_cache

    def etiqueta_ot(self, ot_id: str | None, row: dict | None = None) -> str:
        ot = self.ots().get(ot_id or "", {})
        producto = ot.get("codigo_producto") or (row or {}).get("codigo_producto") or ot_id
        codigo = ot.get("codigo_ot") or (row or {}).get("codigo_ot") or ot_id
        return f"{codigo} ({producto})" if producto and producto != codigo else str(codigo)

    def periodo_texto(self) -> str:
        periodo_id = None
        hoja = self.L.calculado("hoja_costos_ot")
        if hoja:
            periodo_id = hoja[0].get("periodo_id")
        filas = self.L.tabla("periodo", "periodo_id = ?", (periodo_id,)) if periodo_id else self.L.tabla("periodo", order="fecha_inicio")
        if not filas:
            return ""
        p = filas[0]
        rango = f", {p.get('fecha_inicio')} a {p.get('fecha_fin')}" if p.get("fecha_inicio") else ""
        return f"{p.get('periodo')} ({p.get('tipo_periodo')}{rango})"

    def fuente_calculo(self, tabla_bd: str, hojas: list[str]) -> str:
        hojas_txt = ", ".join(hojas)
        return f"Hojas {hojas_txt} del Excel maestro → tabla calculada {tabla_bd} (cálculo N.° {self.run})"

    def filtrar(self, filas: list[dict], campo: str, seleccion: list[str]) -> list[dict]:
        if not seleccion:
            return filas
        return [f for f in filas if f.get(campo) in seleccion]

    def ordenar(self, filas: list[dict], campo: str) -> list[dict]:
        return sorted(filas, key=lambda f: num(f.get(campo)), reverse=self.dir != "MIN")

    def ranking_ot(self, filas: list[dict], campo: str, nombre: str, fmt=soles, total_label: str = "del total",
                   frase_max: str | None = None, frase_min: str | None = None, con_participacion: bool = True,
                   columnas_extra: list[tuple[str, str]] | None = None) -> tuple[str, str]:
        """Devuelve (resumen, detalle) para una metrica por OT."""
        filas = [dict(f) for f in filas]
        total = sum(num(f.get(campo)) for f in filas)
        for f in filas:
            f["_objeto"] = self.etiqueta_ot(f.get("ot_id"), f)
            f["_valor"] = fmt(f.get(campo))
            f["_part"] = pct(num(f.get(campo)) / total) if total else "-"
        columnas = [("_objeto", "Objeto de costo (OT)"), ("_valor", nombre)] + (columnas_extra or [])
        if con_participacion:
            columnas.append(("_part", "Participación"))
        if self.ot_sel:
            elegidas = self.filtrar(filas, "ot_id", self.ot_sel)
            if not elegidas:
                return ("No encontré datos para la OT indicada.", "")
            partes = [f"{f['_objeto']}: {f['_valor']}" + (f" ({f['_part']} {total_label})" if con_participacion else "")
                      for f in elegidas]
            return (f"{nombre}: " + "; ".join(partes) + ".", tabla(self.ordenar(filas, campo), columnas))
        ordenadas = self.ordenar(filas, campo)
        tot = {"_objeto": "TOTAL", "_valor": fmt(total), "_part": pct(1) if total else "-"} if con_participacion else None
        detalle = tabla(ordenadas, columnas, tot)
        if self.dir and ordenadas:
            top = ordenadas[0]
            plantilla = frase_max if self.dir == "MAX" else frase_min
            if plantilla:
                texto = plantilla.format(objeto=top["_objeto"], valor=top["_valor"], part=top["_part"])
            else:
                criterio = "mayor" if self.dir == "MAX" else "menor"
                texto = f"El objeto de costo con {criterio} {nombre.lower()} es {top['_objeto']}: {top['_valor']}"
                if con_participacion:
                    texto += f" ({top['_part']} {total_label})"
                texto += "."
            return texto, detalle
        return (f"{nombre} por objeto de costo" + (f"; total {fmt(total)}." if con_participacion else "."), detalle)

    def ranking_actividad(self, filas: list[dict], campo: str, nombre: str) -> tuple[str, str]:
        filas = [dict(f) for f in filas]
        total = sum(num(f.get(campo)) for f in filas)
        for f in filas:
            f["_act"] = f"{f.get('codigo_actividad')} {f.get('nombre_actividad') or ''}".strip()
            f["_valor"] = soles(f.get(campo))
            f["_part"] = pct(num(f.get(campo)) / total) if total else "-"
        columnas = [("_act", "Actividad"), ("_valor", nombre), ("_part", "Participación")]
        ordenadas = self.ordenar(filas, campo) if self.dir else filas
        detalle = tabla(ordenadas, columnas, {"_act": "TOTAL", "_valor": soles(total), "_part": pct(1) if total else "-"})
        if self.act_sel:
            elegidas = self.filtrar(filas, "actividad_id", self.act_sel)
            resumen = "; ".join(f"{f['_act']}: {f['_valor']} ({f['_part']} del total)" for f in elegidas)
            return (f"{nombre}: {resumen}." if elegidas else "No encontré datos para la actividad indicada.", detalle)
        if self.dir and ordenadas:
            top = ordenadas[0]
            criterio = "mayor" if self.dir == "MAX" else "menor"
            return (f"La actividad con {criterio} {nombre.lower()} es {top['_act']}: {top['_valor']} ({top['_part']} del total).",
                    detalle)
        return (f"{nombre} por actividad; total {soles(total)}.", detalle)

    def tema_importado(self) -> str | None:
        for concepto, tabla_bd in TEMAS_IMPORTADOS:
            if concepto in self.i.conceptos:
                return tabla_bd
        return None

    def tema_calculado(self) -> str | None:
        for concepto, tabla_bd in TEMAS_CALCULADOS:
            if concepto in self.i.conceptos:
                return tabla_bd
        return None

    def filas_de_tema(self, tabla_bd: str) -> list[dict]:
        filas = self.L.tabla(tabla_bd)
        if self.ot_sel and filas and "ot_id" in filas[0]:
            filas = self.filtrar(filas, "ot_id", self.ot_sel)
        if self.act_sel and filas:
            campo = "actividad_id" if "actividad_id" in filas[0] else "actividad_id_destino" if "actividad_id_destino" in filas[0] else None
            if campo:
                filas = self.filtrar(filas, campo, self.act_sel)
        if self.rec_sel and filas and "elemento_costo_id" in filas[0]:
            filas = self.filtrar(filas, "elemento_costo_id", self.rec_sel)
        return filas

    @staticmethod
    def clave_de_fila(fila: dict) -> str:
        for campo in ("venta_id", "material_id", "mod_input_id", "soporte_item_id", "servicio_id", "elemento_costo_id",
                      "actividad_id", "ot_id", "objeto_costo_id", "empresa_id", "periodo_id", "modelo_id"):
            if fila.get(campo):
                return str(fila[campo])
        return "-"

    def ranking_total(self, filas: list[dict], campo: str) -> float:
        return sum(num(f.get(campo)) for f in filas)

    # ------------------------------------------------------------------
    # Datos y trazabilidad
    # ------------------------------------------------------------------
    def r1(self) -> Respuesta:
        empresas = self.L.tabla("empresas")
        periodos = self.L.tabla("periodo", order="fecha_inicio")
        modelos = self.L.tabla("modelos_costeo")
        if not empresas:
            return Respuesta("Empresa", "No hay datos de empresa importados.", sin_datos=True)
        e = empresas[0]
        lineas = [f"Empresa: {e.get('nombre_empresa')} ({e.get('sector')}, {e.get('departamento')}, {e.get('pais')}).",
                  f"Actividad principal: {e.get('actividad_principal')}"]
        detalle = tabla(
            [{"id": p.get("periodo_id"), "p": p.get("periodo"), "t": p.get("tipo_periodo"),
              "r": f"{p.get('fecha_inicio')} a {p.get('fecha_fin')}", "o": p.get("origen_periodo") or "Excel maestro"} for p in periodos],
            [("id", "Periodo"), ("p", "Nombre"), ("t", "Tipo"), ("r", "Rango"), ("o", "Origen")],
        )
        modelo_txt = "; ".join(f"{m.get('modelo_id')} {m.get('nombre_modelo')} ({m.get('ruta_modelo')})" for m in modelos)
        lineas.append(f"Modelo de costeo: Modelo ABC por Órdenes de Trabajo — registrado como {modelo_txt}.")
        return Respuesta("Empresa, periodo y modelo", "\n".join(lineas), detalle,
                         fuente="Hojas Empresa, Periodo y Modelos_Costeo del Excel maestro",
                         observaciones=[f"Último cálculo completado: N.° {self.run}." if self.run else "Aún no hay cálculo completado."])

    def r2(self) -> Respuesta:
        filas = list(self.ots().values())
        detalle = tabla(
            [{"ot": f.get("ot_id"), "cod": f.get("codigo_ot"), "oc": f.get("objeto_costo_id"), "prod": f.get("codigo_producto"),
              "desc": f.get("descripcion_producto"), "per": f.get("periodo_id")} for f in filas],
            [("cod", "Objeto de costo (OT)"), ("ot", "ID"), ("prod", "Producto que fabrica"), ("desc", "Descripción"),
             ("per", "Periodo")],
        )
        return Respuesta("Objetos de costo (OT)", f"Hay {len(filas)} objetos de costo: cada uno es una orden de trabajo, que fabrica un producto del caso.", detalle,
                         fuente="Hojas Objetos_Costo y Ordenes_Trabajo del Excel maestro")

    def r3(self) -> Respuesta:
        filas = [{"ot_id": k, **v} for k, v in self.ots().items()]
        resumen, detalle = self.ranking_ot(filas, "cantidad_producida", "Cantidad producida", fmt=lambda v: f"{cifra(v, 0)} pz",
                                           total_label="de la producción")
        return Respuesta("Unidades producidas", resumen, detalle, fuente="Hoja Ordenes_Trabajo (cantidad_producida)",
                         observaciones=[f"Unidad registrada: {next(iter(self.ots().values()), {}).get('unidad', '-')}."])

    def r4(self) -> Respuesta:
        filas = self.L.tabla("actividades", order="actividad_id")
        detalle = tabla(
            [{"c": f.get("codigo_actividad"), "n": f.get("nombre_actividad"), "p": f.get("proceso"), "cl": f.get("clasificacion"),
              "tn": cifra(f.get("tiempo_normal_min_kg"))} for f in filas],
            [("c", "Código"), ("n", "Actividad"), ("p", "Proceso"), ("cl", "Clasificación"), ("tn", "Tn (min/kg)")],
        )
        return Respuesta("Actividades", f"Intervienen {len(filas)} actividades productivas en el costeo.", detalle,
                         fuente="Hoja Actividades del Excel maestro",
                         observaciones=["Reciben MOD, otros recursos indirectos y soporte (mantenimiento, control de calidad, ingeniería)."])

    def r5(self) -> Respuesta:
        tabla_bd = self.tema_importado()
        calculada = self.tema_calculado()
        if tabla_bd is None and calculada:
            filas = self.filtrar(self.L.calculado(calculada), "ot_id", self.ot_sel)
            formula = filas[0].get("formula") if filas else ""
            return Respuesta("Fuente del dato",
                             f"Ese dato no se importa: es un cálculo generado por el sistema en la tabla {calculada}.",
                             formula=formula, fuente=filas[0].get("fuente") if filas else calculada)
        if tabla_bd is None:
            filas = [{"hoja": hoja, "tabla": t} for hoja, t in SHEET_TO_TABLE.items()]
            detalle = tabla(filas, [("hoja", "Hoja del Excel maestro"), ("tabla", "Tabla SQLite")])
            fuentes = self.L.tabla("fuentes", order="fuente_id")
            detalle += "\n\n" + tabla([{"c": f.get("cuadro_o_seccion"), "d": f.get("dato_usado")} for f in fuentes],
                                      [("c", "Cuadro / sección"), ("d", "Dato usado")])
            return Respuesta("Fuente del dato", "Indique el dato (p. ej. ventas, material directo, mantenimiento) para precisar su origen. "
                             "Este es el mapa de hojas y fuentes:", detalle)
        filas = self.filas_de_tema(tabla_bd)
        hoja = HOJA_DE_TABLA.get(tabla_bd, tabla_bd)
        if not filas:
            return Respuesta("Fuente del dato", f"El dato corresponde a la hoja {hoja} del Excel maestro, pero no tiene filas importadas.",
                             fuente=f"Hoja {hoja}")
        detalle = tabla(
            [{"k": self.clave_de_fila(f), "c": f.get("fuente_cuadro") or "-", "f": f.get("fuente") or "-"} for f in filas[:12]],
            [("k", "Registro"), ("c", "Cuadro"), ("f", "Fuente")],
        )
        traza = self.L.valor(
            "SELECT COUNT(*) FROM source_trace WHERE table_name = ? AND batch_id = (SELECT MAX(batch_id) FROM import_batches WHERE status = 'COMPLETADO')",
            (tabla_bd,), 0)
        return Respuesta("Fuente del dato", f"El dato proviene de la hoja «{hoja}» del Excel maestro (tabla {tabla_bd}).", detalle,
                         fuente=f"Hoja {hoja}", observaciones=[f"Filas trazadas en la última importación: {traza}."])

    def r6(self) -> Respuesta:
        tabla_bd = self.tema_importado()
        calculada = self.tema_calculado()
        objetivo = tabla_bd or calculada
        if objetivo is None:
            return Respuesta("Observaciones", "Indique el dato o cálculo (p. ej. ventas, CMD, MOD, mantenimiento) para mostrar sus observaciones.",
                             sin_datos=True)
        filas = self.filas_de_tema(objetivo) if tabla_bd else self.filtrar(self.L.calculado(objetivo), "ot_id", self.ot_sel)
        obs = [{"k": self.clave_de_fila(f) if tabla_bd else self.etiqueta_ot(f.get("ot_id")), "o": f.get("observacion") or "-"}
               for f in filas]
        unicas = {o["o"] for o in obs}
        resumen = (f"Observación de la hoja {HOJA_DE_TABLA.get(objetivo, objetivo)}: {unicas.pop()}" if len(unicas) == 1
                   else f"Observaciones registradas en {HOJA_DE_TABLA.get(objetivo, objetivo)}:")
        return Respuesta("Observaciones", resumen, "" if len({o['o'] for o in obs}) == 1 else tabla(obs, [("k", "Registro"), ("o", "Observación")]),
                         fuente=f"Columna observacion de {objetivo}")

    def r7(self) -> Respuesta:
        from ..importers import validate_workbook  # lectura del Excel en modo read_only

        lote = self.L.fila("SELECT * FROM import_batches ORDER BY batch_id DESC LIMIT 1") if self.L.existe_tabla("import_batches") else None
        rutas = [Path(lote["excel_path"])] if lote and lote.get("excel_path") else []
        rutas.append(DEFAULT_EXCEL_PATH)
        ruta = next((r for r in rutas if r.exists()), None)
        observaciones = []
        if lote:
            observaciones.append(f"Última importación: {lote.get('imported_at')} — {lote.get('status')} ({lote.get('notes')}).")
        if ruta is None:
            errores = self.L.filas("SELECT severity, sheet_name, message FROM import_errors WHERE batch_id = ?", (lote["batch_id"],)) if lote else []
            resumen = ("No encuentro el archivo Excel en esta PC; según el registro de la última importación " +
                       ("no hubo errores de estructura." if not errores else f"hubo {len(errores)} observaciones."))
            return Respuesta("Validación del Excel maestro", resumen,
                             tabla(errores, [("severity", "Tipo"), ("sheet_name", "Hoja"), ("message", "Mensaje")]) if errores else "",
                             observaciones=observaciones)
        resultado = validate_workbook(ruta)
        errores = [i for i in resultado.issues if i.severity == "ERROR"]
        advertencias = [i for i in resultado.issues if i.severity != "ERROR"]
        from ..constants import SHEET_SPECS

        resumen = (f"Sí: el Excel maestro tiene las {len(SHEET_SPECS)} hojas obligatorias y sus columnas." if resultado.ok
                   else f"No: el Excel maestro tiene {len(errores)} errores de estructura.")
        detalle = tabla([{"t": i.severity, "h": i.sheet_name or "-", "m": i.message} for i in resultado.issues],
                        [("t", "Tipo"), ("h", "Hoja"), ("m", "Mensaje")]) if resultado.issues else ""
        observaciones.insert(0, f"Archivo revisado (solo lectura): {ruta}")
        if advertencias:
            observaciones.append(f"Advertencias: {len(advertencias)}.")
        return Respuesta("Validación del Excel maestro", resumen, detalle, observaciones=observaciones)

    def r8(self) -> Respuesta:
        lote = self.L.fila("SELECT * FROM import_batches WHERE status = 'COMPLETADO' ORDER BY batch_id DESC LIMIT 1") \
            if self.L.existe_tabla("import_batches") else None
        registrados = self.L.filas("SELECT severity, sheet_name, message FROM import_errors WHERE batch_id = ?",
                                   (lote["batch_id"],)) if lote else []
        revisiones = []
        for hoja, grupo, campo, tabla_bd in [
            ("EC_AC_Inductores", "elemento_costo_id", "porcentaje_inductor", "ec_ac_inductores"),
            ("AC_OT_Pesos_Produccion", "actividad_id", "porcentaje_inductor", "ac_ot_pesos_produccion"),
            ("AC_OT_Pesos_Soporte", "actividad_id", "porcentaje_inductor", "ac_ot_pesos_soporte"),
            ("Soporte_CC_ING_Inductores", "proceso_soporte", "porcentaje_distribucion", "soporte_cc_ing_inductores"),
        ]:
            for row in self.L.filas(f"SELECT {grupo} AS g, SUM({campo}) AS s FROM {tabla_bd} GROUP BY {grupo}"):
                total = num(row.get("s"))
                estado = "OK" if abs(total - 1) <= 0.0001 else ("Advertencia" if abs(total - 1) > 0.02 else "Redondeo")
                revisiones.append({"h": hoja, "g": row.get("g"), "s": pct(total), "e": estado})
        relaciones = [r for r in registrados if "Relacion" in (r.get("message") or "")]
        no_ok = [r for r in revisiones if r["e"] != "OK"]
        resumen = (f"Relaciones: {'sin errores' if not relaciones else f'{len(relaciones)} errores'}. "
                   f"Porcentajes de inductores: {len(revisiones) - len(no_ok)} de {len(revisiones)} grupos suman 100 %"
                   + (f"; {len(no_ok)} con diferencia." if no_ok else "."))
        detalle = tabla(no_ok or revisiones, [("h", "Hoja"), ("g", "Grupo"), ("s", "Suma"), ("e", "Estado")])
        if registrados:
            detalle += "\n\nRegistrado en la importación:\n" + tabla(registrados, [("severity", "Tipo"), ("sheet_name", "Hoja"),
                                                                                   ("message", "Mensaje")])
        return Respuesta("Errores y advertencias", resumen, detalle,
                         fuente=f"import_errors del lote {lote['batch_id'] if lote else '-'} y revisión de sumas (solo lectura)")

    # ------------------------------------------------------------------
    # Material directo y personal (usa el kardex)
    # ------------------------------------------------------------------
    def r9(self) -> Respuesta:
        filas = self.L.calculado("cmd_ot")
        _kardexs, obs = kardex_desde_base(self.L)
        for f in filas:
            f["_kg"] = cifra(f.get("cantidad_kg"))
            f["_cu"] = f"S/ {num(f.get('costo_unitario_promedio')):,.4f}"
            f["_mat"] = f.get("material")
        resumen, detalle = self.ranking_ot(
            filas, "cmd_soles", "Material directo (CMD)",
            columnas_extra=[("_mat", "Material"), ("_kg", "Cantidad (kg)"), ("_cu", "C.U. promedio kardex")],
            frase_max="{objeto} consumió más material directo: {valor} ({part} del CMD total).",
            frase_min="{objeto} consumió menos material directo: {valor} ({part} del CMD total).")
        return Respuesta("Material directo consumido", resumen, detalle,
                         formula="CMD_OT = Σ salidas del kardex a la OT × costo promedio ponderado vigente",
                         fuente=self.fuente_calculo("cmd_ot", ["Compras"]), observaciones=obs)

    def r10(self) -> Respuesta:
        cmd = self.filtrar(self.L.calculado("cmd_ot"), "ot_id", self.ot_sel)
        kardexs, obs = kardex_desde_base(self.L)
        compras = {k.material_id: k for k in kardexs}
        filas = []
        for f in cmd:
            k = compras.get(f.get("material_id"))
            entradas = [x for x in k.filas if x.tipo == "ENTRADA"] if k else []
            q_ent = sum(x.ent_cantidad for x in entradas)
            cu_ent = sum(x.ent_total for x in entradas) / q_ent if q_ent else 0.0
            filas.append({"m": f.get("material"), "ot": self.etiqueta_ot(f.get("ot_id")), "kg": cifra(f.get("cantidad_kg")),
                          "cu": f"S/ {cu_ent:,.4f}" if q_ent else "-", "pp": f"S/ {num(f.get('costo_unitario_promedio')):,.4f}",
                          "t": soles(f.get("cmd_soles"))})
        return Respuesta("Cantidad y costo unitario por material", f"Se registran {len(filas)} materiales directos consumidos.",
                         tabla(filas, [("m", "Material"), ("ot", "OT"), ("kg", "Cantidad consumida (kg)"),
                                       ("cu", "C.U. compras netas"), ("pp", "C.U. promedio de salida"), ("t", "CMD")]),
                         formula="CMD = cantidad consumida × costo unitario promedio ponderado del kardex",
                         fuente="Hoja Compras (kardex de materiales directos) y tabla cmd_ot", observaciones=obs)

    def r11(self) -> Respuesta:
        filas = self.L.calculado("hoja_costos_ot")
        for f in filas:
            f["_part_ot"] = pct(num(f.get("cmd_ot")) / num(f.get("costo_total_ot"))) if num(f.get("costo_total_ot")) else "-"
        resumen, detalle = self.ranking_ot(filas, "cmd_ot", "CMD total", columnas_extra=[("_part_ot", "% del costo de la OT")],
                                           total_label="del CMD total")
        return Respuesta("CMD total", resumen, detalle, formula="CMD_OT = Σ cantidad × costo unitario promedio",
                         fuente=self.fuente_calculo("hoja_costos_ot", ["Compras"]))

    def r12(self) -> Respuesta:
        kardexs, obs = kardex_desde_base(self.L)
        if self.ot_sel:
            mats = {m.get("material_id") for m in self.L.calculado("cmd_ot") if m.get("ot_id") in self.ot_sel}
            kardexs = [k for k in kardexs if k.material_id in mats]
        filas = []
        for k in kardexs:
            q, t, destinos = salidas_por_material(k)
            filas.append({"m": k.material, "cu": f"S/ {t / q:,.4f}" if q else "-", "q": cifra(q) if q else "-",
                          "t": soles(t) if q else "-", "ref": ", ".join(destinos) or "-"})
        resumen = "Costo unitario promedio de las salidas del kardex: " + "; ".join(f"{f['m']} {f['cu']}" for f in filas) + "."
        if not obs or all("difiere" not in o for o in obs):
            obs = ["Conciliación kardex recalculado vs. CMD guardado: diferencia S/ 0.00."] + obs
        return Respuesta("Costo promedio del kardex", resumen,
                         tabla(filas, [("m", "Material"), ("cu", "C.U. promedio de salida"), ("q", "Salidas (kg)"),
                                       ("t", "Costo de salidas"), ("ref", "Destino")]),
                         formula="Costo promedio = valor del saldo disponible / cantidad disponible; salida = cantidad × costo promedio",
                         fuente=self.fuente_calculo("kardex_promedio", ["Compras"]), observaciones=obs)

    def r13(self) -> Respuesta:
        servicios = self.filtrar(self.L.tabla("servicios_directos"), "ot_id", self.ot_sel)
        csd = self.filtrar(self.L.calculado("csd_ot"), "ot_id", self.ot_sel)
        total = sum(num(f.get("csd_soles")) for f in csd)
        if not servicios:
            motivo = (csd[0].get("observacion") if csd else "") or "la hoja Servicios_Directos no trae registros"
            return Respuesta("Servicios directos (CSD)",
                             f"No hay servicios directos asignados: el CSD figura en {soles(total)} porque {motivo[0].lower() + motivo[1:]}",
                             formula="CSD_OT = 0 cuando no existen servicios directos en el Excel maestro",
                             fuente="Hoja Servicios_Directos (vacía) → tabla csd_ot",
                             observaciones=["El sistema no inventa servicios directos: si existieran, se registran en la hoja Servicios_Directos."])
        filas = [{"ot": self.etiqueta_ot(f.get("ot_id")), "s": f.get("servicio_directo"), "q": cifra(f.get("cantidad")),
                  "t": soles(f.get("csd_soles"))} for f in csd]
        return Respuesta("Servicios directos (CSD)", f"CSD total: {soles(total)}.",
                         tabla(filas, [("ot", "OT"), ("s", "Servicio"), ("q", "Cantidad"), ("t", "CSD")]),
                         formula="CSD_OT = cantidad × costo unitario", fuente="Hoja Servicios_Directos → tabla csd_ot")

    def _personal(self) -> list[dict]:
        return self.L.tabla("Personal_directos", order="actividad_id")

    def r14(self) -> Respuesta:
        pd = self.filtrar(self._personal(), "actividad_id", self.act_sel)
        filas = []
        for f in pd:
            ts = sum(num(f.get(c)) for c in ("tolerancia_personal_pct", "tolerancia_fatiga_pct", "tolerancia_pie_pct",
                                              "tolerancia_encorvado_pct"))
            dec = int(num(f.get("redondear_te_decimales")) or 2)
            te = round(num(f.get("tiempo_normal_min_kg")) * (1 + ts), dec)
            filas.append({"a": f"{f.get('codigo_actividad')} {f.get('nombre_actividad')}", "tn": cifra(f.get("tiempo_normal_min_kg")),
                          "ts": pct(ts, 0), "te": cifra(te, dec)})
        return Respuesta("Tiempo estándar", "Tiempo estándar (Te) por actividad, en minutos por kg:",
                         tabla(filas, [("a", "Actividad"), ("tn", "Tn (min/kg)"), ("ts", "Ts total"), ("te", "Te (min/kg)")]),
                         formula="Ts = Σ tolerancias (personal, fatiga, de pie, encorvado); Te = Tn × (1 + Ts)",
                         fuente="Hoja Personal_directos (Cuadros 2 y 3)")

    def r15(self) -> Respuesta:
        pd = self._personal()
        mod = self.L.calculado("mod_actividad")
        if not pd:
            return Respuesta("Factor de distribución", "No hay datos de personal directo importados.", sin_datos=True)
        horas_reales = num(pd[0].get("horas_reales_trabajadas_periodo"))
        horas_std = num(pd[0].get("horas_estandar_totales_periodo"))
        fd = horas_reales / horas_std if horas_std else 0
        fd_usado = num(mod[0].get("fd")) if mod else round(fd, 2)
        lectura = ("las horas reales fueron menores que las estándar" if fd < 1 else
                   "las horas reales superaron a las estándar" if fd > 1 else "las horas reales igualan a las estándar")
        return Respuesta("Factor de distribución (Fd)",
                         f"Fd = {cifra(horas_reales)} h reales / {cifra(horas_std)} h estándar = {fd:,.4f} → se usa {fd_usado:,.2f} ({lectura}).",
                         formula="Fd = horas reales trabajadas del periodo / horas estándar totales del periodo",
                         fuente="Hoja Personal_directos → tabla mod_actividad",
                         observaciones=["El Fd ajusta el tiempo estándar al tiempo real antes de repartir la MOD."])

    def r16(self) -> Respuesta:
        mod = self.filtrar(self.L.calculado("mod_actividad"), "actividad_id", self.act_sel)
        filas = [{"a": f"{f.get('codigo_actividad')} {f.get('nombre_actividad')}", "tr": cifra(f.get("tiempo_real_minutos"), 0),
                  "pc": pct(f.get("porcentaje_calculado")), "pu": pct(f.get("porcentaje_usado"), 0)} for f in mod]
        return Respuesta("Porcentaje de MOD por actividad", "Participación de cada actividad en la MOD:",
                         tabla(filas, [("a", "Actividad"), ("tr", "Tiempo real (min)"), ("pc", "% calculado"), ("pu", "% usado")]),
                         formula="% AC = tiempo real AC / Σ tiempo real; tiempo real = Te × kg producidos × Fd",
                         fuente=self.fuente_calculo("mod_actividad", ["Personal_directos"]),
                         observaciones=["El % usado se redondea a entero para replicar el cuadro del caso."])

    def r17(self) -> Respuesta:
        mod = self.L.calculado("mod_actividad")
        resumen, detalle = self.ranking_actividad(mod, "costo_mod_ac", "Costo MOD")
        total_mod = num(mod[0].get("total_mod_soles")) if mod else 0
        return Respuesta("Costo de MOD por actividad", resumen, detalle,
                         formula="Costo MOD AC = % AC × total MOD; total MOD = operarios × salario mensual",
                         fuente=self.fuente_calculo("mod_actividad", ["Personal_directos"]),
                         observaciones=[f"Total MOD del periodo: {soles(total_mod)}."])

    # ------------------------------------------------------------------
    # Recursos indirectos y soporte
    # ------------------------------------------------------------------
    def r18(self) -> Respuesta:
        rec = self.filtrar(self.L.tabla("costo_otros_recursos(CI)", order="elemento_costo_id"), "elemento_costo_id", self.rec_sel)
        total = sum(num(f.get("importe_soles")) for f in rec)
        filas = [{"r": f.get("nombre_recurso"), "i": soles(f.get("importe_soles")), "ind": f.get("inductor_recurso"),
                  "p": pct(num(f.get("importe_soles")) / total) if total else "-", "_v": num(f.get("importe_soles"))} for f in rec]
        if self.dir:
            filas.sort(key=lambda f: f["_v"], reverse=self.dir == "MAX")
        resumen = f"Se importaron {len(filas)} recursos indirectos por {soles(total)}."
        if self.dir and filas:
            resumen = f"El recurso indirecto con {'mayor' if self.dir == 'MAX' else 'menor'} importe es {filas[0]['r']}: {filas[0]['i']} ({filas[0]['p']})."
        return Respuesta("Recursos indirectos", resumen,
                         tabla(filas, [("r", "Recurso"), ("i", "Importe"), ("ind", "Inductor"), ("p", "Participación")],
                               {"r": "TOTAL", "i": soles(total), "p": pct(1) if total else "-"}),
                         fuente="Hoja costo_otros_recursos(CI) (Cuadros 4 y 5)")

    def r19(self) -> Respuesta:
        filas = self.filtrar(self.filtrar(self.L.tabla("ec_ac_inductores"), "actividad_id", self.act_sel), "elemento_costo_id", self.rec_sel)
        nombres = {r["elemento_costo_id"]: r.get("nombre_recurso") for r in self.L.tabla("costo_otros_recursos(CI)")}
        datos = [{"r": nombres.get(f.get("elemento_costo_id"), f.get("codigo_recurso")), "a": f"{f.get('codigo_actividad')} {f.get('nombre_actividad')}",
                  "i": f.get("inductor_recurso"), "p": pct(f.get("porcentaje_inductor"), 0)} for f in filas]
        alcance = "la actividad indicada" if self.act_sel else "cada actividad"
        return Respuesta("Porcentaje de recursos por actividad", f"Porcentaje de cada recurso asignado a {alcance}:",
                         tabla(datos, [("r", "Recurso"), ("a", "Actividad"), ("i", "Inductor"), ("p", "% asignado")]),
                         fuente="Hoja EC_AC_Inductores (Cuadro 6)")

    def r20(self) -> Respuesta:
        filas = self.filtrar(self.filtrar(self.L.calculado("otros_recursos_ac_detalle"), "actividad_id", self.act_sel),
                             "elemento_costo_id", self.rec_sel)
        total = sum(num(f.get("costo_recurso_ac")) for f in filas)
        datos = [{"r": f.get("nombre_recurso"), "a": f"{f.get('codigo_actividad')} {f.get('nombre_actividad')}",
                  "i": soles(f.get("importe_recurso_soles")), "p": pct(f.get("porcentaje_inductor"), 0), "c": soles(f.get("costo_recurso_ac"))}
                 for f in filas]
        if self.act_sel and self.rec_sel and len(filas) == 1:
            f = filas[0]
            resumen = (f"{f.get('nombre_actividad')} recibió {soles(f.get('costo_recurso_ac'))} de {f.get('nombre_recurso')} "
                       f"({soles(f.get('importe_recurso_soles'))} × {pct(f.get('porcentaje_inductor'), 0)}).")
        else:
            resumen = f"Costo indirecto asignado: {soles(total)} en {len(filas)} combinaciones recurso–actividad."
        return Respuesta("Costo indirecto recibido por actividad", resumen,
                         tabla(datos, [("r", "Recurso"), ("a", "Actividad"), ("i", "Importe recurso"), ("p", "%"), ("c", "Costo asignado")],
                               {"r": "TOTAL", "c": soles(total)}),
                         formula="Costo recurso AC = importe del recurso × % del inductor EC-AC",
                         fuente=self.fuente_calculo("otros_recursos_ac_detalle", ["costo_otros_recursos(CI)", "EC_AC_Inductores"]))

    def r21(self) -> Respuesta:
        resumen, detalle = self.ranking_actividad(self.L.calculado("otros_recursos_actividad"), "costo_otros_recursos_ac",
                                                  "Otros recursos indirectos")
        return Respuesta("Otros recursos por actividad", resumen, detalle, formula="Otros recursos AC = Σ costo recurso AC",
                         fuente=self.fuente_calculo("otros_recursos_actividad", ["costo_otros_recursos(CI)", "EC_AC_Inductores"]))

    def r22(self) -> Respuesta:
        filas = self.L.calculado("soporte_mantenimiento_ac")
        nombres = {a["actividad_id"]: a.get("nombre_actividad") for a in self.L.tabla("actividades")}
        for f in filas:
            f["nombre_actividad"] = nombres.get(f.get("actividad_id"), "")
        resumen, detalle = self.ranking_actividad(filas, "costo_mantenimiento_ac", "Mantenimiento")
        equipos = self.filtrar(self.L.tabla("Soporte_Mantenimiento(CI)"), "actividad_id_destino", self.act_sel)
        detalle += "\n\nEquipos:\n" + tabla(
            [{"e": f.get("equipo_descripcion"), "a": f.get("codigo_actividad_destino"), "c": soles(f.get("costo_equipo_soles"))} for f in equipos],
            [("e", "Equipo"), ("a", "Actividad"), ("c", "Costo")])
        return Respuesta("Mantenimiento por actividad", resumen, detalle,
                         formula="Mantenimiento AC = Σ costo de los equipos asociados a la actividad",
                         fuente=self.fuente_calculo("soporte_mantenimiento_ac", ["Soporte_Mantenimiento(CI)"]))

    def _soporte_distribuido(self, tabla_bd: str, campo_total: str, campo_ac: str, fuente_tabla: str, titulo: str,
                             detalle_campo: str) -> Respuesta:
        filas = self.L.calculado(tabla_bd)
        total = num(filas[0].get(campo_total)) if filas else 0
        elegidas = self.filtrar(filas, "actividad_id", self.act_sel)
        datos = [{"a": f"{f.get('codigo_actividad')} {f.get('nombre_actividad')}", "p": pct(f.get("porcentaje_distribucion")),
                  "c": soles(f.get(campo_ac))} for f in elegidas]
        items = self.L.tabla(fuente_tabla)
        origen = tabla([{"d": i.get(detalle_campo), "q": cifra(i.get("cantidad"), 0), "u": soles(i.get("costo_unitario_soles")),
                         "t": soles(num(i.get("cantidad")) * num(i.get("costo_unitario_soles")))} for i in items],
                       [("d", "Concepto"), ("q", "Cantidad"), ("u", "C.U."), ("t", "Importe")],
                       {"d": "TOTAL", "t": soles(total)})
        return Respuesta(titulo, f"Costo total de {titulo.lower()}: {soles(total)}, distribuido a las actividades así:",
                         tabla(datos, [("a", "Actividad"), ("p", "% de distribución"), ("c", "Costo asignado")]) + "\n\nOrigen:\n" + origen,
                         formula="Total = Σ cantidad × costo unitario; costo AC = total × % del cuadro 12",
                         fuente=self.fuente_calculo(tabla_bd, [HOJA_DE_TABLA.get(fuente_tabla, fuente_tabla), "Soporte_CC_ING_Inductores"]))

    def r23(self) -> Respuesta:
        return self._soporte_distribuido("soporte_control_calidad_ac", "costo_total_control_calidad", "costo_control_calidad_ac",
                                         "Soporte_Control_Calidad(CI)", "Control de calidad", "prueba_laboratorio")

    def r24(self) -> Respuesta:
        return self._soporte_distribuido("soporte_ingenieria_ac", "costo_total_ingenieria", "costo_ingenieria_ac",
                                         "Soporte_Ingenieria(CI)", "Ingeniería y desarrollo", "tipo_plano")

    def r25(self) -> Respuesta:
        filas = self.L.calculado("soporte_actividad")
        resumen, _ = self.ranking_actividad(filas, "costo_soporte_ac", "Soporte total")
        if self.dir:
            filas = self.ordenar(filas, "costo_soporte_ac")
        datos = [{"a": f"{f.get('codigo_actividad')} {f.get('nombre_actividad')}", "m": soles(f.get("costo_mantenimiento_ac")),
                  "cc": soles(f.get("costo_control_calidad_ac")), "i": soles(f.get("costo_ingenieria_ac")), "t": soles(f.get("costo_soporte_ac"))}
                 for f in self.filtrar(filas, "actividad_id", self.act_sel)]
        tot = {"a": "TOTAL", **{k: soles(self.ranking_total(filas, c)) for k, c in
                                [("m", "costo_mantenimiento_ac"), ("cc", "costo_control_calidad_ac"), ("i", "costo_ingenieria_ac"),
                                 ("t", "costo_soporte_ac")]}}
        return Respuesta("Soporte por actividad", resumen,
                         tabla(datos, [("a", "Actividad"), ("m", "Mantenimiento"), ("cc", "Control calidad"), ("i", "Ingeniería"),
                                       ("t", "Soporte total")], None if self.act_sel else tot),
                         formula="Soporte AC = mantenimiento AC + control de calidad AC + ingeniería AC",
                         fuente=self.fuente_calculo("soporte_actividad", ["Soporte_Mantenimiento(CI)", "Soporte_Control_Calidad(CI)",
                                                                          "Soporte_Ingenieria(CI)"]))

    def r26(self) -> Respuesta:
        filas = self.filtrar(self.L.calculado("actividad_total"), "actividad_id", self.act_sel)
        datos = []
        for f in filas:
            t = num(f.get("costo_total_ac"))
            datos.append({"a": f"{f.get('codigo_actividad')} {f.get('nombre_actividad')}", "mod": soles(f.get("costo_mod_ac")),
                          "or": soles(f.get("costo_otros_recursos_ac")), "s": soles(f.get("costo_soporte_ac")), "t": soles(t),
                          "pm": pct(num(f.get("costo_mod_ac")) / t) if t else "-"})
        if len(filas) == 1:
            f = filas[0]
            t = num(f.get("costo_total_ac"))
            resumen = (f"{f.get('nombre_actividad')} cuesta {soles(t)}: MOD {soles(f.get('costo_mod_ac'))} ({pct(num(f.get('costo_mod_ac')) / t)}), "
                       f"otros recursos {soles(f.get('costo_otros_recursos_ac'))} ({pct(num(f.get('costo_otros_recursos_ac')) / t)}) y "
                       f"soporte {soles(f.get('costo_soporte_ac'))} ({pct(num(f.get('costo_soporte_ac')) / t)}).")
        else:
            resumen = f"Costo total de las actividades: {soles(self.ranking_total(filas, 'costo_total_ac'))}."
        return Respuesta("Costo total por actividad", resumen,
                         tabla(datos, [("a", "Actividad"), ("mod", "MOD"), ("or", "Otros recursos"), ("s", "Soporte"), ("t", "Total"),
                                       ("pm", "% MOD")]),
                         formula="Costo total AC = MOD AC + otros recursos AC + soporte AC",
                         fuente=self.fuente_calculo("actividad_total", ["Personal_directos", "costo_otros_recursos(CI)", "Soporte_*"]))

    def r27(self) -> Respuesta:
        if self.dir is None:
            self.dir = "MAX"
        resumen, detalle = self.ranking_actividad(self.L.calculado("actividad_total"), "costo_total_ac", "Costo total")
        return Respuesta("Ranking de actividades", resumen, detalle,
                         formula="Costo total AC = MOD AC + otros recursos AC + soporte AC",
                         fuente=self.fuente_calculo("actividad_total", ["Personal_directos", "costo_otros_recursos(CI)", "Soporte_*"]))

    # ------------------------------------------------------------------
    # Asignacion a ordenes y hoja de costos
    # ------------------------------------------------------------------
    def r28(self) -> Respuesta:
        prod = self.filtrar(self.filtrar(self.L.tabla("ac_ot_pesos_produccion"), "actividad_id", self.act_sel), "ot_id", self.ot_sel)
        sop = {(f.get("actividad_id"), f.get("ot_id")): f for f in self.L.tabla("ac_ot_pesos_soporte")}
        datos = [{"a": f"{f.get('codigo_actividad')} {f.get('nombre_actividad')}", "ot": self.etiqueta_ot(f.get("ot_id")),
                  "pp": pct(f.get("porcentaje_inductor")),
                  "ps": pct(sop.get((f.get("actividad_id"), f.get("ot_id")), {}).get("porcentaje_inductor"))} for f in prod]
        resumen = (f"Peso de {datos[0]['a']} a {datos[0]['ot']}: producción {datos[0]['pp']}, soporte {datos[0]['ps']}."
                   if len(datos) == 1 else "Pesos de distribución actividad → OT:")
        return Respuesta("Pesos AC → OT", resumen,
                         tabla(datos, [("a", "Actividad"), ("ot", "OT"), ("pp", "Peso producción"), ("ps", "Peso soporte")]),
                         fuente="Hojas AC_OT_Pesos_Produccion y AC_OT_Pesos_Soporte (Cuadros 15 y 16)")

    def _transferencia(self, tabla_bd: str, campo_ac: str, campo_ot: str, titulo: str, hojas: list[str]) -> Respuesta:
        filas = self.filtrar(self.filtrar(self.L.calculado(tabla_bd), "actividad_id", self.act_sel), "ot_id", self.ot_sel)
        total = sum(num(f.get(campo_ot)) for f in filas)
        datos = [{"a": f"{f.get('codigo_actividad')} {f.get('nombre_actividad')}", "ot": self.etiqueta_ot(f.get("ot_id")),
                  "c": soles(f.get(campo_ac)), "p": pct(f.get("peso_ac_ot")), "t": soles(f.get(campo_ot))} for f in filas]
        if len(filas) == 1:
            f = filas[0]
            resumen = (f"{f.get('nombre_actividad')} transfirió {soles(f.get(campo_ot))} a {self.etiqueta_ot(f.get('ot_id'))} "
                       f"({soles(f.get(campo_ac))} × {pct(f.get('peso_ac_ot'))}).")
        else:
            resumen = f"{titulo} transferido: {soles(total)}."
        return Respuesta(titulo, resumen,
                         tabla(datos, [("a", "Actividad"), ("ot", "OT"), ("c", "Costo de la actividad"), ("p", "Peso"),
                                       ("t", "Transferido")], {"a": "TOTAL", "t": soles(total)}),
                         formula=f"Transferido = costo de la actividad × peso AC-OT",
                         fuente=self.fuente_calculo(tabla_bd, hojas))

    def r29(self) -> Respuesta:
        return self._transferencia("produccion_ot_detalle", "costo_conversion_ac", "costo_conversion_ot",
                                   "Costo de producción", ["AC_OT_Pesos_Produccion"])

    def r30(self) -> Respuesta:
        return self._transferencia("soporte_ot_detalle", "costo_soporte_ac", "costo_soporte_ot", "Costo de soporte",
                                   ["AC_OT_Pesos_Soporte"])

    def r31(self) -> Respuesta:
        filas = self.filtrar(self.L.calculado("hoja_costos_ot"), "ot_id", self.ot_sel)
        datos = [{"ot": self.etiqueta_ot(f.get("ot_id")), "mod": soles(f.get("mod_ot")), "or": soles(f.get("otros_recursos_ot")),
                  "s": soles(f.get("costo_soporte_ot")),
                  "t": soles(num(f.get("mod_ot")) + num(f.get("otros_recursos_ot")) + num(f.get("costo_soporte_ot")))} for f in filas]
        if len(filas) == 1:
            f = filas[0]
            resumen = (f"{self.etiqueta_ot(f.get('ot_id'))} recibió MOD {soles(f.get('mod_ot'))}, recursos indirectos "
                       f"{soles(f.get('otros_recursos_ot'))} y soporte {soles(f.get('costo_soporte_ot'))}.")
        else:
            resumen = "MOD, recursos indirectos y soporte recibidos por cada OT:"
        return Respuesta("Costos recibidos por la OT", resumen,
                         tabla(datos, [("ot", "OT"), ("mod", "MOD"), ("or", "Recursos indirectos"), ("s", "Soporte"), ("t", "Total")]),
                         fuente=self.fuente_calculo("hoja_costos_ot", ["AC_OT_Pesos_Produccion", "AC_OT_Pesos_Soporte"]))

    def r32(self) -> Respuesta:
        resumen, detalle = self.ranking_ot(self.L.calculado("hoja_costos_ot"), "costo_total_ot", "Costo total",
                                           frase_max="El objeto de costo con mayor costo total es {objeto}: {valor} ({part} del total).",
                                           frase_min="El objeto de costo con menor costo total es {objeto}: {valor} ({part} del total).")
        return Respuesta("Costo total de la OT", resumen, detalle,
                         formula="Costo total OT = CMD + costo de producción + costo de soporte + CSD",
                         fuente=self.fuente_calculo("hoja_costos_ot", ["Compras", "Personal_directos", "costo_otros_recursos(CI)", "Soporte_*"]))

    def r33(self) -> Respuesta:
        filas = self.L.calculado("hoja_costos_ot")
        for f in filas:
            f["_q"] = cifra(f.get("cantidad_producida"), 0)
            f["_t"] = soles(f.get("costo_total_ot"))
        resumen, detalle = self.ranking_ot(filas, "costo_unitario_ot", "Costo unitario", fmt=lambda v: f"S/ {num(v):,.4f}",
                                           con_participacion=False, columnas_extra=[("_t", "Costo total"), ("_q", "Cantidad")])
        return Respuesta("Costo unitario", resumen, detalle, formula="Costo unitario = costo total OT / cantidad producida",
                         fuente=self.fuente_calculo("hoja_costos_ot", ["Ordenes_Trabajo"]))

    def _componentes(self, f: dict) -> list[tuple[str, float]]:
        return [("CMD", num(f.get("cmd_ot"))), ("Producción (MOD + otros recursos)", num(f.get("costo_produccion_ot"))),
                ("Soporte", num(f.get("costo_soporte_ot"))), ("CSD", num(f.get("csd_ot")))]

    def r34(self) -> Respuesta:
        filas = self.filtrar(self.L.calculado("hoja_costos_ot"), "ot_id", self.ot_sel)
        bloques = []
        for f in filas:
            t = num(f.get("costo_total_ot"))
            datos = [{"c": n, "v": soles(v), "p": pct(v / t) if t else "-"} for n, v in self._componentes(f)]
            bloques.append(f"{self.etiqueta_ot(f.get('ot_id'))}\n" + tabla(datos, [("c", "Componente"), ("v", "Importe"), ("p", "Participación")],
                                                                          {"c": "TOTAL", "v": soles(t), "p": pct(1) if t else "-"}))
        return Respuesta("Composición del costo", "Composición del costo entre CMD, producción, soporte y CSD:", "\n\n".join(bloques),
                         formula="Participación = componente / costo total OT",
                         fuente=self.fuente_calculo("hoja_costos_ot", ["Compras", "Personal_directos", "costo_otros_recursos(CI)", "Soporte_*", "AC_OT_Pesos_*"]))

    def r35(self) -> Respuesta:
        filas = self.filtrar(self.L.calculado("hoja_costos_ot"), "ot_id", self.ot_sel)
        datos, frases = [], []
        for f in filas:
            t = num(f.get("costo_total_ot"))
            comps = [c for c in self._componentes(f) if self.dir == "MAX" or self.dir is None or c[1] > 0]
            elegido = (min if self.dir == "MIN" else max)(comps, key=lambda c: c[1])
            datos.append({"ot": self.etiqueta_ot(f.get("ot_id")), "c": elegido[0], "v": soles(elegido[1]), "p": pct(elegido[1] / t) if t else "-"})
            frases.append(f"{self.etiqueta_ot(f.get('ot_id'))}: {elegido[0]} ({pct(elegido[1] / t) if t else '-'})")
        criterio = "menor" if self.dir == "MIN" else "mayor"
        return Respuesta("Componente dominante", f"Componente de {criterio} participación — " + "; ".join(frases) + ".",
                         tabla(datos, [("ot", "OT"), ("c", "Componente"), ("v", "Importe"), ("p", "Participación")]),
                         formula="Participación = componente / costo total OT", fuente=self.fuente_calculo("hoja_costos_ot", ["Compras", "Personal_directos", "costo_otros_recursos(CI)", "Soporte_*", "AC_OT_Pesos_*"]))

    def r36(self) -> Respuesta:
        filas = self.L.calculado("hoja_costos_ot")
        elegir = min if self.dir == "MIN" else max
        criterio = "menor" if self.dir == "MIN" else "mayor"
        t = elegir(filas, key=lambda f: num(f.get("costo_total_ot")))
        u = elegir(filas, key=lambda f: num(f.get("costo_unitario_ot")))
        resumen = (f"{criterio.capitalize()} costo total: {self.etiqueta_ot(t.get('ot_id'))} con {soles(t.get('costo_total_ot'))}. "
                   f"{criterio.capitalize()} costo unitario: {self.etiqueta_ot(u.get('ot_id'))} con S/ {num(u.get('costo_unitario_ot')):,.4f}.")
        datos = [{"ot": self.etiqueta_ot(f.get("ot_id")), "t": soles(f.get("costo_total_ot")), "u": f"S/ {num(f.get('costo_unitario_ot')):,.4f}"}
                 for f in self.ordenar(filas, "costo_total_ot")]
        return Respuesta("OT de mayor costo", resumen, tabla(datos, [("ot", "OT"), ("t", "Costo total"), ("u", "Costo unitario")]),
                         formula="Costo unitario = costo total / cantidad producida", fuente=self.fuente_calculo("hoja_costos_ot", ["Compras", "Personal_directos", "costo_otros_recursos(CI)", "Soporte_*", "AC_OT_Pesos_*"]),
                         observaciones=["Un costo total alto puede deberse al volumen; compare también el costo unitario."])

    # ------------------------------------------------------------------
    # Ventas y rentabilidad
    # ------------------------------------------------------------------
    def r37(self) -> Respuesta:
        filas = self.L.calculado("estado_resultados_ot")
        ventas = {v.get("ot_id"): v for v in self.L.tabla("ventas")}
        extra = []
        if any(v.get("cantidad_vendida") for v in ventas.values()):
            for f in filas:
                v = ventas.get(f.get("ot_id"), {})
                f["_q"] = cifra(v.get("cantidad_vendida"), 0) if v.get("cantidad_vendida") else "-"
                f["_vvu"] = soles(v.get("valor_venta_unitario_soles")) if v.get("valor_venta_unitario_soles") else "-"
            extra = [("_q", "Cantidad vendida"), ("_vvu", "Valor venta unit.")]
        resumen, detalle = self.ranking_ot(filas, "ventas_netas_soles", "Ventas netas (ingresos)", columnas_extra=extra,
                                           total_label="de las ventas netas",
                                           frase_max="El objeto de costo que genera más ingresos es {objeto}: {valor} ({part} de las ventas netas).",
                                           frase_min="El objeto de costo que genera menos ingresos es {objeto}: {valor} ({part} de las ventas netas).")
        return Respuesta("Ventas netas", resumen, detalle, fuente=self.fuente_calculo("estado_resultados_ot", ["Ventas (Cuadro 18)"]),
                         observaciones=["Las ventas netas se importan del Excel maestro; el sistema no las calcula."])

    def _estado(self, campo: str, nombre: str, formula: str, frase_max: str | None = None, frase_min: str | None = None,
                fmt=soles, con_part: bool = True) -> Respuesta:
        filas = self.L.calculado("estado_resultados_ot")
        for f in filas:
            f["_v"] = soles(f.get("ventas_netas_soles"))
        resumen, detalle = self.ranking_ot(filas, campo, nombre, fmt=fmt, con_participacion=con_part, frase_max=frase_max,
                                           frase_min=frase_min, columnas_extra=[("_v", "Ventas netas")])
        return Respuesta(nombre, resumen, detalle, formula=formula, fuente=self.fuente_calculo("estado_resultados_ot", ["Ventas", "hoja de costos"]))

    def r38(self) -> Respuesta:
        return self._estado("margen_bruto", "Margen bruto", "Margen bruto = ventas netas − materia prima directa − costos de conversión")

    def r39(self) -> Respuesta:
        return self._estado("margen_operativo", "Margen operativo", "Margen operativo = margen bruto − procesos de soporte")

    def r40(self) -> Respuesta:
        return self._estado("rentabilidad_operativa", "Rentabilidad operativa", "Rentabilidad operativa = margen operativo / ventas netas",
                            fmt=pct, con_part=False,
                            frase_max="La mayor rentabilidad operativa la tiene {objeto}: {valor}.",
                            frase_min="La menor rentabilidad operativa la tiene {objeto}: {valor}.")

    def r41(self) -> Respuesta:
        filas = self.L.calculado("estado_resultados_ot")
        alto = max(filas, key=lambda f: num(f.get("rentabilidad_operativa")))
        bajo = min(filas, key=lambda f: num(f.get("rentabilidad_operativa")))
        if self.dir == "MIN":
            resumen = f"El objeto de costo con menor rentabilidad operativa es {self.etiqueta_ot(bajo.get('ot_id'))}: {pct(bajo.get('rentabilidad_operativa'))}."
        elif self.dir == "MAX":
            resumen = f"El objeto de costo con mayor rentabilidad operativa es {self.etiqueta_ot(alto.get('ot_id'))}: {pct(alto.get('rentabilidad_operativa'))}."
        else:
            resumen = (f"Mayor: {self.etiqueta_ot(alto.get('ot_id'))} ({pct(alto.get('rentabilidad_operativa'))}). "
                       f"Menor: {self.etiqueta_ot(bajo.get('ot_id'))} ({pct(bajo.get('rentabilidad_operativa'))}).")
        self.dir = self.dir or "MAX"
        datos = [{"ot": self.etiqueta_ot(f.get("ot_id")), "v": soles(f.get("ventas_netas_soles")), "m": soles(f.get("margen_operativo")),
                  "r": pct(f.get("rentabilidad_operativa"))} for f in self.ordenar(filas, "rentabilidad_operativa")]
        return Respuesta("Rentabilidad por producto", resumen,
                         tabla(datos, [("ot", "Objeto de costo (OT)"), ("v", "Ventas netas"), ("m", "Margen operativo"), ("r", "Rentabilidad")]),
                         formula="Rentabilidad operativa = margen operativo / ventas netas",
                         fuente=self.fuente_calculo("estado_resultados_ot", ["Ventas", "hoja de costos"]))

    def r42(self) -> Respuesta:
        from .propuestas import priorizar_ots

        filas = priorizar_ots(self.L)
        if not filas:
            return Respuesta("Prioridad de revisión", "No hay resultados para priorizar.", sin_datos=True)
        top = filas[0]
        datos = [{"n": i + 1, "ot": self.etiqueta_ot(f["ot_id"]), "r": pct(f["rentabilidad"]), "cmd": pct(f["cmd_sobre_ventas"]),
                  "s": pct(f["soporte_sobre_costo"]), "m": "; ".join(f["motivos"]) or "-"} for i, f in enumerate(filas)]
        return Respuesta("Prioridad de revisión",
                         f"Conviene revisar primero {self.etiqueta_ot(top['ot_id'])}: " + "; ".join(top["motivos"]) + ".",
                         tabla(datos, [("n", "N.°"), ("ot", "OT"), ("r", "Rentabilidad"), ("cmd", "CMD / ventas"), ("s", "Soporte / costo"),
                                       ("m", "Motivos")]),
                         formula="Prioridad: menor rentabilidad operativa; desempate por mayor CMD/ventas y soporte/costo",
                         fuente="estado_resultados_ot y hoja_costos_ot (solo lectura)")

    def r43(self) -> Respuesta:
        from .propuestas import generar_propuestas, propuestas_como_texto

        alertas = self.L.calculado("dashboard_alertas")
        texto_alertas = "\n".join(
            f"- [{a.get('severidad')}] {a.get('mensaje')} ({a.get('observacion')})" for a in alertas) or "- Sin alertas registradas."
        props = generar_propuestas(self.L)
        return Respuesta("Alertas y propuestas de mejora", f"El dashboard registra {len(alertas)} alertas; el agente propone {len(props)} mejoras.",
                         "ALERTAS DEL DASHBOARD\n" + texto_alertas + "\n\nPROPUESTAS DE MEJORA (no se aplican: requieren decisión del área)\n"
                         + propuestas_como_texto(props),
                         fuente="dashboard_alertas, hoja_costos_ot, estado_resultados_ot, actividad_total y kardex (solo lectura)")

    # ------------------------------------------------------------------
    # Salidas del sistema
    # ------------------------------------------------------------------
    def _archivo(self, nombre: str) -> str:
        ruta = REPORTS_DIR / nombre
        if ruta.exists():
            from datetime import datetime

            return f"existe ({datetime.fromtimestamp(ruta.stat().st_mtime):%Y-%m-%d %H:%M})"
        return "aún no generado"

    def r44(self) -> Respuesta:
        detalle_rows = self.L.calculado("otros_recursos_ac_detalle")
        mod = self.L.calculado("mod_actividad")
        sop = self.L.calculado("soporte_actividad")
        total_rec = self.ranking_total(detalle_rows, "costo_recurso_ac")
        resumen = (f"El reporte de asignación de recursos a actividades (Proceso 1 EC → AC) muestra cómo llegan a cada actividad la MOD "
                   f"({soles(self.ranking_total(mod, 'costo_mod_ac'))}), los otros recursos indirectos ({soles(total_rec)}) y el soporte "
                   f"({soles(self.ranking_total(sop, 'costo_soporte_ac'))}), con fórmula, datos de entrada y resultado.")
        return Respuesta("Reporte Proceso 1 EC-AC", resumen,
                         f"Archivo Reporte_Proceso_1_EC_AC.xlsx: {self._archivo('Reporte_Proceso_1_EC_AC.xlsx')}.",
                         fuente="Pestaña Proceso 1 EC-AC y reporte Excel")

    def r45(self) -> Respuesta:
        hojas = self.filtrar(self.L.calculado("hoja_costos_ot"), "ot_id", self.ot_sel)
        estados = {f.get("ot_id"): f for f in self.L.calculado("estado_resultados_ot")}
        bloques = []
        for h in hojas:
            e = estados.get(h.get("ot_id"), {})
            filas = [
                ("Cantidad producida", f"{cifra(h.get('cantidad_producida'), 0)} pz"),
                ("CMD", soles(h.get("cmd_ot"))), ("MOD", soles(h.get("mod_ot"))),
                ("Otros recursos indirectos", soles(h.get("otros_recursos_ot"))),
                ("Costo de producción", soles(h.get("costo_produccion_ot"))), ("Soporte", soles(h.get("costo_soporte_ot"))),
                ("CSD", soles(h.get("csd_ot"))), ("COSTO TOTAL", soles(h.get("costo_total_ot"))),
                ("Costo unitario", f"S/ {num(h.get('costo_unitario_ot')):,.4f}"),
                ("", ""),
                ("Ventas netas", soles(e.get("ventas_netas_soles"))), ("(−) Materia prima directa", soles(e.get("materia_prima_directa"))),
                ("(−) Costos de conversión", soles(e.get("costos_conversion"))), ("MARGEN BRUTO", soles(e.get("margen_bruto"))),
                ("(−) Procesos de soporte", soles(e.get("procesos_soporte"))), ("MARGEN OPERATIVO", soles(e.get("margen_operativo"))),
                ("Rentabilidad operativa", pct(e.get("rentabilidad_operativa"))),
            ]
            bloques.append(f"{self.etiqueta_ot(h.get('ot_id'))}\n" + tabla([{"c": c, "v": v} for c, v in filas], [("c", "Concepto"), ("v", "Importe")]))
        return Respuesta("Hoja de costos y estado de resultados",
                         "Hoja de costos y estado de resultados" + (" de la OT indicada:" if self.ot_sel else " por OT:"),
                         "\n\n".join(bloques), fuente="hoja_costos_ot y estado_resultados_ot (solo lectura)",
                         observaciones=[f"Reporte_Hoja_Costos_OT.xlsx: {self._archivo('Reporte_Hoja_Costos_OT.xlsx')}; "
                                        f"Reporte_Estado_Resultados.xlsx: {self._archivo('Reporte_Estado_Resultados.xlsx')}."])

    def r46(self) -> Respuesta:
        hoja = self.L.calculado("hoja_costos_ot")
        estado = self.L.calculado("estado_resultados_ot")
        ventas = self.ranking_total(estado, "ventas_netas_soles")
        mo = self.ranking_total(estado, "margen_operativo")
        kpis = [("Costo total de OT", soles(self.ranking_total(hoja, "costo_total_ot"))), ("Ventas netas", soles(ventas)),
                ("Margen operativo", soles(mo)), ("Rentabilidad operativa global", pct(mo / ventas) if ventas else "-")]
        graficos = ["Costo total por OT (barras)", "Estructura de costo por OT: CMD, producción, soporte y CSD (barras apiladas)",
                    "Ranking de actividades por costo total (barras horizontales)", "Rentabilidad operativa por OT (%)",
                    "Panel de alertas y recomendaciones"]
        return Respuesta("Dashboard", "Indicadores (KPI) actuales: " + "; ".join(f"{k}: {v}" for k, v in kpis) + ".",
                         "Gráficos:\n" + "\n".join(f"  - {g}" for g in graficos),
                         fuente="Pestaña Dashboard (dashboard.py) — datos de la última ejecución" if self.run else "Pestaña Dashboard")

    def r47(self) -> Respuesta:
        filas = [{"t": "Excel", "n": n, "e": self._archivo(n)} for n in OUTPUT_REPORTS]
        filas += [{"t": "Word", "n": n, "e": self._archivo(n)} for n in OUTPUT_DOCUMENTS]
        return Respuesta("Reportes y documentos", f"El aplicativo genera {len(OUTPUT_REPORTS)} reportes Excel y {len(OUTPUT_DOCUMENTS)} documentos Word "
                         "(más el dashboard en PNG) desde la pestaña Reportes.",
                         tabla(filas, [("t", "Tipo"), ("n", "Archivo"), ("e", "Estado en la carpeta reports")]),
                         observaciones=["El agente no genera archivos: use los botones de la pestaña Reportes."])

    def r48(self) -> Respuesta:
        ruta = REPORTS_DIR / "Informe_pruebas.docx"
        if not ruta.exists():
            return Respuesta("Informe de pruebas", "Aún no hay un informe de pruebas generado. Ejecútelo desde la pestaña Reportes "
                             "(el agente no corre pruebas porque estas reimportan el Excel y recalculan).", sin_datos=True)
        try:
            from docx import Document

            doc = Document(str(ruta))
            resumen = next((p.text for p in doc.paragraphs if p.text.startswith("Pruebas ejecutadas")), "")
            filas = []
            for t in doc.tables:
                encabezado = [c.text.strip().lower() for c in t.rows[0].cells]
                if "prueba" in encabezado:
                    for row in t.rows[1:]:
                        celdas = dict(zip(encabezado, [c.text.strip() for c in row.cells]))
                        ok = celdas.get("ok", "")
                        filas.append({"p": celdas.get("prueba", ""), "r": "Conforme" if ok in {"True", "SI", "Sí", "OK"} else ok or "-",
                                      "d": celdas.get("detalle", "")})
        except Exception as exc:  # lectura de un archivo externo
            return Respuesta("Informe de pruebas", f"No pude leer el informe de pruebas: {exc}", sin_datos=True)
        return Respuesta("Informe de pruebas", resumen or f"Informe encontrado ({self._archivo('Informe_pruebas.docx')}).",
                         tabla(filas, [("p", "Prueba"), ("r", "Resultado"), ("d", "Detalle")]) if filas else "",
                         fuente=f"{ruta.name} {self._archivo('Informe_pruebas.docx')} (lectura)")

    # ------------------------------------------------------------------
    # Ampliaciones
    # ------------------------------------------------------------------
    def rE1(self) -> Respuesta:
        filas = self.L.calculado("hoja_costos_ot")
        estado = {f.get("ot_id"): f for f in self.L.calculado("estado_resultados_ot")}
        for f in filas:
            f["costo_ventas"] = num(f.get("costo_total_ot"))
            ventas = num(estado.get(f.get("ot_id"), {}).get("ventas_netas_soles"))
            f["_cv_v"] = pct(f["costo_ventas"] / ventas) if ventas else "-"
            f["_cmd"] = soles(f.get("cmd_ot"))
            f["_conv"] = soles(f.get("costo_produccion_ot"))
            f["_sop"] = soles(f.get("costo_soporte_ot"))
        resumen, detalle = self.ranking_ot(
            filas, "costo_ventas", "Costo de ventas", total_label="del costo de ventas",
            columnas_extra=[("_cmd", "CMD (kardex)"), ("_conv", "Conversión"), ("_sop", "Soporte"), ("_cv_v", "% de las ventas")],
            frase_max="El objeto de costo que generó más costo de ventas es {objeto}: {valor} ({part} del costo de ventas).",
            frase_min="El objeto de costo que generó menos costo de ventas es {objeto}: {valor} ({part} del costo de ventas).")
        return Respuesta("Costo de ventas", resumen, detalle,
                         formula="Costo de ventas = inventario inicial de PT + costo de producción − inventario final de PT; "
                                 "con II = IF = 0, costo de ventas = costo total de la OT (CMD + conversión + soporte + CSD)",
                         fuente=self.fuente_calculo("hoja_costos_ot", ["Compras", "Ventas"]),
                         observaciones=["El Excel maestro no registra inventarios de productos terminados: se asume que toda la producción "
                                        "de la OT se vendió en el periodo.",
                                        "En el estado de resultados del sistema el soporte se presenta debajo del margen bruto."])

    def _kardex_compras(self) -> str:
        from ..kardex_compras import conciliacion_cmd, kardex_desde_base as kardex_compras

        if not self.L.existe_tabla("compras"):
            return ""
        conn = self.L.conectar()
        items = kardex_compras(conn)
        if not items:
            return ""
        conc = {c["material_id"]: c for c in conciliacion_cmd(conn, items)}
        filas = []
        for k in items:
            c = conc.get(k.item_codigo)
            filas.append({"c": k.item_codigo, "d": k.item_descripcion, "ot": c["codigo_ot"] if c else "",
                          "si": f"{cifra(k.saldo_inicial_cantidad)} / {soles(k.saldo_inicial_total)}",
                          "co": f"{cifra(k.ingresos_cantidad)} / {soles(k.ingresos_total)}",
                          "cs": f"{cifra(k.salidas_cantidad)} / {soles(k.salidas_total)}",
                          "cmd": soles(c["cmd_soles"]) if c else "—",
                          "sf": f"{cifra(k.saldo_final_cantidad)} / {soles(k.saldo_final_total)}",
                          "ok": ("coincide" if c and c["cuadra"] else "NO coincide") if c else "sin OT"})
        total = sum(k.saldo_final_total for k in items)
        return ("KARDEX DE COMPRAS — almacén de materias primas (promedio ponderado; datos ficticios)\n"
                + tabla(filas, [("c", "Material"), ("d", "Descripción"), ("ot", "OT"), ("si", "Saldo inicial"),
                                ("co", "Compras netas"), ("cs", "Salidas a la OT"), ("cmd", "CMD de la OT"),
                                ("sf", "Saldo final"), ("ok", "Salidas vs CMD")])
                + f"\nLas salidas del kardex son el consumo de material directo de cada OT. Saldo final de materias primas: "
                  f"{soles(total)}. Detalle por movimiento: pestaña Datos importados → kardex_compras, "
                  "o Reportes → Generar kardex promedio.")

    def rE2(self) -> Respuesta:
        kardexs, obs = kardex_desde_base(self.L)
        if self.ot_sel:
            mats = {m.get("material_id") for m in self.L.calculado("cmd_ot") if m.get("ot_id") in self.ot_sel}
            kardexs = [k for k in kardexs if k.material_id in mats]
        bloques = []
        for k in kardexs:
            filas = [{"d": f.detalle, "ref": f.referencia, "eq": cifra(f.ent_cantidad) if f.ent_cantidad else "",
                      "ecu": f"{f.ent_costo_unitario:,.4f}" if f.ent_cantidad else "", "et": cifra(f.ent_total) if f.ent_cantidad else "",
                      "sq": cifra(f.sal_cantidad) if f.sal_cantidad else "", "scu": f"{f.sal_costo_unitario:,.4f}" if f.sal_cantidad else "",
                      "st": cifra(f.sal_total) if f.sal_cantidad else "", "q": cifra(f.saldo_cantidad),
                      "cu": f"{f.saldo_costo_promedio:,.4f}", "t": cifra(f.saldo_total)} for f in k.filas]
            ec = k.ecuacion_almacen()
            bloques.append(
                f"KARDEX — {k.material} (promedio ponderado, S/)\n"
                + tabla(filas, [("d", "Detalle"), ("ref", "Ref."), ("eq", "Ent. kg"), ("ecu", "Ent. C.U."), ("et", "Ent. total"),
                                ("sq", "Sal. kg"), ("scu", "Sal. C.U."), ("st", "Sal. total"), ("q", "Saldo kg"), ("cu", "C.U. prom."),
                                ("t", "Saldo S/")])
                + f"\nII {soles(ec['inventario_inicial'])} + Compras {soles(ec['compras'])} − IF {soles(ec['inventario_final'])} = "
                  f"Consumo {soles(ec['consumo'])} ({'cuadra' if ec['cuadra'] else 'NO cuadra'})")
        saldo = sum(k.inventario_final for k in kardexs)
        consumo = sum(k.consumo for k in kardexs)
        compras = self._kardex_compras()
        if compras:
            bloques.append(compras)
        return Respuesta("Kardex de materiales y de compras",
                         f"Consumo valorizado de materiales (CMD): {soles(consumo)}. Saldo final de inventario de materiales: {soles(saldo)}."
                         + (" Abajo, el kardex de la hoja Compras (almacén de materias primas) y su conciliación con el CMD."
                            if compras else ""),
                         "\n\n".join(bloques),
                         formula="Entrada: C.U. = (cantidad × precio + flete) / cantidad; salida al costo promedio vigente; "
                                 "II + Compras − IF = Consumo",
                         fuente=self.fuente_calculo("kardex_promedio", ["Compras"]), observaciones=obs)
