"""Agente IA de Finanzas y Contabilidad (solo lectura).

Uso:
    agente = AgenteFinanzas("costos_abc_frenosa.sqlite")
    resultado = agente.consultar("¿cual es el producto con mas ventas?")
    print(resultado.texto())
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from .catalogo import PreguntaBase, cargar_catalogo, categorias
from .formato import Respuesta
from .interprete import CatalogoEntidades, InterpreteFinanzas, Interpretacion
from .lectura import LectorSoloLectura
from .propuestas import Propuesta, generar_propuestas, propuestas_como_texto
from .respuestas import Consultor

AVISO_SOLO_LECTURA = (
    "Soy un agente de solo lectura: interpreto preguntas, consulto los resultados y propongo mejoras, "
    "pero no modifico datos, cálculos, archivos ni configuraciones."
)


@dataclass
class ResultadoAgente:
    interpretacion: Interpretacion
    respuesta: Respuesta | None
    mensaje: str = ""

    def texto(self) -> str:
        i = self.interpretacion
        lineas: list[str] = []
        if i.es_pedido_de_cambio:
            lineas += ["PEDIDO DE CAMBIO NO PERMITIDO", AVISO_SOLO_LECTURA,
                       "Si el cambio es necesario, regístrelo en el Excel maestro, vuelva a importarlo y ejecute el cálculo.", ""]
        if not i.entendida:
            lineas += ["NO LOGRÉ INTERPRETAR LA PREGUNTA",
                       "Reformúlela usando términos como objeto de costo / producto, OT, actividad, recurso, ventas, "
                       "costo total, costo unitario, margen o rentabilidad."]
            if i.alternativas:
                lineas.append("¿Quiso decir…?")
                lineas += [f"  - {p.texto} ({p.etiqueta})" for p, _ in i.alternativas]
            return "\n".join(lineas)
        p = i.pregunta
        lineas += ["PREGUNTA INTERPRETADA", i.interpretada,
                   f"({p.etiqueta} · {p.categoria} · confianza {i.confianza * 100:,.0f} %)"]
        if i.terminos:
            lineas.append("Términos interpretados: " + "; ".join(f"{a} → {b}" for a, b in i.terminos))
        if not i.segura and i.alternativas:
            lineas.append("También podría referirse a: " + " | ".join(a.texto for a, _ in i.alternativas))
        if self.respuesta is not None:
            lineas += ["", "RESPUESTA", self.respuesta.como_texto()]
        if self.mensaje:
            lineas += ["", self.mensaje]
        return "\n".join(lineas)


class AgenteFinanzas:
    nombre = "Agente IA de Finanzas y Contabilidad"

    def __init__(self, db_path: str | Path | None = None, catalogo_txt: str | Path | None = None):
        self.db_path = db_path
        self.catalogo: list[PreguntaBase] = cargar_catalogo(catalogo_txt)
        self.interprete = InterpreteFinanzas(self.catalogo)

    def categorias(self) -> list[str]:
        return categorias(self.catalogo)

    def interpretar(self, pregunta: str) -> Interpretacion:
        with LectorSoloLectura(self.db_path) as lector:
            entidades = CatalogoEntidades.desde_base(lector) if lector.disponible else CatalogoEntidades([])
            return self.interprete.interpretar(pregunta, entidades)

    def consultar(self, pregunta: str, responder: bool = True) -> ResultadoAgente:
        with LectorSoloLectura(self.db_path) as lector:
            entidades = CatalogoEntidades.desde_base(lector) if lector.disponible else CatalogoEntidades([])
            interpretacion = self.interprete.interpretar(pregunta, entidades)
            respuesta = None
            if responder and interpretacion.entendida:
                try:
                    respuesta = Consultor(lector, interpretacion).responder()
                except Exception as exc:  # el agente nunca debe romper la aplicacion
                    respuesta = Respuesta("Error", f"No pude armar la respuesta con los datos actuales: {exc}", sin_datos=True)
            return ResultadoAgente(interpretacion, respuesta)

    def propuestas(self) -> list[Propuesta]:
        with LectorSoloLectura(self.db_path) as lector:
            if not lector.disponible or lector.ultimo_run() is None:
                return []
            return generar_propuestas(lector)

    def propuestas_texto(self) -> str:
        props = self.propuestas()
        encabezado = "PROPUESTAS DE MEJORA — " + self.nombre + "\n" + AVISO_SOLO_LECTURA + "\n\n"
        if not props:
            return encabezado + "No hay un cálculo completado para sustentar propuestas. Ejecute el Modelo ABC por Órdenes de Trabajo."
        return encabezado + propuestas_como_texto(props)
