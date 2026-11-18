"""Catalogo base del agente: se construye a partir de preguntas_base.txt.

El archivo de texto es la fuente: categorias (lineas sin numero) y preguntas
numeradas ("12. ¿...?"). Aqui solo se agregan, por numero de pregunta:

- ejemplos: otras formas en que un usuario de Finanzas/Contabilidad haria la
  misma pregunta (mejoran la interpretacion);
- ranking: plantillas para reformular cuando el usuario pide "el mayor" o
  "el menor" (p. ej. "¿cual es el producto con mas ventas?").

Dos ampliaciones (E1 y E2) cubren preguntas frecuentes que el txt no trae de
forma literal: costo de ventas por objeto de costo y kardex completo.
"""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

ARCHIVO_BASE = "preguntas_base.txt"
SLOTS = {"{OT}": "ENT_OT", "{actividad}": "ENT_ACTIVIDAD", "{recurso}": "ENT_RECURSO"}


@dataclass
class PreguntaBase:
    codigo: str  # "1".."48" o "E1", "E2"
    categoria: str
    texto: str
    ejemplos: list[str] = field(default_factory=list)
    ranking_max: str | None = None
    ranking_min: str | None = None
    relacionada: str | None = None
    dominante: str | None = None  # concepto que, si aparece, apunta casi siempre a esta pregunta

    @property
    def slots(self) -> list[str]:
        return [slot for slot in SLOTS if slot in self.texto]

    @property
    def es_ampliacion(self) -> bool:
        return self.codigo.startswith("E")

    @property
    def etiqueta(self) -> str:
        if self.es_ampliacion:
            return f"Ampliación {self.codigo}, relacionada con la pregunta {self.relacionada}"
        return f"Pregunta {self.codigo} del catálogo base"


# Ejemplos de parafraseo y plantillas de ranking por numero de pregunta.
META: dict[str, dict] = {
    "1": {"ejemplos": ["que empresa estamos costeando", "de que compania es el analisis y que metodo de costeo se usa",
                       "cual es el periodo evaluado"]},
    "2": {"ejemplos": ["que productos hay", "lista de ordenes de trabajo", "cuales son los objetos de costo",
                       "que articulos se costean"]},
    "3": {"ejemplos": ["cuantas piezas se fabricaron de cada producto", "volumen de produccion por orden",
                       "cantidad producida de {OT}"],
          "max": "¿Qué objeto de costo tuvo la mayor cantidad producida?",
          "min": "¿Qué objeto de costo tuvo la menor cantidad producida?"},
    "4": {"ejemplos": ["que procesos productivos hay", "cuales son las etapas de produccion", "lista de actividades"]},
    "5": {"dominante": "FUENTE", "ejemplos": ["de donde sale el dato de ventas", "que cuadro respalda el material directo",
                       "origen del dato de mantenimiento", "en que hoja esta el dato"]},
    "6": {"dominante": "OBSERVACION", "ejemplos": ["que notas tiene el dato de ventas", "comentarios del calculo de mano de obra",
                       "aclaraciones de los datos importados"]},
    "7": {"ejemplos": ["el excel esta completo", "faltan columnas en el maestro", "la plantilla tiene todas las hojas",
                       "se valido bien el archivo de entrada"]},
    "8": {"dominante": "ERROR", "ejemplos": ["hay inconsistencias en los inductores", "los pesos suman 100 %", "errores de importacion",
                       "advertencias del archivo"]},
    "9": {"ejemplos": ["cuanta materia prima uso {OT}", "consumo de insumos de la orden", "material consumido por producto"],
          "max": "¿Qué objeto de costo consumió más material directo (CMD)?",
          "min": "¿Qué objeto de costo consumió menos material directo (CMD)?"},
    "10": {"ejemplos": ["precio unitario de cada insumo", "kilos y precio de cada material",
                        "cantidad y precio unitario de las materias primas"]},
    "11": {"ejemplos": ["cuanto cuesta la materia prima de {OT}", "importe del material directo de la orden",
                        "costo de material directo total"],
           "max": "¿Qué objeto de costo tiene el mayor CMD total?",
           "min": "¿Qué objeto de costo tiene el menor CMD total?"},
    "12": {"ejemplos": ["que costo unitario promedio dio el kardex", "a que costo unitario se valorizo la salida del almacen",
                        "costo unitario de valorizacion por promedio ponderado", "precio promedio de los insumos"]},
    "13": {"ejemplos": ["hay servicios de terceros en la orden", "por que no hay servicio directo",
                        "csd de {OT}", "servicios tercerizados"]},
    "14": {"ejemplos": ["cuanto demora cada actividad por kg", "te de cada proceso", "tolerancias y tiempo normal"]},
    "15": {"ejemplos": ["cual es el fd", "factor de eficiencia de la mano de obra", "relacion horas reales y estandar"]},
    "16": {"ejemplos": ["que proporcion de la mano de obra va a cada proceso", "porcentaje de mod por actividad",
                        "participacion de cada actividad en la planilla"]},
    "17": {"ejemplos": ["cuanto se paga de mano de obra en cada proceso", "sueldos por actividad",
                        "costo de personal directo por actividad"],
           "max": "¿Qué actividad tiene el mayor costo de MOD asignado?",
           "min": "¿Qué actividad tiene el menor costo de MOD asignado?"},
    "18": {"ejemplos": ["cuales son los gastos indirectos y sus montos", "lista de cif", "monto de cada recurso indirecto",
                        "costo de cada recurso indirecto"],
           "max": "¿Qué recurso indirecto tiene el mayor importe?",
           "min": "¿Qué recurso indirecto tiene el menor importe?"},
    "19": {"ejemplos": ["que porcentaje de cada recurso va a {actividad}", "inductores de recursos de {actividad}",
                        "pesos ec ac de {actividad}"]},
    "20": {"ejemplos": ["cuanto de {recurso} se cargo a {actividad}", "monto de {recurso} asignado a {actividad}"]},
    "21": {"ejemplos": ["suma de cif por proceso", "total de recursos indirectos de cada actividad",
                        "gastos indirectos por actividad",
                        "cuantos recursos indirectos recibe cada actividad"],
           "max": "¿Qué actividad recibe más otros recursos indirectos?",
           "min": "¿Qué actividad recibe menos otros recursos indirectos?"},
    "22": {"ejemplos": ["cuanto mantenimiento tiene cada proceso", "costo de mantenimiento de maquinas por actividad",
                        "gasto de mantenimiento de equipos"],
           "max": "¿Qué actividad tiene el mayor costo de mantenimiento?",
           "min": "¿Qué actividad tiene el menor costo de mantenimiento?"},
    "23": {"ejemplos": ["cuanto cuestan las pruebas de laboratorio y como se reparten", "costo de calidad por actividad",
                        "distribucion del control de calidad"]},
    "24": {"ejemplos": ["costo de los planos y su reparto", "cuanto se gasta en ingenieria", "distribucion de ingenieria"]},
    "25": {"ejemplos": ["cuanto soporte recibe cada proceso", "costo de apoyo por actividad", "soporte total por actividad"],
           "max": "¿Qué actividad recibe el mayor costo de soporte?",
           "min": "¿Qué actividad recibe el menor costo de soporte?"},
    "26": {"ejemplos": ["cuanto cuesta {actividad}", "cuanto cuesta {actividad} y de que se compone", "desglose del costo de {actividad}",
                        "costo total de la actividad {actividad}"]},
    "27": {"ejemplos": ["que procesos son los mas caros", "ranking de actividades por costo",
                        "actividad que mas cuesta", "actividades mas costosas"],
           "max": "¿Cuáles son las actividades de mayor costo?",
           "min": "¿Cuáles son las actividades de menor costo?"},
    "28": {"ejemplos": ["que porcentaje de {actividad} va a {OT}", "inductor ac ot de {actividad} para {OT}",
                        "peso de distribucion de la actividad a la orden"]},
    "29": {"ejemplos": ["cuanto costo de conversion paso de {actividad} a {OT}",
                        "costo de produccion que recibio {OT} desde {actividad}"]},
    "30": {"ejemplos": ["cuanto soporte paso de {actividad} a {OT}", "soporte que recibio {OT} desde {actividad}"]},
    "31": {"ejemplos": ["cuanta mano de obra y cif recibio {OT}", "costos indirectos y mod que absorbio la orden",
                        "cuanto cif recibio {OT}", "cuanto soporte recibio la orden"]},
    "32": {"ejemplos": ["cuanto costo la orden {OT}", "costo total de produccion del producto", "costo de {OT}"],
           "max": "¿Qué objeto de costo tiene el mayor costo total?",
           "min": "¿Qué objeto de costo tiene el menor costo total?"},
    "33": {"ejemplos": ["cuanto cuesta cada pieza de {OT}", "costo por unidad del producto", "costo unitario por orden"],
           "max": "¿Qué objeto de costo tiene el mayor costo unitario?",
           "min": "¿Qué objeto de costo tiene el menor costo unitario?"},
    "34": {"ejemplos": ["desglose del costo de {OT}", "estructura de costos de la orden", "de que se compone el costo del producto"]},
    "35": {"ejemplos": ["que elemento pesa mas en el costo de {OT}", "que representa la mayor parte del costo",
                        "componente dominante del costo"]},
    "36": {"ejemplos": ["orden con mayor costo total", "que producto es el mas caro de fabricar",
                        "ot con el costo unitario mas alto"]},
    "37": {"ejemplos": ["cuanto se vendio de cada producto", "ingresos por producto", "facturacion por orden",
                        "ventas de {OT}"],
           "max": "¿Cuál es el objeto de costo que genera más ingresos?",
           "min": "¿Cuál es el objeto de costo que genera menos ingresos?"},
    "38": {"ejemplos": ["utilidad bruta de {OT}", "ganancia bruta por producto"],
           "max": "¿Qué objeto de costo tiene el mayor margen bruto?",
           "min": "¿Qué objeto de costo tiene el menor margen bruto?"},
    "39": {"ejemplos": ["cuanto gano {OT}", "utilidad de la orden", "ganancia por producto", "beneficio operativo"],
           "max": "¿Qué objeto de costo genera el mayor margen operativo?",
           "min": "¿Qué objeto de costo genera el menor margen operativo?"},
    "40": {"ejemplos": ["que porcentaje de ganancia tiene {OT}", "rentabilidad de la orden", "margen porcentual del producto"],
           "max": "¿Qué objeto de costo tiene la mayor rentabilidad operativa porcentual?",
           "min": "¿Qué objeto de costo tiene la menor rentabilidad operativa porcentual?"},
    "41": {"ejemplos": ["cual es el producto mas rentable", "que producto rinde menos", "producto menos rentable"]},
    "42": {"ejemplos": ["que orden deberia analizar primero", "donde hay que poner atencion", "que producto es preocupante",
                        "que ot priorizar"]},
    "43": {"dominante": "PROPUESTA", "ejemplos": ["que recomendaciones das", "como mejorar la rentabilidad", "propuestas de mejora", "que riesgos hay",
                        "que hago para ganar mas", "como reducir costos"]},
    "44": {"ejemplos": ["que contiene el reporte del proceso 1", "reporte ec ac de recursos", "reporte de recursos a actividades"]},
    "45": {"ejemplos": ["muestrame la hoja de costos de {OT}", "estado de resultados del producto"]},
    "46": {"ejemplos": ["que kpis tiene el tablero", "que graficos muestra el dashboard", "indicadores del panel"]},
    "47": {"ejemplos": ["que archivos puedo exportar", "que documentos genera el sistema", "que informes saca el aplicativo"]},
    "48": {"dominante": "PRUEBAS", "ejemplos": ["pasaron los tests", "resultados de las pruebas funcionales", "el sistema aprobo las pruebas"]},
}

AMPLIACIONES: list[PreguntaBase] = [
    PreguntaBase(
        codigo="E1",
        categoria="Ventas y rentabilidad",
        texto="¿Cuál es el costo de ventas de {OT}?",
        ejemplos=["que producto genero mas costos", "costo de lo vendido por producto", "que articulo gasto mas",
                  "costo de ventas", "que producto tiene mas costos"],
        ranking_max="¿Cuál objeto de costo generó más costo de ventas?",
        ranking_min="¿Cuál objeto de costo generó menos costo de ventas?",
        relacionada="32",
    ),
    PreguntaBase(
        codigo="E2",
        categoria="Material directo y personal",
        texto="¿Cómo queda el kardex promedio ponderado de los materiales (entradas, salidas y saldo final)?",
        ejemplos=["muestrame el kardex", "saldo de inventario de materiales", "existencias finales del almacen",
                  "inventario final de materia prima", "kardex de {OT}"],
        relacionada="12",
    ),
]


def _ruta_archivo_base() -> Path:
    candidatos = [Path(__file__).with_name(ARCHIVO_BASE)]
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidatos.append(Path(meipass) / "app" / "agente_ia" / ARCHIVO_BASE)
    for ruta in candidatos:
        if ruta.exists():
            return ruta
    raise FileNotFoundError(f"No se encontro el catalogo base del agente: {ARCHIVO_BASE}")


def leer_txt(ruta: str | Path | None = None) -> list[PreguntaBase]:
    """Lee el txt: lineas 'N. pregunta' son preguntas; las demas, categorias."""
    path = Path(ruta) if ruta else _ruta_archivo_base()
    preguntas: list[PreguntaBase] = []
    categoria = "General"
    for linea in path.read_text(encoding="utf-8-sig").splitlines():
        linea = linea.strip()
        if not linea:
            continue
        match = re.match(r"^(\d+)\s*[.)-]\s*(.+)$", linea)
        if match:
            preguntas.append(PreguntaBase(codigo=match.group(1), categoria=categoria, texto=match.group(2).strip()))
        else:
            categoria = linea
    return preguntas


def cargar_catalogo(ruta: str | Path | None = None) -> list[PreguntaBase]:
    preguntas = leer_txt(ruta)
    for pregunta in preguntas:
        meta = META.get(pregunta.codigo, {})
        pregunta.ejemplos = list(meta.get("ejemplos", []))
        pregunta.ranking_max = meta.get("max")
        pregunta.ranking_min = meta.get("min")
        pregunta.dominante = meta.get("dominante")
    return preguntas + [PreguntaBase(**vars(item)) for item in AMPLIACIONES]


def categorias(catalogo: list[PreguntaBase]) -> list[str]:
    vistas: list[str] = []
    for pregunta in catalogo:
        if pregunta.categoria not in vistas:
            vistas.append(pregunta.categoria)
    return vistas
