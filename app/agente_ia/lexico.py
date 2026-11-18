"""Lexico de Finanzas y Contabilidad de costos.

Traduce el vocabulario cotidiano del usuario ("producto", "ventas", "ganancia")
al vocabulario tecnico del Modelo ABC por Ordenes de Trabajo ("objeto de costo",
"ingresos / ventas netas", "margen operativo"). Cada termino reconocido se
convierte en un CONCEPTO que luego usa el interprete para comparar la pregunta
con el catalogo base (preguntas_base.txt).

No contiene datos de la empresa: las OT, actividades y recursos se leen de la
base SQLite en tiempo de ejecucion (ver interprete.py).
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass

# ---------------------------------------------------------------------------
# Normalizacion
# ---------------------------------------------------------------------------

STOPWORDS = {
    "el", "la", "los", "las", "un", "una", "unos", "unas", "de", "del", "al", "a",
    "en", "y", "o", "u", "e", "que", "cual", "cuales", "es", "son", "fue", "fueron",
    "por", "para", "con", "se", "su", "sus", "lo", "le", "les", "me", "mi", "mis",
    "como", "donde", "quien", "hay", "tiene", "tienen", "tuvo", "esta", "estan",
    "este", "esta", "esto", "eso", "ese", "esa", "cada", "sobre", "entre", "segun",
    "cuanto", "cuanta", "cuantos", "cuantas", "muestra", "muestrame", "dime", "dame",
    "quiero", "saber", "ver", "podrias", "puedes", "favor", "porfa", "hola", "buenas",
    "genera", "genero", "generan", "generaron", "registrado", "registrados",
    "registrada", "registradas", "usado", "usados", "uso", "utilizo", "utilizado",
    "cual", "cuál", "qué", "and", "the", "ha", "han", "sido", "ser", "seria",
}


def strip_accents(text: str) -> str:
    text = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in text if not unicodedata.combining(ch))


def normalize(text: str) -> str:
    """Minusculas, sin tildes y sin signos (conserva letras, numeros y %)."""
    text = strip_accents(str(text or "")).lower()
    text = text.replace("ñ", "n")
    text = re.sub(r"[{}]", " ", text)
    text = re.sub(r"[^a-z0-9%\s\-_]", " ", text)
    text = re.sub(r"[\-_]", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def tokenize(text: str) -> list[str]:
    return normalize(text).split()


def tokenize_display(text: str) -> list[str]:
    """Tokens con tildes (para mostrar), alineados con tokenize()."""
    text = str(text or "").lower()
    text = re.sub(r"[{}]", " ", text)
    text = re.sub(r"[^\w%\s\-]", " ", text)
    text = re.sub(r"[\-_]", " ", text)
    tokens = text.split()
    return tokens if len(tokens) == len(tokenize(" ".join(tokens))) else tokenize(text)


def compact(text: str) -> str:
    """Forma compacta para comparar codigos: 'OT-FRENOSA_SW 50' -> 'otfrenosasw50'."""
    return re.sub(r"[^a-z0-9]", "", strip_accents(str(text or "")).lower())


# ---------------------------------------------------------------------------
# Conceptos
# ---------------------------------------------------------------------------
# Cada concepto tiene un nombre tecnico legible (para explicar la
# interpretacion al usuario) y un peso base (importancia al comparar).

CONCEPTOS: dict[str, tuple[str, float]] = {
    "OBJETO": ("objeto de costo = orden de trabajo (OT)", 1.0),
    "VENTAS": ("ingresos (ventas netas)", 2.2),
    "COSTO": ("costo", 1.2),
    "COSTO_VENTAS": ("costo de ventas", 2.4),
    "UNITARIO": ("costo unitario", 2.0),
    "MARGEN": ("margen", 1.6),
    "MARGEN_BRUTO": ("margen bruto", 2.6),
    "MARGEN_OPERATIVO": ("margen operativo", 2.4),
    "RENTABILIDAD": ("rentabilidad operativa (%)", 2.4),
    "MAX": ("criterio: el mayor", 0.6),
    "MIN": ("criterio: el menor", 0.6),
    "CMD": ("costo de material directo (CMD)", 2.2),
    "CONSUMO": ("consumo", 1.6),
    "KARDEX": ("kardex / valorizacion promedio de inventarios", 2.2),
    "SALDO": ("saldo / inventario final", 2.0),
    "EXCEL": ("archivo Excel", 1.0),
    "MOD": ("mano de obra directa (MOD) / personal directo", 2.2),
    "TIEMPO_STD": ("tiempo estandar (Te)", 2.6),
    "FD": ("factor de distribucion (Fd)", 2.8),
    "PORCENTAJE": ("porcentaje / proporcion", 1.4),
    "ACTIVIDAD": ("actividad productiva", 1.4),
    "PRODUCCION": ("actividades productivas / costo de conversion", 1.6),
    "RECURSO": ("recursos indirectos (CIF)", 1.8),
    "SOPORTE": ("procesos de soporte", 1.8),
    "MANTENIMIENTO": ("soporte: mantenimiento", 2.4),
    "CALIDAD": ("soporte: control de calidad", 2.4),
    "INGENIERIA": ("soporte: ingenieria y desarrollo", 2.4),
    "CSD": ("costo de servicio directo (CSD)", 2.6),
    "PESO": ("peso / inductor de distribucion", 1.8),
    "DISTRIBUCION": ("distribucion / asignacion", 1.0),
    "TRANSFERENCIA": ("costo transferido / recibido", 1.6),
    "COMPOSICION": ("composicion / estructura del costo", 1.8),
    "TOTAL": ("total", 0.8),
    "UNIDADES": ("unidades producidas", 2.0),
    "EMPRESA": ("empresa", 1.6),
    "PERIODO": ("periodo de costeo", 1.4),
    "MODELO": ("modelo de costeo", 1.4),
    "FUENTE": ("hoja de origen / fuente del dato", 2.2),
    "OBSERVACION": ("observaciones del dato", 2.4),
    "VALIDACION": ("validacion del Excel maestro", 2.2),
    "ERROR": ("errores / advertencias de importacion", 2.2),
    "ALERTA": ("alertas del dashboard", 2.0),
    "PROPUESTA": ("propuestas de mejora", 2.2),
    "REVISAR": ("prioridad de revision", 1.8),
    "REPORTE": ("reportes y documentos", 1.8),
    "DASHBOARD": ("dashboard / indicadores", 2.0),
    "HOJA_COSTOS": ("hoja de costos", 2.2),
    "ESTADO_RESULTADOS": ("estado de resultados", 2.2),
    "PRUEBAS": ("informe de pruebas funcionales", 2.6),
    "ASIGNACION_REPORTE": ("reporte de asignacion de recursos", 1.6),
    "LISTA": ("listado", 0.4),
    "IMPORTE": ("importe / monto", 0.5),
    "ENT_OT": ("[OT indicada]", 0.8),
    "ENT_ACTIVIDAD": ("[actividad indicada]", 0.8),
    "ENT_RECURSO": ("[recurso indicado]", 0.8),
}

# Frases de varias palabras (se evaluan primero, de la mas larga a la mas corta).
FRASES: list[tuple[str, tuple[str, ...]]] = [
    ("costo de ventas", ("COSTO_VENTAS", "COSTO")),
    ("costos de ventas", ("COSTO_VENTAS", "COSTO")),
    ("costo de venta", ("COSTO_VENTAS", "COSTO")),
    ("costo de lo vendido", ("COSTO_VENTAS", "COSTO")),
    ("costo de los productos vendidos", ("COSTO_VENTAS", "COSTO")),
    ("objeto de costo", ("OBJETO",)),
    ("objetos de costo", ("OBJETO",)),
    ("orden de trabajo", ("OBJETO",)),
    ("ordenes de trabajo", ("OBJETO",)),
    ("margen bruto", ("MARGEN_BRUTO",)),
    ("utilidad bruta", ("MARGEN_BRUTO",)),
    ("ganancia bruta", ("MARGEN_BRUTO",)),
    ("resultado bruto", ("MARGEN_BRUTO",)),
    ("margen operativo", ("MARGEN_OPERATIVO",)),
    ("margen operacional", ("MARGEN_OPERATIVO",)),
    ("utilidad operativa", ("MARGEN_OPERATIVO",)),
    ("utilidad operacional", ("MARGEN_OPERATIVO",)),
    ("ganancia operativa", ("MARGEN_OPERATIVO",)),
    ("resultado operativo", ("MARGEN_OPERATIVO",)),
    ("margen porcentual", ("RENTABILIDAD",)),
    ("rentabilidad operativa porcentual", ("RENTABILIDAD",)),
    ("rentabilidad operativa", ("RENTABILIDAD",)),
    ("porcentaje de ganancia", ("RENTABILIDAD",)),
    ("porcentaje de utilidad", ("RENTABILIDAD",)),
    ("materia prima", ("CMD",)),
    ("materias primas", ("CMD",)),
    ("material directo", ("CMD",)),
    ("materiales directos", ("CMD",)),
    ("promedio ponderado", ("KARDEX",)),
    ("costo promedio", ("KARDEX",)),
    ("costo unitario promedio", ("KARDEX", "UNITARIO")),
    ("mano de obra", ("MOD",)),
    ("personal directo", ("MOD",)),
    ("tiempo estandar", ("TIEMPO_STD",)),
    ("tiempos estandar", ("TIEMPO_STD",)),
    ("tiempo normal", ("TIEMPO_STD",)),
    ("tiempo suplementario", ("TIEMPO_STD",)),
    ("factor de distribucion", ("FD",)),
    ("costo de conversion", ("PRODUCCION",)),
    ("costos de conversion", ("PRODUCCION",)),
    ("costo de produccion", ("PRODUCCION", "COSTO")),
    ("precio de venta", ("VENTAS", "UNITARIO")),
    ("precios de venta", ("VENTAS", "UNITARIO")),
    ("una pieza", ("UNITARIO",)),
    ("una unidad", ("UNITARIO",)),
    ("la pieza", ("UNITARIO",)),
    ("por kilo", ("UNITARIO",)),
    ("al cierre", ("SALDO",)),
    ("otros recursos", ("RECURSO",)),
    ("gastos indirectos", ("RECURSO",)),
    ("costos indirectos", ("RECURSO",)),
    ("costo indirecto", ("RECURSO",)),
    ("elemento de costo", ("RECURSO",)),
    ("elementos de costo", ("RECURSO",)),
    ("procesos de soporte", ("SOPORTE",)),
    ("areas de apoyo", ("SOPORTE",)),
    ("control de calidad", ("CALIDAD",)),
    ("prueba de laboratorio", ("CALIDAD",)),
    ("pruebas de laboratorio", ("CALIDAD",)),
    ("ingenieria y desarrollo", ("INGENIERIA",)),
    ("servicio directo", ("CSD",)),
    ("servicios directos", ("CSD",)),
    ("servicio de terceros", ("CSD",)),
    ("servicios de terceros", ("CSD",)),
    ("hoja de costos", ("HOJA_COSTOS",)),
    ("hoja de costo", ("HOJA_COSTOS",)),
    ("hojas de costos", ("HOJA_COSTOS",)),
    ("estado de resultados", ("ESTADO_RESULTADOS",)),
    ("estado de resultado", ("ESTADO_RESULTADOS",)),
    ("estado de ganancias y perdidas", ("ESTADO_RESULTADOS",)),
    ("pruebas funcionales", ("PRUEBAS",)),
    ("prueba funcional", ("PRUEBAS",)),
    ("informe de pruebas", ("PRUEBAS",)),
    ("asignacion de recursos", ("ASIGNACION_REPORTE", "RECURSO")),
    ("ec ac", ("ASIGNACION_REPORTE",)),
    ("proceso 1", ("ASIGNACION_REPORTE",)),
    ("ac ot", ("PESO",)),
    ("inventario final", ("KARDEX", "SALDO")),
    ("saldo final", ("KARDEX", "SALDO")),
    ("excel maestro", ("VALIDACION",)),
    ("hoja del excel", ("FUENTE",)),
    ("de donde", ("FUENTE",)),
    ("por unidad", ("UNITARIO",)),
    ("por pieza", ("UNITARIO",)),
    ("cada pieza", ("UNITARIO",)),
    ("cada unidad", ("UNITARIO",)),
    ("por kg", ("UNITARIO",)),
    ("modelo de costeo", ("MODELO",)),
    ("sistema de costeo", ("MODELO",)),
    ("que hacer", ("PROPUESTA",)),
    ("que hago", ("PROPUESTA",)),
    ("como mejorar", ("PROPUESTA",)),
    ("como aumentar", ("PROPUESTA",)),
    ("como reducir", ("PROPUESTA",)),
    ("como bajar", ("PROPUESTA",)),
    ("en que", ("LISTA",)),
]

# Palabras exactas.
PALABRAS: dict[str, tuple[str, ...]] = {
    "producto": ("OBJETO",), "productos": ("OBJETO",), "articulo": ("OBJETO",),
    "articulos": ("OBJETO",), "item": ("OBJETO",), "items": ("OBJETO",),
    "forro": ("OBJETO",), "forros": ("OBJETO",), "ot": ("OBJETO",), "ots": ("OBJETO",),
    "orden": ("OBJETO",), "ordenes": ("OBJETO",), "lote": ("OBJETO",), "lotes": ("OBJETO",),
    "pedido": ("OBJETO",), "pedidos": ("OBJETO",), "referencia": ("OBJETO",),
    "mas": ("MAX",), "mayor": ("MAX",), "mayores": ("MAX",), "top": ("MAX",),
    "lider": ("MAX",), "alto": ("MAX",), "alta": ("MAX",), "altos": ("MAX",),
    "altas": ("MAX",), "mejor": ("MAX",), "mejores": ("MAX",), "principal": ("MAX",),
    "principales": ("MAX",), "elevado": ("MAX",), "elevada": ("MAX",),
    "menos": ("MIN",), "menor": ("MIN",), "menores": ("MIN",), "bajo": ("MIN",),
    "baja": ("MIN",), "bajos": ("MIN",), "bajas": ("MIN",), "peor": ("MIN",),
    "peores": ("MIN",), "ultimo": ("MIN",), "reducido": ("MIN",),
    "caro": ("UNITARIO", "MAX"), "cara": ("UNITARIO", "MAX"), "caros": ("UNITARIO", "MAX"),
    "barato": ("UNITARIO", "MIN"), "barata": ("UNITARIO", "MIN"), "baratos": ("UNITARIO", "MIN"),
    "margen": ("MARGEN",), "margenes": ("MARGEN",),
    "cmd": ("CMD",), "mp": ("CMD",), "mpd": ("CMD",),
    "kardex": ("KARDEX",), "stock": ("KARDEX",), "ponderado": ("KARDEX",),
    "promedio": ("KARDEX",),
    "repuesto": ("KARDEX",), "repuestos": ("KARDEX",),
    "salida": ("KARDEX",), "salidas": ("KARDEX",), "entradas": ("KARDEX",),
    "saldo": ("SALDO",), "saldos": ("SALDO",), "final": ("SALDO",), "finales": ("SALDO",),
    "cierre": ("SALDO",), "precio": ("UNITARIO",), "precios": ("UNITARIO",),
    "factor": ("FD",), "critica": ("REVISAR",), "critico": ("REVISAR",), "criticas": ("REVISAR",),
    "criticos": ("REVISAR",), "usaron": ("CONSUMO",), "usa": ("CONSUMO",), "usan": ("CONSUMO",),
    "queda": ("SALDO",), "quedan": ("SALDO",), "quedo": ("SALDO",),
    "participan": ("LISTA",), "participa": ("LISTA",), "intervienen": ("LISTA",), "interviene": ("LISTA",),
    "hoja": ("FUENTE",), "viene": ("FUENTE",), "vienen": ("FUENTE",),
    "excel": ("EXCEL",), "pesa": ("COMPOSICION",), "elemento": ("COMPOSICION",),
    "mod": ("MOD",), "pd": ("MOD",), "planilla": ("MOD",), "trabajadores": ("MOD",),
    "te": ("TIEMPO_STD",), "estandar": ("TIEMPO_STD",),
    "fd": ("FD",), "eficiencia": ("FD",), "%": ("PORCENTAJE",), "proporcion": ("PORCENTAJE",),
    "cuota": ("PORCENTAJE",),
    "proceso": ("ACTIVIDAD",), "procesos": ("ACTIVIDAD",), "tarea": ("ACTIVIDAD",),
    "tareas": ("ACTIVIDAD",), "etapa": ("ACTIVIDAD",), "etapas": ("ACTIVIDAD",),
    "operacion": ("ACTIVIDAD",), "operaciones": ("ACTIVIDAD",),
    "produccion": ("PRODUCCION",), "productiva": ("PRODUCCION",),
    "productivas": ("PRODUCCION",), "productivo": ("PRODUCCION",), "conversion": ("PRODUCCION",),
    "cif": ("RECURSO",), "ci": ("RECURSO",), "overhead": ("RECURSO",),
    "soporte": ("SOPORTE",), "apoyo": ("SOPORTE",),
    "calidad": ("CALIDAD",), "laboratorio": ("CALIDAD",),
    "csd": ("CSD",),
    "reparto": ("DISTRIBUCION",), "mix": ("COMPOSICION",), "estructura": ("COMPOSICION",),
    "suma": ("TOTAL",), "global": ("TOTAL",),
    "unidades": ("UNIDADES",), "piezas": ("UNIDADES",), "produjo": ("UNIDADES",),
    "produjeron": ("UNIDADES",), "volumen": ("UNIDADES",), "kilos": ("UNIDADES",),
    "fabrico": ("UNIDADES",), "fabricaron": ("UNIDADES",), "cantidad": ("UNIDADES",),
    "cantidades": ("UNIDADES",),
    "empresa": ("EMPRESA",), "compania": ("EMPRESA",), "negocio": ("EMPRESA",),
    "mes": ("PERIODO",), "meses": ("PERIODO",), "fecha": ("PERIODO",),
    "costeo": ("LISTA",), "costear": ("LISTA",), "costea": ("LISTA",), "metodo": ("MODELO",), "metodologia": ("MODELO",),
    "origen": ("FUENTE",), "proviene": ("FUENTE",), "provienen": ("FUENTE",),
    "procede": ("FUENTE",), "sale": ("FUENTE",),
    "columnas": ("VALIDACION",), "columna": ("VALIDACION",), "maestro": ("VALIDACION",),
    "hojas": ("VALIDACION",), "requeridas": ("VALIDACION",), "obligatorias": ("VALIDACION",),
    "mejora": ("PROPUESTA",), "mejoras": ("PROPUESTA",), "mejorar": ("PROPUESTA",),
    "plan": ("PROPUESTA",), "acciones": ("PROPUESTA",),
    "conviene": ("REVISAR",), "atencion": ("REVISAR",), "atender": ("REVISAR",),
    "primero": ("REVISAR",),
    "word": ("REPORTE",), "exportar": ("REPORTE",), "aplicativo": ("REPORTE",),
    "dashboard": ("DASHBOARD",), "tablero": ("DASHBOARD",), "panel": ("DASHBOARD",),
    "kpi": ("DASHBOARD",), "kpis": ("DASHBOARD",),
    "test": ("PRUEBAS",), "tests": ("PRUEBAS",),
    "lista": ("LISTA",), "listado": ("LISTA",), "listar": ("LISTA",),
    "importe": ("IMPORTE",), "monto": ("IMPORTE",), "valor": ("IMPORTE",),
    "soles": ("IMPORTE",), "dinero": ("IMPORTE",),
    "recibio": ("TRANSFERENCIA",), "recibe": ("TRANSFERENCIA",), "recibieron": ("TRANSFERENCIA",),
    "llego": ("TRANSFERENCIA",), "llega": ("TRANSFERENCIA",), "cargo": ("TRANSFERENCIA",),
    "absorbio": ("TRANSFERENCIA",),
}

# Prefijos (raices) -> conceptos. Se aplican a palabras no reconocidas antes.
PREFIJOS: list[tuple[str, tuple[str, ...]]] = [
    ("vent", ("VENTAS",)), ("vend", ("VENTAS",)), ("ingres", ("VENTAS",)),
    ("factur", ("VENTAS",)), ("recaud", ("VENTAS",)),
    ("cost", ("COSTO",)), ("gast", ("COSTO",)), ("cuest", ("COSTO",)), ("egres", ("COSTO",)),
    ("unitari", ("UNITARIO",)),
    ("gana", ("MARGEN",)), ("gano", ("MARGEN",)), ("utilid", ("MARGEN",)),
    ("benefic", ("MARGEN",)), ("lucr", ("MARGEN",)),
    ("operativ", ("MARGEN_OPERATIVO",)),
    ("rentab", ("RENTABILIDAD",)), ("rendimient", ("RENTABILIDAD",)), ("retorno", ("RENTABILIDAD",)),
    ("maxim", ("MAX",)), ("minim", ("MIN",)),
    ("material", ("CMD",)), ("insum", ("CMD",)),
    ("inventari", ("KARDEX",)), ("existenci", ("KARDEX",)), ("almacen", ("KARDEX",)),
    ("valoriz", ("KARDEX",)), ("valora", ("KARDEX",)), ("valuac", ("KARDEX",)),
    ("operari", ("MOD",)), ("sueld", ("MOD",)), ("salari", ("MOD",)), ("remuner", ("MOD",)),
    ("toleranci", ("TIEMPO_STD",)),
    ("porcentaj", ("PORCENTAJE",)), ("porcentual", ("PORCENTAJE",)),
    ("actividad", ("ACTIVIDAD",)),
    ("recurso", ("RECURSO",)), ("indirect", ("RECURSO",)),
    ("mantenim", ("MANTENIMIENTO",)), ("maquin", ("MANTENIMIENTO",)), ("equipo", ("MANTENIMIENTO",)),
    ("ensay", ("CALIDAD",)),
    ("ingenier", ("INGENIERIA",)), ("plano", ("INGENIERIA",)),
    ("tercer", ("CSD",)),
    ("peso", ("PESO",)), ("inductor", ("PESO",)), ("driver", ("PESO",)), ("criterio", ("PESO",)),
    ("prorrat", ("PESO", "DISTRIBUCION")),
    ("distribu", ("DISTRIBUCION",)), ("repart", ("DISTRIBUCION",)), ("asign", ("DISTRIBUCION",)),
    ("transfi", ("TRANSFERENCIA",)), ("trasla", ("TRANSFERENCIA",)), ("traspas", ("TRANSFERENCIA",)),
    ("compon", ("COMPOSICION",)), ("composic", ("COMPOSICION",)), ("desglos", ("COMPOSICION",)),
    ("particip", ("COMPOSICION",)), ("divid", ("COMPOSICION",)),
    ("total", ("TOTAL",)), ("acumulad", ("TOTAL",)),
    ("producid", ("UNIDADES",)), ("fabricad", ("UNIDADES",)),
    ("periodo", ("PERIODO",)),
    ("modelo", ("MODELO",)),
    ("fuente", ("FUENTE",)), ("origin", ("FUENTE",)),
    ("consum", ("CONSUMO",)), ("emple", ("CONSUMO",)), ("cuadro", ("FUENTE",)), ("trazab", ("FUENTE",)),
    ("observac", ("OBSERVACION",)), ("nota", ("OBSERVACION",)), ("comentari", ("OBSERVACION",)),
    ("aclarac", ("OBSERVACION",)),
    ("valid", ("VALIDACION",)),
    ("error", ("ERROR",)), ("advertenci", ("ERROR",)), ("inconsisten", ("ERROR",)),
    ("relacion", ("ERROR",)),
    ("alert", ("ALERTA",)), ("riesg", ("ALERTA",)), ("aviso", ("ALERTA",)),
    ("propuest", ("PROPUESTA",)), ("recomend", ("PROPUESTA",)), ("suger", ("PROPUESTA",)),
    ("optimiz", ("PROPUESTA",)),
    ("revis", ("REVISAR",)), ("prioriz", ("REVISAR",)), ("preocup", ("REVISAR",)),
    ("reporte", ("REPORTE",)), ("document", ("REPORTE",)), ("manual", ("REPORTE",)),
    ("informe", ("REPORTE",)),
    ("grafic", ("DASHBOARD",)), ("indicador", ("DASHBOARD",)),
]

# Sinonimos de palabras que el usuario usa y su equivalente tecnico, para
# mostrar "como se entendio" la pregunta.
GLOSARIO_VISIBLE: dict[str, str] = {
    "producto": "objeto de costo (OT)", "productos": "objetos de costo (OT)", "articulo": "objeto de costo (OT)",
    "item": "objeto de costo (OT)", "forro": "objeto de costo (OT)", "lote": "objeto de costo (OT)",
    "pedido": "objeto de costo (OT)",
    "ventas": "ingresos (ventas netas)", "venta": "ingresos (ventas netas)",
    "vendio": "ingresos (ventas netas)", "vende": "ingresos (ventas netas)",
    "vendieron": "ingresos (ventas netas)", "facturo": "ingresos (ventas netas)",
    "facturacion": "ingresos (ventas netas)", "ingresos": "ingresos (ventas netas)",
    "costos": "costo de ventas", "gastos": "costo de ventas", "gasto": "costo de ventas",
    "cuesta": "costo", "caro": "costo unitario mayor", "barato": "costo unitario menor",
    "ganancia": "margen operativo", "ganancias": "margen operativo", "utilidad": "margen operativo",
    "beneficio": "margen operativo", "gano": "margen operativo",
    "rentable": "rentabilidad operativa (%)", "rentables": "rentabilidad operativa (%)",
    "materia prima": "material directo (CMD)", "insumos": "material directo (CMD)",
    "insumo": "material directo (CMD)", "materiales": "material directo (CMD)",
    "inventario": "kardex promedio ponderado", "inventarios": "kardex promedio ponderado",
    "stock": "kardex promedio ponderado", "existencias": "kardex promedio ponderado",
    "mano de obra": "MOD (personal directo)", "operarios": "MOD (personal directo)",
    "sueldos": "MOD (personal directo)", "salarios": "MOD (personal directo)",
    "planilla": "MOD (personal directo)",
    "proceso": "actividad", "procesos": "actividades", "tarea": "actividad",
    "etapa": "actividad", "operacion": "actividad",
    "gastos indirectos": "recursos indirectos (CIF)", "indirectos": "recursos indirectos (CIF)",
    "cif": "recursos indirectos (CIF)",
    "terceros": "servicio directo (CSD)",
    "prorrateo": "inductor / peso de distribución", "criterio": "inductor",
    "mas": "criterio: máximo", "mayor": "criterio: máximo", "mejor": "criterio: máximo",
    "menos": "criterio: mínimo", "menor": "criterio: mínimo", "peor": "criterio: mínimo",
}

# Verbos que indican que el usuario pide MODIFICAR algo. El agente es de solo
# lectura: los detecta para responder que no realiza cambios.
VERBOS_CAMBIO = {
    "modifica", "modificar", "modifique", "modifiquen", "cambia", "cambiar", "cambie",
    "actualiza", "actualizar", "actualice", "borra", "borrar", "borre", "elimina",
    "eliminar", "elimine", "edita", "editar", "edite", "inserta", "insertar", "agrega",
    "agregar", "agregue", "anade", "anadir", "registra", "registrar", "registre",
    "importa", "importar", "importe", "ejecuta", "ejecutar", "recalcula", "recalcular",
    "guarda", "guardar", "sobrescribe", "sobrescribir", "reemplaza", "reemplazar",
    "corrige", "corregir", "ajusta", "ajustar", "pon", "poner", "fija", "fijar",
    "sube", "baja", "aumenta", "disminuye", "reduce", "crea", "crear", "cree",
    "exporta", "genera", "generar", "compila", "compilar", "cierra", "cerrar",
}

MARCAS_PREGUNTA = {
    "como", "deberia", "deberiamos", "conviene", "convendria", "pasaria", "seria",
    "podria", "podriamos", "propuesta", "propuestas", "recomiendas", "recomendacion",
    "sugieres", "sugerencia", "que", "cual", "cuanto", "por", "puede", "puedo",
    "pueden", "debo", "debemos",
}


# Adjetivos que ya traen el criterio: "mas barato" = el menor costo unitario.
ADJETIVOS_CON_CRITERIO = {"barato", "barata", "baratos", "caro", "cara", "caros", "bajo", "baja", "bajos",
                          "alto", "alta", "altos", "peor", "mejor", "elevado", "elevada", "reducido"}
INVERSO = {"barato": "caro", "barata": "cara", "baratos": "caros", "caro": "barato", "cara": "barata",
           "caros": "baratos", "bajo": "alto", "baja": "alta", "bajos": "altos", "alto": "bajo", "alta": "baja",
           "altos": "bajos", "peor": "mejor", "mejor": "peor", "elevado": "reducido", "elevada": "reducido",
           "reducido": "elevado"}


def preprocesar(tokens: list[str]) -> list[str]:
    """'mas barato' -> 'barato'; 'menos caro' -> 'barato'."""
    salida: list[str] = []
    i = 0
    while i < len(tokens):
        tok = tokens[i]
        sig = tokens[i + 1] if i + 1 < len(tokens) else ""
        if tok in {"mas", "menos"} and sig in ADJETIVOS_CON_CRITERIO:
            salida.append(sig if tok == "mas" else INVERSO.get(sig, sig))
            i += 2
            continue
        salida.append(tok)
        i += 1
    return salida


# Reglas de composicion: combinaciones que equivalen a otro concepto.
def componer(conceptos: set[str]) -> set[str]:
    extra = set()
    if "MARGEN" in conceptos and "PORCENTAJE" in conceptos:
        extra.add("RENTABILIDAD")
    return conceptos | extra


@dataclass(frozen=True)
class TerminoReconocido:
    texto: str
    conceptos: tuple[str, ...]


def marcar_frases(tokens: list[str], usados: list[bool]) -> tuple[set[str], list[TerminoReconocido]]:
    """Primera pasada: frases de varias palabras (la mas larga primero)."""
    conceptos: set[str] = set()
    reconocidos: list[TerminoReconocido] = []
    frases = sorted(FRASES, key=lambda item: len(item[0].split()), reverse=True)
    for frase, tags in frases:
        partes = frase.split()
        n = len(partes)
        for i in range(len(tokens) - n + 1):
            if any(usados[i : i + n]):
                continue
            if tokens[i : i + n] == partes:
                conceptos.update(tags)
                reconocidos.append(TerminoReconocido(frase, tags))
                for j in range(i, i + n):
                    usados[j] = True
    return conceptos, reconocidos


def marcar_palabras(tokens: list[str], usados: list[bool]) -> tuple[set[str], list[TerminoReconocido]]:
    """Ultima pasada: palabras sueltas y raices."""
    conceptos: set[str] = set()
    reconocidos: list[TerminoReconocido] = []
    for i, tok in enumerate(tokens):
        if usados[i]:
            continue
        tags = PALABRAS.get(tok)
        if tags is None:
            for pref, ptags in PREFIJOS:
                if tok.startswith(pref):
                    tags = ptags
                    break
        if tags:
            conceptos.update(tags)
            reconocidos.append(TerminoReconocido(tok, tags))
            usados[i] = True
    return conceptos, reconocidos


def extraer_conceptos(tokens: list[str], usados: list[bool] | None = None) -> tuple[set[str], list[TerminoReconocido]]:
    """Convierte tokens normalizados en conceptos (frases y luego palabras)."""
    usados = usados if usados is not None else [False] * len(tokens)
    c1, r1 = marcar_frases(tokens, usados)
    c2, r2 = marcar_palabras(tokens, usados)
    return c1 | c2, r1 + r2


def nombre_concepto(concepto: str) -> str:
    return CONCEPTOS.get(concepto, (concepto.lower(), 1.0))[0]


def peso_concepto(concepto: str) -> float:
    return CONCEPTOS.get(concepto, ("", 1.0))[1]
