"""Interprete de preguntas de Finanzas y Contabilidad.

Pasos:
1. Normaliza la pregunta (minusculas, sin tildes ni signos).
2. Reconoce frases tecnicas, entidades de la base (OT, actividades, recursos)
   y palabras/raices, y las convierte en conceptos (lexico.py).
3. Compara esos conceptos con cada pregunta del catalogo base (txt) y sus
   ejemplos usando un F1 ponderado (peso del concepto x rareza en el catalogo).
4. Reformula la pregunta con vocabulario tecnico del Modelo ABC por OT, por
   ejemplo: "¿cual es producto con mas ventas?" ->
   "¿Cuál es el objeto de costo que genera más ingresos?".

Funciona sin internet y sin modelos externos, para poder ir dentro del .exe.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field

from . import lexico
from .catalogo import SLOTS, PreguntaBase, cargar_catalogo

UMBRAL_SEGURO = 0.55
UMBRAL_MINIMO = 0.30

_MARCADORES = {"zzentot": "ENT_OT", "zzentactividad": "ENT_ACTIVIDAD", "zzentrecurso": "ENT_RECURSO"}
_BASE_DE_ENTIDAD = {"ENT_OT": "OBJETO", "ENT_ACTIVIDAD": "ACTIVIDAD", "ENT_RECURSO": "RECURSO"}
_NO_PUNTUAN = {"LISTA", "IMPORTE"}


# ---------------------------------------------------------------------------
# Entidades leidas de la base (no hay datos de la empresa en el codigo)
# ---------------------------------------------------------------------------

@dataclass
class Entidad:
    tipo: str  # OT | ACTIVIDAD | RECURSO
    id: str
    nombre: str
    alias: set[str] = field(default_factory=set)
    raiz: str = ""
    extra: dict = field(default_factory=dict)

    @property
    def concepto(self) -> str:
        return {"OT": "ENT_OT", "ACTIVIDAD": "ENT_ACTIVIDAD", "RECURSO": "ENT_RECURSO"}[self.tipo]


class CatalogoEntidades:
    def __init__(self, entidades: list[Entidad] | None = None):
        self.entidades = entidades or []
        self._alias: dict[str, Entidad] = {}
        conteo: dict[str, set[str]] = {}
        for ent in self.entidades:
            for alias in ent.alias:
                conteo.setdefault(alias, set()).add(f"{ent.tipo}:{ent.id}")
        for ent in self.entidades:
            for alias in ent.alias:
                if len(conteo[alias]) == 1 and len(alias) >= 3:
                    self._alias[alias] = ent

    @classmethod
    def desde_base(cls, lector) -> "CatalogoEntidades":
        entidades: list[Entidad] = []
        if lector is None or not lector.disponible:
            return cls([])
        objetos = {row["objeto_costo_id"]: row for row in lector.tabla("objetos_costo")}
        materiales: dict = {}
        codigo_a_ot = {str(ot.get("codigo_ot") or "").strip().lower(): ot.get("ot_id") for ot in lector.tabla("ordenes_trabajo")}
        compras = lector.tabla("compras") if lector.existe_tabla("compras") else []
        for row in compras:
            destino = str(row.get("destino") or "").strip().lower()
            ot_id = codigo_a_ot.get(destino)
            if ot_id and row.get("item_descripcion"):
                lista = materiales.setdefault(ot_id, [])
                if all(m.get("material") != row["item_descripcion"] for m in lista):
                    lista.append({"material": row["item_descripcion"]})
        for ot in lector.tabla("ordenes_trabajo", order="ot_id"):
            obj = objetos.get(ot.get("objeto_costo_id"), {})
            nombre = f"{ot.get('codigo_ot') or ot.get('ot_id')} ({obj.get('codigo_producto') or ot.get('objeto_costo_id')})"
            alias: set[str] = set()
            valores = [ot.get("ot_id"), ot.get("codigo_ot"), ot.get("objeto_costo_id"), obj.get("codigo_producto"),
                       obj.get("descripcion_producto"), obj.get("material_base")]
            valores += [m.get("material") for m in materiales.get(ot.get("ot_id"), [])]
            for valor in valores:
                if not valor:
                    continue
                alias.add(lexico.compact(valor))
                for parte in lexico.normalize(valor).split():
                    if any(ch.isdigit() for ch in parte) and len(parte) >= 3:
                        alias.add(lexico.compact(parte))
                        digitos = "".join(ch for ch in parte if ch.isdigit())
                        if len(digitos) >= 4:
                            alias.add(digitos)
                partes = lexico.normalize(valor).split()
                for i in range(len(partes) - 1):
                    par = partes[i] + partes[i + 1]
                    if any(ch.isdigit() for ch in par):
                        alias.add(par)
            entidades.append(Entidad("OT", ot["ot_id"], nombre, alias, extra={"codigo_ot": ot.get("codigo_ot"),
                                                                                 "codigo_producto": obj.get("codigo_producto")}))
        for act in lector.tabla("actividades", order="actividad_id"):
            nombre_norm = lexico.normalize(act.get("nombre_actividad"))
            alias = {lexico.compact(act.get("actividad_id")), lexico.compact(act.get("codigo_actividad")),
                     lexico.compact(nombre_norm)}
            raiz = nombre_norm.replace(" ", "")
            raiz = raiz[: max(4, len(raiz) - 3)] if " " not in nombre_norm else ""
            entidades.append(Entidad("ACTIVIDAD", act["actividad_id"],
                                     f"{act.get('codigo_actividad')} {act.get('nombre_actividad')}", alias, raiz))
        for rec in lector.tabla("costo_otros_recursos(CI)", order="elemento_costo_id"):
            alias = {lexico.compact(rec.get("elemento_costo_id")), lexico.compact(rec.get("codigo_recurso")),
                     lexico.compact(rec.get("nombre_recurso"))}
            entidades.append(Entidad("RECURSO", rec["elemento_costo_id"], rec.get("nombre_recurso") or rec["elemento_costo_id"],
                                     alias))
        return cls(entidades)

    def buscar(self, tokens: list[str], usados: list[bool]) -> list[tuple[Entidad, str]]:
        """Marca en ``usados`` los tokens que nombran una entidad."""
        encontradas: list[tuple[Entidad, str]] = []
        for n in (4, 3, 2, 1):
            for i in range(len(tokens) - n + 1):
                if any(usados[i : i + n]):
                    continue
                clave = "".join(tokens[i : i + n])
                ent = self._alias.get(clave)
                if ent is None and n == 1:
                    ent = self._por_raiz(clave)
                if ent is not None:
                    encontradas.append((ent, " ".join(tokens[i : i + n])))
                    for j in range(i, i + n):
                        usados[j] = True
        unicas: dict[str, tuple[Entidad, str]] = {}
        for ent, texto in encontradas:
            unicas.setdefault(f"{ent.tipo}:{ent.id}", (ent, texto))
        return list(unicas.values())

    def _por_raiz(self, token: str) -> Entidad | None:
        if len(token) < 4:
            return None
        candidatas = [e for e in self.entidades if e.raiz and token.startswith(e.raiz)]
        return candidatas[0] if len(candidatas) == 1 else None


# ---------------------------------------------------------------------------
# Resultado de la interpretacion
# ---------------------------------------------------------------------------

@dataclass
class Interpretacion:
    original: str
    interpretada: str
    pregunta: PreguntaBase | None
    confianza: float
    conceptos: set[str]
    entidades: list[Entidad]
    terminos: list[tuple[str, str]]
    direccion: str | None  # "MAX" | "MIN" | None
    alternativas: list[tuple[PreguntaBase, float]]
    es_pedido_de_cambio: bool = False

    @property
    def entendida(self) -> bool:
        return self.pregunta is not None and self.confianza >= UMBRAL_MINIMO

    @property
    def segura(self) -> bool:
        return self.pregunta is not None and self.confianza >= UMBRAL_SEGURO

    def entidades_de(self, tipo: str) -> list[Entidad]:
        return [e for e in self.entidades if e.tipo == tipo]


# ---------------------------------------------------------------------------
# Interprete
# ---------------------------------------------------------------------------

def _texto_con_marcadores(texto: str) -> str:
    for slot, _concepto in SLOTS.items():
        texto = texto.replace(slot, " zz" + "ent" + slot.strip("{}").lower() + " ")
    return texto


def _conceptos_de_texto(texto: str) -> set[str]:
    tokens = lexico.tokenize(_texto_con_marcadores(texto))
    usados = [False] * len(tokens)
    conceptos: set[str] = set()
    for i, tok in enumerate(tokens):
        if tok in _MARCADORES:
            conceptos.add(_MARCADORES[tok])
            usados[i] = True
    extra, _ = lexico.extraer_conceptos(tokens, usados)
    return conceptos | extra


def _para_puntaje(conceptos: set[str]) -> set[str]:
    salida = set()
    for c in conceptos:
        if c in _NO_PUNTUAN:
            continue
        salida.add(_BASE_DE_ENTIDAD.get(c, c))
    return salida


class InterpreteFinanzas:
    def __init__(self, catalogo: list[PreguntaBase] | None = None):
        self.catalogo = catalogo or cargar_catalogo()
        self._firmas: list[tuple[PreguntaBase, list[set[str]], set[str]]] = []
        df: dict[str, int] = {}
        for pregunta in self.catalogo:
            textos = [pregunta.texto] + pregunta.ejemplos
            conjuntos = [_para_puntaje(_conceptos_de_texto(t)) for t in textos]
            union = set().union(*conjuntos)
            self._firmas.append((pregunta, conjuntos, union))
            for c in union:
                df[c] = df.get(c, 0) + 1
        n = len(self.catalogo)
        self._peso = {c: lexico.peso_concepto(c) * (1.0 + math.log(n / df[c])) for c in df}

    # -- puntaje -----------------------------------------------------------
    def _w(self, conceptos: set[str]) -> float:
        return sum(self._peso.get(c, lexico.peso_concepto(c) * 1.5) for c in conceptos)

    def _f1(self, q: set[str], t: set[str]) -> float:
        if not q or not t:
            return 0.0
        inter = self._w(q & t)
        if inter == 0:
            return 0.0
        p = inter / self._w(q)
        r = inter / self._w(t)
        return 2 * p * r / (p + r)

    RANKING_PURO = {"27", "36", "41"}

    def puntuar(self, conceptos: set[str], entidades: list[Entidad],
                direccion: str | None = None) -> list[tuple[PreguntaBase, float]]:
        q = _para_puntaje(conceptos)
        tipos = {e.tipo for e in entidades}
        resultados = []
        for pregunta, conjuntos, _union in self._firmas:
            # el texto del txt pesa un poco mas que los ejemplos (desempate)
            score = max([self._f1(q, conjuntos[0]) + 0.002] + [self._f1(q, t) for t in conjuntos[1:]])
            requeridos = {{"{OT}": "OT", "{actividad}": "ACTIVIDAD", "{recurso}": "RECURSO"}[slot] for slot in pregunta.slots}
            if requeridos and requeridos <= tipos:
                score += 0.03 * len(requeridos)
            # una pregunta que exige actividad o recurso pierde peso si no se menciono ninguno
            for tipo, concepto in (("ACTIVIDAD", "ACTIVIDAD"), ("RECURSO", "RECURSO")):
                if tipo in requeridos and tipo not in tipos and concepto not in q:
                    score *= 0.85
            # superlativo sin entidad: preferir preguntas que pueden ordenar (ranking)
            if direccion and not entidades and (pregunta.ranking_max or pregunta.codigo in self.RANKING_PURO):
                score += 0.05
            # entidad concreta sin superlativo: no es una pregunta de ranking
            if entidades and not direccion and pregunta.codigo in self.RANKING_PURO:
                score *= 0.8
            if pregunta.dominante and pregunta.dominante in q:
                score = max(score, 0.9)
            resultados.append((pregunta, score))
        resultados.sort(key=lambda item: item[1], reverse=True)
        return [(p, min(s, 1.0)) for p, s in resultados]

    # -- interpretacion -----------------------------------------------------
    def interpretar(self, texto: str, entidades_catalogo: CatalogoEntidades | None = None) -> Interpretacion:
        entidades_catalogo = entidades_catalogo or CatalogoEntidades([])
        tokens_crudos = lexico.tokenize(texto)
        visibles = lexico.tokenize_display(texto)
        mapa_visible = dict(zip(tokens_crudos, visibles))
        tokens = lexico.preprocesar(tokens_crudos)
        usados = [False] * len(tokens)
        c_frases, r_frases = lexico.marcar_frases(tokens, usados)
        halladas = entidades_catalogo.buscar(tokens, usados)
        c_palabras, r_palabras = lexico.marcar_palabras(tokens, usados)
        conceptos = lexico.componer(c_frases | c_palabras | {ent.concepto for ent, _ in halladas})
        # "caro/barato" es costo unitario para productos, pero costo a secas para recursos o actividades
        if ({"RECURSO", "ACTIVIDAD", "ENT_RECURSO", "ENT_ACTIVIDAD"} & conceptos and "OBJETO" not in conceptos
                and any(t in lexico.ADJETIVOS_CON_CRITERIO for t in tokens)
                and not any(t.startswith("unitari") for t in tokens)):
            conceptos = (conceptos - {"UNITARIO"}) | {"COSTO"}
        entidades = [ent for ent, _ in halladas]

        direccion = None
        if "MAX" in conceptos and "MIN" not in conceptos:
            direccion = "MAX"
        elif "MIN" in conceptos and "MAX" not in conceptos:
            direccion = "MIN"

        ranking = self.puntuar(conceptos, entidades, direccion)
        mejor, score = ranking[0] if ranking else (None, 0.0)
        if score < UMBRAL_MINIMO:
            mejor = None
        alternativas = [(p, s) for p, s in ranking[1:6] if s >= max(UMBRAL_MINIMO, score * 0.75)][:3]

        terminos = self._terminos(r_frases + r_palabras, halladas, mejor, mapa_visible)
        interpretada = self.reformular(mejor, entidades, direccion, conceptos) if mejor else ""
        return Interpretacion(
            original=texto,
            interpretada=interpretada,
            pregunta=mejor,
            confianza=score,
            conceptos=conceptos,
            entidades=entidades,
            terminos=terminos,
            direccion=direccion,
            alternativas=alternativas,
            es_pedido_de_cambio=self.es_pedido_de_cambio(tokens_crudos),
        )

    def _terminos(self, reconocidos, halladas, pregunta: PreguntaBase | None,
                  visible: dict[str, str] | None = None) -> list[tuple[str, str]]:
        visible = visible or {}
        salida: list[tuple[str, str]] = []
        vistos = set()
        for termino in reconocidos:
            texto = termino.texto
            tecnico = lexico.GLOSARIO_VISIBLE.get(texto)
            if texto in {"costos", "gastos", "gasto"} and not (pregunta and pregunta.codigo == "E1"):
                tecnico = "costo"
            if tecnico and texto not in vistos and lexico.normalize(tecnico) != texto:
                mostrado = " ".join(visible.get(t, t) for t in texto.split())
                salida.append((mostrado, tecnico))
                vistos.add(texto)
        for ent, texto in halladas:
            salida.append((" ".join(visible.get(t, t) for t in texto.split()), ent.nombre))
        return salida

    @staticmethod
    def es_pedido_de_cambio(tokens: list[str]) -> bool:
        if not tokens:
            return False
        significativos = [t for t in tokens if t not in {"por", "favor", "porfa", "ahora", "ya", "oye", "hola", "agente"}]
        if not significativos:
            return False
        primero = significativos[0]
        if primero in lexico.VERBOS_CAMBIO:
            return True
        texto = " ".join(significativos)
        for intro in ("quiero que", "necesito que", "puedes", "podrias", "me puedes", "me podrias", "hazme", "anda y"):
            if texto.startswith(intro):
                resto = texto[len(intro):].split()
                resto = [t for t in resto if t not in {"me", "lo", "la", "los", "las", "le", "les"}]
                if resto and (resto[0] in lexico.VERBOS_CAMBIO or resto[0].rstrip("s") in lexico.VERBOS_CAMBIO):
                    return True
        return False

    # -- reformulacion --------------------------------------------------------
    @staticmethod
    def reformular(pregunta: PreguntaBase, entidades: list[Entidad], direccion: str | None,
                   conceptos: set[str]) -> str:
        tipos = {e.tipo for e in entidades}
        plantilla = pregunta.ranking_max if direccion == "MAX" else pregunta.ranking_min if direccion == "MIN" else None
        if plantilla:
            sujeto = "ACTIVIDAD" if "actividad" in plantilla.lower() else "RECURSO" if "recurso" in plantilla.lower() else "OT"
            if sujeto not in tipos:
                return plantilla
        texto = pregunta.texto
        reemplazos = {"{OT}": "OT", "{actividad}": "ACTIVIDAD", "{recurso}": "RECURSO"}
        genericos = {"{OT}": "cada OT", "{actividad}": "cada actividad", "{recurso}": "cada recurso"}
        for slot, tipo in reemplazos.items():
            if slot in texto:
                nombres = [e.nombre for e in entidades if e.tipo == tipo]
                texto = texto.replace(slot, " y ".join(nombres) if nombres else genericos[slot])
        for tipo, generico in (("OT", "cada OT"), ("ACTIVIDAD", "cada actividad"), ("RECURSO", "cada recurso")):
            nombres = [e.nombre for e in entidades if e.tipo == tipo]
            if nombres and generico in texto:
                texto = texto.replace(generico, " y ".join(nombres))
        if direccion and "mayor o menor" in texto:
            texto = texto.replace("mayor o menor", "mayor" if direccion == "MAX" else "menor")
        if "total o unitario" in texto:
            if "UNITARIO" in conceptos and "TOTAL" not in conceptos:
                texto = texto.replace("total o unitario", "unitario")
            elif "TOTAL" in conceptos and "UNITARIO" not in conceptos:
                texto = texto.replace("total o unitario", "total")
            if direccion == "MIN":
                texto = texto.replace("el mayor", "el menor")
        texto = texto.replace("Qué producto", "Qué objeto de costo").replace(" producto ", " objeto de costo ")
        if direccion == "MIN" and pregunta.codigo == "27":
            texto = texto.replace("mayor costo", "menor costo")
        if direccion == "MIN" and pregunta.codigo == "35":
            texto = texto.replace("mayor participación", "menor participación")
        return texto
