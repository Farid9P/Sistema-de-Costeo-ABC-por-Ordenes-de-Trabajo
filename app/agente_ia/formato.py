"""Formato de cifras y tablas de texto para las respuestas del agente."""

from __future__ import annotations

from dataclasses import dataclass, field


def num(value) -> float:
    try:
        if value is None or value == "":
            return 0.0
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def soles(value) -> str:
    return f"S/ {num(value):,.2f}"


def pct(value, decimals: int = 2) -> str:
    return f"{num(value) * 100:,.{decimals}f} %"


def cifra(value, decimals: int = 2) -> str:
    return f"{num(value):,.{decimals}f}"


def tabla(filas: list[dict], columnas: list[tuple[str, str]], total: dict | None = None) -> str:
    """Tabla de texto monoespaciada. ``columnas`` = [(clave, titulo)]."""
    if not filas:
        return "(sin filas)"
    datos = [[str(fila.get(clave, "")) for clave, _ in columnas] for fila in filas]
    if total:
        datos.append([str(total.get(clave, "")) for clave, _ in columnas])
    titulos = [titulo for _, titulo in columnas]
    anchos = [max(len(titulos[i]), *(len(row[i]) for row in datos)) for i in range(len(columnas))]

    def es_numero(texto: str) -> bool:
        limpio = texto.replace("S/", "").replace(",", "").replace("%", "").replace(" ", "").replace("-", "")
        return limpio.replace(".", "", 1).isdigit()

    def linea(row: list[str]) -> str:
        celdas = []
        for i, celda in enumerate(row):
            celdas.append(celda.rjust(anchos[i]) if es_numero(celda) else celda.ljust(anchos[i]))
        return "  ".join(celdas).rstrip()

    sep = "  ".join("-" * a for a in anchos)
    salida = [linea(titulos), sep]
    cuerpo = datos[:-1] if total else datos
    salida += [linea(row) for row in cuerpo]
    if total:
        salida += [sep, linea(datos[-1])]
    return "\n".join(salida)


@dataclass
class Respuesta:
    titulo: str
    resumen: str
    detalle: str = ""
    formula: str = ""
    fuente: str = ""
    periodo: str = ""
    observaciones: list[str] = field(default_factory=list)
    sin_datos: bool = False

    def como_texto(self) -> str:
        partes = [self.resumen]
        if self.detalle:
            partes += ["", self.detalle]
        if self.formula:
            partes += ["", f"Fórmula: {self.formula}"]
        if self.periodo:
            partes.append(f"Periodo: {self.periodo}")
        if self.fuente:
            partes.append(f"Fuente: {self.fuente}")
        if self.observaciones:
            partes.append("Observaciones:")
            partes += [f"  - {obs}" for obs in self.observaciones]
        return "\n".join(partes)
