from __future__ import annotations

import math
import unicodedata
from collections import defaultdict
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Iterable


def normalize_header(value: Any) -> str:
    text = "" if value is None else str(value).strip()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return text.lower()


def safe_float(value: Any, default: float = 0.0) -> float:
    if value is None or value == "":
        return default
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return default
        return float(value)
    text = str(value).strip().replace(",", "")
    try:
        return float(text)
    except ValueError:
        return default


def safe_int(value: Any, default: int = 0) -> int:
    try:
        return int(round(safe_float(value, default)))
    except (TypeError, ValueError):
        return default


def round_half_up(value: float, decimals: int = 0) -> float:
    quant = Decimal("1") if decimals == 0 else Decimal("1").scaleb(-decimals)
    return float(Decimal(str(value)).quantize(quant, rounding=ROUND_HALF_UP))


def round_percentage_fraction(value: float, decimals_as_percentage: int) -> float:
    return round_half_up(value * 100.0, decimals_as_percentage) / 100.0


def is_blank_row(row: dict[str, Any]) -> bool:
    return all(value is None or str(value).strip() == "" for value in row.values())


def group_sum(rows: Iterable[dict[str, Any]], key: str, value: str) -> dict[str, float]:
    totals: dict[str, float] = defaultdict(float)
    for row in rows:
        totals[str(row.get(key) or "")] += safe_float(row.get(value))
    return dict(totals)


def ensure_dirs(*paths: str | Path) -> None:
    for path in paths:
        Path(path).mkdir(parents=True, exist_ok=True)


def money(value: float) -> str:
    return f"S/ {value:,.2f}"


def pct(value: float) -> str:
    return f"{value * 100:,.2f}%"



# ---------------------------------------------------------------------------
# Objeto de costo = Orden de trabajo (OT)
# ---------------------------------------------------------------------------
# En el modelo, cada OT es el objeto de costo. Estas funciones solo arman
# etiquetas para pantallas y reportes; no intervienen en ningun calculo.

def codigos_ot(conn) -> dict[str, str]:
    """ot_id -> codigo_ot leido de la tabla ordenes_trabajo."""
    try:
        rows = conn.execute("SELECT ot_id, codigo_ot FROM ordenes_trabajo").fetchall()
    except Exception:
        return {}
    codigos = {row[0]: row[1] or row[0] for row in rows}
    try:  # consolidado de varios meses: las OT del periodo vienen en la hoja de costos de la ultima ejecucion
        for row in conn.execute(
            """SELECT ot_id, codigo_ot FROM hoja_costos_ot WHERE run_id = (
                   SELECT MAX(run_id) FROM calculation_runs WHERE status = 'COMPLETADO')"""
        ).fetchall():
            codigos.setdefault(row[0], row[1] or row[0])
    except Exception:
        pass
    return codigos


def con_codigo_ot(rows: list[dict], codigos: dict[str, str]) -> list[dict]:
    """Agrega 'codigo_ot' (despues de 'ot_id') a filas que solo traen ot_id."""
    salida = []
    for row in rows:
        if "ot_id" not in row or row.get("codigo_ot"):
            salida.append(row)
            continue
        nueva: dict = {}
        for key, value in row.items():
            nueva[key] = value
            if key == "ot_id":
                nueva["codigo_ot"] = codigos.get(value, value)
        salida.append(nueva)
    return salida


def etiqueta_objeto_costo(row: dict | None, con_producto: bool = True) -> str:
    """Etiqueta del objeto de costo: la OT, con el producto que fabrica entre parentesis."""
    if not row:
        return ""
    codigo = row.get("codigo_ot") or row.get("ot_id") or row.get("objeto_costo_id") or ""
    producto = row.get("codigo_producto")
    if con_producto and producto and producto != codigo:
        return f"{codigo} ({producto})"
    return str(codigo)
