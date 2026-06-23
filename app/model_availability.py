from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from .constants import MODEL_1_VISUAL_NAME, MODEL_2_BLOCK_MESSAGE, MODEL_3_BLOCK_MESSAGE
from .database import db_session, fetch_one_value, quote_identifier


@dataclass
class CostingModelStatus:
    code: str
    label: str
    visual_name: str
    status: str
    available: bool
    reason: str
    flow: str


def _has_rows(conn, table: str) -> bool:
    return bool(fetch_one_value(conn, f"SELECT COUNT(*) FROM {quote_identifier(table)}", default=0))


def evaluate_models(db_path: str | Path) -> list[CostingModelStatus]:
    with db_session(db_path) as conn:
        model_1_requirements = {
            "actividades": _has_rows(conn, "actividades"),
            "elementos de costo": _has_rows(conn, "costo_otros_recursos(CI)"),
            "inductores EC-AC": _has_rows(conn, "ec_ac_inductores"),
            "ordenes/objetos de costo": _has_rows(conn, "ordenes_trabajo") and _has_rows(conn, "objetos_costo"),
            "inductores AC-OC/OT": _has_rows(conn, "ac_ot_pesos_produccion"),
        }
        missing_1 = [name for name, present in model_1_requirements.items() if not present]
        model_1_available = not missing_1

    reason_1 = (
        "Disponible: el Excel importado contiene actividades, elementos de costo, inductores EC-AC, "
        "ordenes de trabajo/objetos de costo e inductores AC-OC/OT."
        if model_1_available
        else "No disponible: faltan " + ", ".join(missing_1) + "."
    )
    return [
        CostingModelStatus(
            code="MODELO_1",
            label="Modelo 1",
            visual_name=MODEL_1_VISUAL_NAME,
            status="Disponible" if model_1_available else "No disponible por datos faltantes",
            available=model_1_available,
            reason=reason_1,
            flow="Proceso 1: EC -> AC; Proceso 2 AC-OT (Produccion); Proceso 2 AC-OT (Soporte)",
        ),
        CostingModelStatus(
            code="MODELO_2",
            label="Modelo 2",
            visual_name="Modelo 2 teorico",
            status="No disponible por datos faltantes",
            available=False,
            reason=MODEL_2_BLOCK_MESSAGE,
            flow="EC -> CA -> AC -> OC",
        ),
        CostingModelStatus(
            code="MODELO_3",
            label="Modelo 3",
            visual_name="Modelo 3 teorico",
            status="No disponible por datos faltantes",
            available=False,
            reason=MODEL_3_BLOCK_MESSAGE,
            flow="Ruta escalonada de centros y actividades de apoyo, operacion y estrategia",
        ),
    ]
