from __future__ import annotations

import sqlite3

from .utils import safe_float


def calculate_average_unit_cost(total_cost: float, total_quantity: float) -> float:
    if total_quantity == 0:
        return 0.0
    return total_cost / total_quantity


def build_kardex_from_materiales(conn: sqlite3.Connection, run_id: int) -> int:
    conn.execute("DELETE FROM kardex_promedio")
    rows = conn.execute("SELECT * FROM materiales_directos").fetchall()
    count = 0
    for row in rows:
        qty = safe_float(row["cantidad_kg"])
        unit = safe_float(row["costo_unitario_soles"])
        total = qty * unit
        average = calculate_average_unit_cost(total, qty)
        conn.execute(
            """
            INSERT INTO kardex_promedio
                (run_id, material_id, ot_id, material, entrada_cantidad,
                 entrada_costo_unitario, costo_promedio, salida_cantidad,
                 costo_salida, formula, observacion)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                run_id,
                row["material_id"],
                row["ot_id"],
                row["material"],
                qty,
                unit,
                average,
                qty,
                qty * average,
                "Costo promedio = costo total disponible / cantidad total disponible",
                "El Excel trae costo unitario publicado; se registra como costo promedio validado para la salida.",
            ),
        )
        count += 1
    return count

