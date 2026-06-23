from __future__ import annotations

import sqlite3

from .utils import safe_float


def calculate_cmd(conn: sqlite3.Connection, run_id: int) -> int:
    conn.execute("DELETE FROM cmd_ot")
    rows = conn.execute("SELECT * FROM materiales_directos").fetchall()
    for row in rows:
        qty = safe_float(row["cantidad_kg"])
        unit_cost = safe_float(row["costo_unitario_soles"])
        cmd = qty * unit_cost
        conn.execute(
            """
            INSERT INTO cmd_ot
                (run_id, ot_id, objeto_costo_id, material_id, material, cantidad_kg,
                 costo_unitario_promedio, cmd_soles, formula, fuente, observacion)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                run_id,
                row["ot_id"],
                row["objeto_costo_id"],
                row["material_id"],
                row["material"],
                qty,
                unit_cost,
                cmd,
                "CMD_OT = cantidad_consumida * costo_unitario_promedio",
                row["fuente"],
                row["observacion"],
            ),
        )
    return len(rows)


def calculate_csd(conn: sqlite3.Connection, run_id: int) -> int:
    conn.execute("DELETE FROM csd_ot")
    rows = conn.execute("SELECT * FROM servicios_directos WHERE servicio_id IS NOT NULL AND TRIM(servicio_id) <> ''").fetchall()
    if rows:
        for row in rows:
            qty = safe_float(row["cantidad"])
            unit_cost = safe_float(row["costo_unitario_soles"])
            conn.execute(
                """
                INSERT INTO csd_ot
                    (run_id, ot_id, servicio_id, servicio_directo, cantidad,
                     costo_unitario_soles, csd_soles, formula, fuente, observacion)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    run_id,
                    row["ot_id"],
                    row["servicio_id"],
                    row["servicio_directo"],
                    qty,
                    unit_cost,
                    qty * unit_cost,
                    "CSD_OT = cantidad * costo_unitario_soles",
                    row["fuente"],
                    row["observacion"],
                ),
            )
        return len(rows)

    ots = conn.execute("SELECT ot_id FROM ordenes_trabajo").fetchall()
    for row in ots:
        conn.execute(
            """
            INSERT INTO csd_ot
                (run_id, ot_id, servicio_id, servicio_directo, cantidad,
                 costo_unitario_soles, csd_soles, formula, fuente, observacion)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                run_id,
                row["ot_id"],
                None,
                None,
                0,
                0,
                0,
                "CSD_OT = 0 cuando no existen servicios directos publicados",
                "frenosa_datos_entrada_abc.xlsx",
                "Sin datos de servicio directo publicado en la fuente.",
            ),
        )
    return len(ots)


def calculate_direct_costs(conn: sqlite3.Connection, run_id: int) -> dict[str, int]:
    return {
        "cmd_ot": calculate_cmd(conn, run_id),
        "csd_ot": calculate_csd(conn, run_id),
    }

