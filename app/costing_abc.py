from __future__ import annotations

import sqlite3
from collections import defaultdict
from pathlib import Path

from .costing_direct import calculate_direct_costs
from .costing_mod import calculate_mod_to_activities
from .database import (
    clear_calculated_tables,
    db_session,
    finish_calculation_run,
    initialize_database,
    start_calculation_run,
)
from .inventory_average import build_kardex_from_compras
from .model_availability import evaluate_models
from .utils import safe_float


def calculate_otros_recursos_to_activities(conn: sqlite3.Connection, run_id: int) -> int:
    conn.execute("DELETE FROM otros_recursos_ac_detalle")
    conn.execute("DELETE FROM otros_recursos_actividad")
    rows = conn.execute(
        """
        SELECT i.*, c.nombre_recurso, c.importe_soles, c.fuente AS recurso_fuente,
               c.observacion AS recurso_observacion
        FROM ec_ac_inductores i
        JOIN "costo_otros_recursos(CI)" c
          ON c.elemento_costo_id = i.elemento_costo_id
        """
    ).fetchall()
    totals: dict[tuple[str, str, str], float] = defaultdict(float)
    for row in rows:
        importe = safe_float(row["importe_soles"])
        porcentaje = safe_float(row["porcentaje_inductor"])
        costo = importe * porcentaje
        key = (row["actividad_id"], row["codigo_actividad"], row["nombre_actividad"])
        totals[key] += costo
        conn.execute(
            """
            INSERT INTO otros_recursos_ac_detalle
                (run_id, elemento_costo_id, codigo_recurso, nombre_recurso,
                 actividad_id, codigo_actividad, nombre_actividad, importe_recurso_soles,
                 porcentaje_inductor, costo_recurso_ac, formula, fuente, observacion)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                run_id,
                row["elemento_costo_id"],
                row["codigo_recurso"],
                row["nombre_recurso"],
                row["actividad_id"],
                row["codigo_actividad"],
                row["nombre_actividad"],
                importe,
                porcentaje,
                costo,
                "costo_recurso_AC = importe_recurso * porcentaje_inductor_recurso_AC",
                row["fuente"],
                row["observacion"],
            ),
        )
    for (actividad_id, codigo, nombre), total in totals.items():
        conn.execute(
            """
            INSERT INTO otros_recursos_actividad
                (run_id, actividad_id, codigo_actividad, nombre_actividad,
                 costo_otros_recursos_ac, formula, fuente, observacion)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                run_id,
                actividad_id,
                codigo,
                nombre,
                total,
                "costo_otros_recursos_AC = SUM(costo_recurso_AC)",
                "EC_AC_Inductores + costo_otros_recursos(CI)",
                "Calculo generado por el sistema.",
            ),
        )
    return len(rows)


def calculate_support_to_activities(conn: sqlite3.Connection, run_id: int) -> int:
    conn.execute("DELETE FROM soporte_mantenimiento_ac")
    conn.execute("DELETE FROM soporte_control_calidad_ac")
    conn.execute("DELETE FROM soporte_ingenieria_ac")
    conn.execute("DELETE FROM soporte_actividad")

    mantenimiento_rows = conn.execute(
        """
        SELECT actividad_id_destino AS actividad_id, codigo_actividad_destino AS codigo_actividad,
               SUM(costo_equipo_soles) AS costo_mantenimiento
        FROM "Soporte_Mantenimiento(CI)"
        GROUP BY actividad_id_destino, codigo_actividad_destino
        """
    ).fetchall()
    mantenimiento = {}
    for row in mantenimiento_rows:
        key = row["actividad_id"]
        mantenimiento[key] = safe_float(row["costo_mantenimiento"])
        conn.execute(
            """
            INSERT INTO soporte_mantenimiento_ac
                (run_id, actividad_id, codigo_actividad, costo_mantenimiento_ac,
                 formula, fuente, observacion)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                run_id,
                row["actividad_id"],
                row["codigo_actividad"],
                mantenimiento[key],
                "costo_mantenimiento_AC = SUM(costo_mantenimiento_equipo asociado a la actividad)",
                "Soporte_Mantenimiento(CI)",
                "Calculo agrupado por actividad destino.",
            ),
        )

    cc_total = safe_float(
        conn.execute(
            'SELECT SUM(cantidad * costo_unitario_soles) FROM "Soporte_Control_Calidad(CI)"'
        ).fetchone()[0]
    )
    ing_total = safe_float(
        conn.execute('SELECT SUM(cantidad * costo_unitario_soles) FROM "Soporte_Ingenieria(CI)"').fetchone()[0]
    )

    cc_by_activity = {}
    ing_by_activity = {}
    dist_rows = conn.execute("SELECT * FROM soporte_cc_ing_inductores").fetchall()
    for row in dist_rows:
        proceso = str(row["proceso_soporte"] or "").lower()
        pct = safe_float(row["porcentaje_distribucion"])
        if "calidad" in proceso:
            costo = cc_total * pct
            cc_by_activity[row["actividad_id"]] = costo
            conn.execute(
                """
                INSERT INTO soporte_control_calidad_ac
                    (run_id, actividad_id, codigo_actividad, nombre_actividad,
                     costo_total_control_calidad, porcentaje_distribucion,
                     costo_control_calidad_ac, formula, fuente, observacion)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    run_id,
                    row["actividad_id"],
                    row["codigo_actividad"],
                    row["nombre_actividad"],
                    cc_total,
                    pct,
                    costo,
                    "costo_control_calidad_AC = costo_control_calidad_total * porcentaje_inductor_CC_AC",
                    row["fuente"],
                    row["observacion"],
                ),
            )
        elif "ingen" in proceso:
            costo = ing_total * pct
            ing_by_activity[row["actividad_id"]] = costo
            conn.execute(
                """
                INSERT INTO soporte_ingenieria_ac
                    (run_id, actividad_id, codigo_actividad, nombre_actividad,
                     costo_total_ingenieria, porcentaje_distribucion,
                     costo_ingenieria_ac, formula, fuente, observacion)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    run_id,
                    row["actividad_id"],
                    row["codigo_actividad"],
                    row["nombre_actividad"],
                    ing_total,
                    pct,
                    costo,
                    "costo_ingenieria_AC = costo_ingenieria_total * porcentaje_inductor_ING_AC",
                    row["fuente"],
                    row["observacion"],
                ),
            )

    actividades = conn.execute("SELECT * FROM actividades").fetchall()
    for row in actividades:
        mant = mantenimiento.get(row["actividad_id"], 0.0)
        cc = cc_by_activity.get(row["actividad_id"], 0.0)
        ing = ing_by_activity.get(row["actividad_id"], 0.0)
        total = mant + cc + ing
        conn.execute(
            """
            INSERT INTO soporte_actividad
                (run_id, actividad_id, codigo_actividad, nombre_actividad,
                 costo_mantenimiento_ac, costo_control_calidad_ac, costo_ingenieria_ac,
                 costo_soporte_ac, formula, fuente, observacion)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                run_id,
                row["actividad_id"],
                row["codigo_actividad"],
                row["nombre_actividad"],
                mant,
                cc,
                ing,
                total,
                "costo_soporte_AC = mantenimiento_AC + control_calidad_AC + ingenieria_AC",
                "Soporte_Mantenimiento(CI) + Soporte_Control_Calidad(CI) + Soporte_Ingenieria(CI)",
                "Calculo generado por el sistema.",
            ),
        )
    return len(actividades)


def calculate_activity_totals(conn: sqlite3.Connection, run_id: int) -> int:
    conn.execute("DELETE FROM actividad_total")
    rows = conn.execute("SELECT * FROM actividades").fetchall()
    for row in rows:
        actividad_id = row["actividad_id"]
        mod = safe_float(
            conn.execute(
                "SELECT SUM(costo_mod_ac) FROM mod_actividad WHERE actividad_id = ?",
                (actividad_id,),
            ).fetchone()[0]
        )
        otros = safe_float(
            conn.execute(
                "SELECT SUM(costo_otros_recursos_ac) FROM otros_recursos_actividad WHERE actividad_id = ?",
                (actividad_id,),
            ).fetchone()[0]
        )
        soporte = safe_float(
            conn.execute(
                "SELECT SUM(costo_soporte_ac) FROM soporte_actividad WHERE actividad_id = ?",
                (actividad_id,),
            ).fetchone()[0]
        )
        conversion = mod + otros
        total = conversion + soporte
        conn.execute(
            """
            INSERT INTO actividad_total
                (run_id, actividad_id, codigo_actividad, nombre_actividad,
                 costo_mod_ac, costo_otros_recursos_ac, costo_conversion_ac,
                 costo_soporte_ac, costo_total_ac, formula, fuente, observacion)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                run_id,
                actividad_id,
                row["codigo_actividad"],
                row["nombre_actividad"],
                mod,
                otros,
                conversion,
                soporte,
                total,
                "costo_total_AC = PD_AC + otros_recursos_AC + soporte_AC",
                "Calculos PD, EC-AC y Soporte-AC",
                "Total por actividad generado por el sistema.",
            ),
        )
    return len(rows)


def calculate_ac_to_ot(conn: sqlite3.Connection, run_id: int) -> int:
    conn.execute("DELETE FROM produccion_ot_detalle")
    conn.execute("DELETE FROM produccion_ot")
    conn.execute("DELETE FROM soporte_ot_detalle")
    conn.execute("DELETE FROM soporte_ot")

    prod_rows = conn.execute(
        """
        SELECT p.*, a.costo_conversion_ac, a.costo_mod_ac, a.costo_otros_recursos_ac
        FROM ac_ot_pesos_produccion p
        JOIN actividad_total a ON a.actividad_id = p.actividad_id AND a.run_id = ?
        """,
        (run_id,),
    ).fetchall()
    prod_totals: dict[str, dict[str, float | str]] = defaultdict(lambda: {"mod": 0.0, "otros": 0.0, "prod": 0.0, "objeto": ""})
    for row in prod_rows:
        weight = safe_float(row["porcentaje_inductor"])
        conversion_ot = safe_float(row["costo_conversion_ac"]) * weight
        mod_ot = safe_float(row["costo_mod_ac"]) * weight
        otros_ot = safe_float(row["costo_otros_recursos_ac"]) * weight
        prod_totals[row["ot_id"]]["objeto"] = row["objeto_costo_id"]
        prod_totals[row["ot_id"]]["mod"] += mod_ot
        prod_totals[row["ot_id"]]["otros"] += otros_ot
        prod_totals[row["ot_id"]]["prod"] += conversion_ot
        conn.execute(
            """
            INSERT INTO produccion_ot_detalle
                (run_id, actividad_id, codigo_actividad, nombre_actividad, ot_id,
                 objeto_costo_id, codigo_producto, peso_ac_ot, costo_conversion_ac,
                 costo_mod_ac, costo_otros_recursos_ac, costo_conversion_ot,
                 costo_mod_ot, costo_otros_recursos_ot, formula, fuente, observacion)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                run_id,
                row["actividad_id"],
                row["codigo_actividad"],
                row["nombre_actividad"],
                row["ot_id"],
                row["objeto_costo_id"],
                row["codigo_producto"],
                weight,
                row["costo_conversion_ac"],
                row["costo_mod_ac"],
                row["costo_otros_recursos_ac"],
                conversion_ot,
                mod_ot,
                otros_ot,
                "costo_produccion_OT = SUM(costo_conversion_AC * peso_AC_OT_produccion)",
                row["fuente"],
                row["observacion"],
            ),
        )
    for ot_id, values in prod_totals.items():
        conn.execute(
            """
            INSERT INTO produccion_ot
                (run_id, ot_id, objeto_costo_id, costo_mod_ot, costo_otros_recursos_ot,
                 costo_produccion_ot, formula, fuente, observacion)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                run_id,
                ot_id,
                values["objeto"],
                values["mod"],
                values["otros"],
                values["prod"],
                "costo_produccion_OT = SUM(costo_actividad_productiva_AC * peso_AC_OT_produccion)",
                "AC_OT_Pesos_Produccion",
                "Incluye MOD y otros recursos indirectos asignados a actividades.",
            ),
        )

    support_rows = conn.execute(
        """
        SELECT p.*, s.costo_soporte_ac
        FROM ac_ot_pesos_soporte p
        JOIN soporte_actividad s ON s.actividad_id = p.actividad_id AND s.run_id = ?
        """,
        (run_id,),
    ).fetchall()
    support_totals: dict[str, dict[str, float | str]] = defaultdict(lambda: {"soporte": 0.0, "objeto": ""})
    for row in support_rows:
        weight = safe_float(row["porcentaje_inductor"])
        support_ot = safe_float(row["costo_soporte_ac"]) * weight
        support_totals[row["ot_id"]]["objeto"] = row["objeto_costo_id"]
        support_totals[row["ot_id"]]["soporte"] += support_ot
        conn.execute(
            """
            INSERT INTO soporte_ot_detalle
                (run_id, actividad_id, codigo_actividad, nombre_actividad, ot_id,
                 objeto_costo_id, codigo_producto, peso_ac_ot, costo_soporte_ac,
                 costo_soporte_ot, formula, fuente, observacion)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                run_id,
                row["actividad_id"],
                row["codigo_actividad"],
                row["nombre_actividad"],
                row["ot_id"],
                row["objeto_costo_id"],
                row["codigo_producto"],
                weight,
                row["costo_soporte_ac"],
                support_ot,
                "costo_soporte_OT = SUM(costo_soporte_AC * peso_AC_OT_soporte)",
                row["fuente"],
                row["observacion"],
            ),
        )
    for ot_id, values in support_totals.items():
        conn.execute(
            """
            INSERT INTO soporte_ot
                (run_id, ot_id, objeto_costo_id, costo_soporte_ot, formula, fuente, observacion)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                run_id,
                ot_id,
                values["objeto"],
                values["soporte"],
                "costo_soporte_OT = SUM(costo_soporte_AC * peso_AC_OT_soporte)",
                "AC_OT_Pesos_Soporte",
                "Costos de mantenimiento, control de calidad e ingenieria transferidos a OT.",
            ),
        )
    return len(prod_rows) + len(support_rows)


def calculate_cost_sheets(conn: sqlite3.Connection, run_id: int) -> int:
    conn.execute("DELETE FROM hoja_costos_ot")
    conn.execute("DELETE FROM estado_resultados_ot")
    ots = conn.execute(
        """
        SELECT ot.*, oc.codigo_producto
        FROM ordenes_trabajo ot
        LEFT JOIN objetos_costo oc ON oc.objeto_costo_id = ot.objeto_costo_id
        """
    ).fetchall()
    for row in ots:
        ot_id = row["ot_id"]
        cmd = safe_float(conn.execute("SELECT SUM(cmd_soles) FROM cmd_ot WHERE run_id = ? AND ot_id = ?", (run_id, ot_id)).fetchone()[0])
        csd = safe_float(conn.execute("SELECT SUM(csd_soles) FROM csd_ot WHERE run_id = ? AND ot_id = ?", (run_id, ot_id)).fetchone()[0])
        prod = conn.execute("SELECT * FROM produccion_ot WHERE run_id = ? AND ot_id = ?", (run_id, ot_id)).fetchone()
        support = conn.execute("SELECT * FROM soporte_ot WHERE run_id = ? AND ot_id = ?", (run_id, ot_id)).fetchone()
        mod = safe_float(prod["costo_mod_ot"]) if prod else 0.0
        otros = safe_float(prod["costo_otros_recursos_ot"]) if prod else 0.0
        costo_prod = safe_float(prod["costo_produccion_ot"]) if prod else 0.0
        costo_support = safe_float(support["costo_soporte_ot"]) if support else 0.0
        total = cmd + costo_prod + costo_support + csd
        qty = safe_float(row["cantidad_producida"])
        unit = total / qty if qty else 0.0
        obs_csd = "" if csd else " Sin datos de servicio directo publicado en la fuente."
        conn.execute(
            """
            INSERT INTO hoja_costos_ot
                (run_id, empresa_id, periodo_id, modelo_id, ot_id, codigo_ot,
                 objeto_costo_id, codigo_producto, cantidad_producida, cmd_ot,
                 mod_ot, otros_recursos_ot, costo_produccion_ot, costo_soporte_ot,
                 csd_ot, costo_total_ot, costo_unitario_ot, formula, fuente, observacion)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                run_id,
                row["empresa_id"],
                row["periodo_id"],
                row["modelo_id"],
                ot_id,
                row["codigo_ot"],
                row["objeto_costo_id"],
                row["codigo_producto"],
                qty,
                cmd,
                mod,
                otros,
                costo_prod,
                costo_support,
                csd,
                total,
                unit,
                "costo_total_OT = CMD_OT + costo_produccion_OT + costo_soporte_OT + CSD_OT; costo_unitario_OT = costo_total_OT / cantidad_producida_OT",
                "Datos importados y calculos generados desde SQLite",
                "Hoja de costos generada por el sistema." + obs_csd,
            ),
        )
        ventas = safe_float(conn.execute("SELECT SUM(ventas_netas_soles) FROM ventas WHERE ot_id = ?", (ot_id,)).fetchone()[0])
        margen_bruto = ventas - cmd - costo_prod
        margen_operativo = margen_bruto - costo_support
        rent = margen_operativo / ventas if ventas else 0.0
        conn.execute(
            """
            INSERT INTO estado_resultados_ot
                (run_id, ot_id, objeto_costo_id, codigo_producto, ventas_netas_soles,
                 materia_prima_directa, costos_conversion, margen_bruto,
                 procesos_soporte, margen_operativo, rentabilidad_operativa,
                 formula, fuente, observacion)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                run_id,
                ot_id,
                row["objeto_costo_id"],
                row["codigo_producto"],
                ventas,
                cmd,
                costo_prod,
                margen_bruto,
                costo_support,
                margen_operativo,
                rent,
                "Ventas netas - materia prima directa - costos de conversion = margen bruto; margen bruto - procesos de soporte = margen operativo",
                "Ventas + hoja de costos OT",
                "Estado de resultados generado por el sistema.",
            ),
        )
    return len(ots)


def generate_dashboard_alerts(conn: sqlite3.Connection, run_id: int) -> int:
    conn.execute("DELETE FROM dashboard_alertas")
    top_activity = conn.execute(
        """
        SELECT nombre_actividad, costo_total_ac
        FROM actividad_total
        WHERE run_id = ?
        ORDER BY costo_total_ac DESC
        LIMIT 1
        """,
        (run_id,),
    ).fetchone()
    if top_activity:
        conn.execute(
            """
            INSERT INTO dashboard_alertas (run_id, tipo, severidad, mensaje, valor_referencia, observacion)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                run_id,
                "Actividad relevante",
                "Media",
                f"La actividad con mayor costo total es {top_activity['nombre_actividad']}.",
                top_activity["costo_total_ac"],
                "Revisar consumo de MOD, otros recursos y soporte.",
            ),
        )
    low_margin = conn.execute(
        """
        SELECT e.codigo_producto, e.rentabilidad_operativa, COALESCE(ot.codigo_ot, e.ot_id) AS codigo_ot
        FROM estado_resultados_ot e
        LEFT JOIN ordenes_trabajo ot ON ot.ot_id = e.ot_id
        WHERE e.run_id = ?
        ORDER BY e.rentabilidad_operativa ASC
        LIMIT 1
        """,
        (run_id,),
    ).fetchone()
    if low_margin:
        conn.execute(
            """
            INSERT INTO dashboard_alertas (run_id, tipo, severidad, mensaje, valor_referencia, observacion)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                run_id,
                "Rentabilidad",
                "Alta",
                f"El objeto de costo con menor rentabilidad operativa es la {low_margin['codigo_ot']} ({low_margin['codigo_producto']}).",
                low_margin["rentabilidad_operativa"],
                "Evaluar precio, mezcla de costos y actividades de mayor peso.",
            ),
        )
    return conn.execute("SELECT COUNT(*) FROM dashboard_alertas WHERE run_id = ?", (run_id,)).fetchone()[0]


COLUMNAS_AJUSTABLES_AC = ("costo_mod_ac", "costo_otros_recursos_ac", "costo_soporte_ac")
NOTA_SIMULACION = " Valor ajustado por simulacion del usuario."


def apply_direct_cost_overrides(conn: sqlite3.Connection, run_id: int, ajustes: list[dict]) -> None:
    """Simulacion: fija el CMD o el CSD total de una OT (se reparte en proporcion a sus filas)."""
    for aj in ajustes:
        tabla = aj.get("tabla")
        if tabla not in ("cmd_ot", "csd_ot"):
            continue
        campo = "cmd_soles" if tabla == "cmd_ot" else "csd_soles"
        filas = conn.execute(f"SELECT rowid, {campo} FROM {tabla} WHERE run_id = ? AND ot_id = ?", (run_id, aj["clave"])).fetchall()
        nuevo = safe_float(aj.get("valor"))
        actual = sum(safe_float(f[1]) for f in filas)
        if not filas:
            continue
        for i, fila in enumerate(filas):
            valor = nuevo * safe_float(fila[1]) / actual if actual else (nuevo if i == 0 else 0.0)
            conn.execute(
                f"UPDATE {tabla} SET {campo} = ?, observacion = COALESCE(observacion, '') || ? WHERE rowid = ?",
                (valor, NOTA_SIMULACION, fila[0]),
            )


def apply_activity_overrides(conn: sqlite3.Connection, run_id: int, ajustes: list[dict]) -> None:
    """Simulacion: fija PD (MOD), otros recursos o soporte de una actividad y recalcula sus totales."""
    for aj in ajustes:
        if aj.get("tabla") != "actividad_total" or aj.get("columna") not in COLUMNAS_AJUSTABLES_AC:
            continue
        actividad_id, columna, valor = aj["clave"], aj["columna"], safe_float(aj.get("valor"))
        conn.execute(
            f"UPDATE actividad_total SET {columna} = ?, observacion = COALESCE(observacion, '') || ? "
            "WHERE run_id = ? AND actividad_id = ?",
            (valor, NOTA_SIMULACION, run_id, actividad_id),
        )
        if columna == "costo_soporte_ac":
            fila = conn.execute(
                "SELECT costo_mantenimiento_ac, costo_control_calidad_ac, costo_ingenieria_ac, costo_soporte_ac "
                "FROM soporte_actividad WHERE run_id = ? AND actividad_id = ?",
                (run_id, actividad_id),
            ).fetchone()
            if fila:
                anterior = safe_float(fila[3])
                factor = valor / anterior if anterior else 0.0
                mant = safe_float(fila[0]) * factor if anterior else valor
                conn.execute(
                    """UPDATE soporte_actividad SET costo_mantenimiento_ac = ?, costo_control_calidad_ac = ?,
                       costo_ingenieria_ac = ?, costo_soporte_ac = ?, observacion = COALESCE(observacion, '') || ?
                       WHERE run_id = ? AND actividad_id = ?""",
                    (mant, safe_float(fila[1]) * factor, safe_float(fila[2]) * factor, valor, NOTA_SIMULACION,
                     run_id, actividad_id),
                )
        conn.execute(
            """UPDATE actividad_total SET costo_conversion_ac = costo_mod_ac + costo_otros_recursos_ac,
                   costo_total_ac = costo_mod_ac + costo_otros_recursos_ac + costo_soporte_ac
               WHERE run_id = ? AND actividad_id = ?""",
            (run_id, actividad_id),
        )


def run_model_1_calculation(db_path: str | Path, ajustes: list[dict] | None = None) -> int:
    """Calculo ABC del Modelo 1. ``ajustes`` (simulacion, opcional) fija valores intermedios:
    {"tabla": "actividad_total"|"cmd_ot"|"csd_ot", "clave": actividad_id u ot_id, "columna": ..., "valor": ...}.
    Sin ajustes el calculo es el de siempre."""
    initialize_database(db_path)
    models = evaluate_models(db_path)
    model_1 = next(item for item in models if item.code == "MODELO_1")
    if not model_1.available:
        raise RuntimeError(model_1.reason)

    with db_session(db_path) as conn:
        clear_calculated_tables(conn)
        run_id = start_calculation_run(
            conn, "MODELO_1", "Modelo ABC por Ordenes de Trabajo" + (" (simulacion)" if ajustes else "")
        )
        try:
            build_kardex_from_compras(conn, run_id)
            calculate_direct_costs(conn, run_id)
            if ajustes:
                apply_direct_cost_overrides(conn, run_id, ajustes)
            calculate_mod_to_activities(conn, run_id)
            calculate_otros_recursos_to_activities(conn, run_id)
            calculate_support_to_activities(conn, run_id)
            calculate_activity_totals(conn, run_id)
            if ajustes:
                apply_activity_overrides(conn, run_id, ajustes)
            calculate_ac_to_ot(conn, run_id)
            calculate_cost_sheets(conn, run_id)
            generate_dashboard_alerts(conn, run_id)
            finish_calculation_run(conn, run_id, "COMPLETADO", "Calculo ABC ejecutado correctamente.")
            return run_id
        except Exception as exc:
            finish_calculation_run(conn, run_id, "ERROR", str(exc))
            raise
