from __future__ import annotations

import sqlite3

from .utils import round_half_up, round_percentage_fraction, safe_float, safe_int


MOD_FORMULA = (
    "Ts_total = SUM(tolerancias); Te = ROUND(tiempo_normal_min_kg * (1 + Ts_total), "
    "redondear_te_decimales); Fd = horas_reales / horas_estandar; "
    "tiempo_real_AC = Te * kg_producidos_periodo * Fd; porcentaje_AC = tiempo_real_AC / "
    "tiempo_real_total_minutos_periodo; total_PD = cantidad_operarios * salario_mensual; "
    "costo_PD_AC = porcentaje_AC * total_PD"
)


def calculate_mod_to_activities(conn: sqlite3.Connection, run_id: int) -> int:
    conn.execute("DELETE FROM mod_actividad")
    rows = conn.execute('SELECT * FROM "Personal_directos"').fetchall()
    for row in rows:
        tolerances = [
            safe_float(row["tolerancia_personal_pct"]),
            safe_float(row["tolerancia_fatiga_pct"]),
            safe_float(row["tolerancia_pie_pct"]),
            safe_float(row["tolerancia_encorvado_pct"]),
        ]
        ts_total = sum(tolerances)
        te_decimals = safe_int(row["redondear_te_decimales"], 2)
        pct_decimals = safe_int(row["redondear_porcentaje_decimales"], 0)
        tn = safe_float(row["tiempo_normal_min_kg"])
        te = round_half_up(tn * (1 + ts_total), te_decimals)
        horas_reales = safe_float(row["horas_reales_trabajadas_periodo"])
        horas_estandar = safe_float(row["horas_estandar_totales_periodo"])
        fd_raw = horas_reales / horas_estandar if horas_estandar else 0
        use_pdf_rounding = str(row["usar_porcentaje_redondeado_pdf"] or "").strip().upper() == "SI"
        fd = round_half_up(fd_raw, 2) if use_pdf_rounding else fd_raw
        tiempo_real = te * safe_float(row["kg_producidos_periodo"]) * fd
        tiempo_total = safe_float(row["tiempo_real_total_minutos_periodo"])
        porcentaje_calculado = tiempo_real / tiempo_total if tiempo_total else 0
        porcentaje_usado = (
            round_percentage_fraction(porcentaje_calculado, pct_decimals)
            if use_pdf_rounding
            else porcentaje_calculado
        )
        total_mod = safe_float(row["cantidad_operarios"]) * safe_float(row["salario_mensual_operario_soles"])
        costo_mod = porcentaje_usado * total_mod
        if use_pdf_rounding:
            tiempo_real = round_half_up(tiempo_real, 0)
            costo_mod = round_half_up(costo_mod, 0)
        conn.execute(
            """
            INSERT INTO mod_actividad
                (run_id, actividad_id, codigo_actividad, nombre_actividad, ts_total,
                 te, fd, tiempo_real_minutos, porcentaje_calculado, porcentaje_usado,
                 total_mod_soles, costo_mod_ac, formula, fuente, observacion)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                run_id,
                row["actividad_id"],
                row["codigo_actividad"],
                row["nombre_actividad"],
                ts_total,
                te,
                fd,
                tiempo_real,
                porcentaje_calculado,
                porcentaje_usado,
                total_mod,
                costo_mod,
                MOD_FORMULA,
                row["fuente"],
                row["observacion"],
            ),
        )
    return len(rows)
