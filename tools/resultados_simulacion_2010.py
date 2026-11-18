"""Importa y calcula con el sistema cada Excel simulado de 2010 y devuelve sus resultados.

    python tools/resultados_simulacion_2010.py            # imprime el resumen
    python tools/resultados_simulacion_2010.py --json r.json

Cada mes se procesa en una base SQLite temporal (no toca la base de la aplicación).
"""

from __future__ import annotations

import argparse
import json
import sqlite3
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from app.costing_abc import run_model_1_calculation  # noqa: E402
from app.importers import import_excel_to_sqlite, validate_workbook  # noqa: E402
from app.kardex_compras import conciliacion_cmd, kardex_desde_base  # noqa: E402

CARPETA = ROOT / "data" / "simulacion_2010"


def resultados_mes(excel: Path) -> dict:
    validacion = validate_workbook(excel)
    with tempfile.TemporaryDirectory() as tmp:
        db = Path(tmp) / "mes.sqlite"
        imp = import_excel_to_sqlite(excel, db)
        if not imp.ok:
            raise RuntimeError(f"{excel.name}: {[i.message for i in imp.errors]}")
        run_model_1_calculation(db)
        conn = sqlite3.connect(db)
        conn.row_factory = sqlite3.Row
        run = conn.execute("SELECT MAX(run_id) FROM calculation_runs WHERE status = 'COMPLETADO'").fetchone()[0]
        periodo = dict(conn.execute("SELECT * FROM periodo").fetchone())
        estado = [dict(r) for r in conn.execute(
            """SELECT e.*, o.codigo_ot, o.cantidad_producida FROM estado_resultados_ot e
               JOIN ordenes_trabajo o ON o.ot_id = e.ot_id WHERE e.run_id = ? ORDER BY e.ot_id""", (run,))]
        hoja = [dict(r) for r in conn.execute("SELECT * FROM hoja_costos_ot WHERE run_id = ? ORDER BY ot_id", (run,))]
        avisos = [dict(r) for r in conn.execute("SELECT severity, sheet_name, message FROM import_errors")]
        kardex = kardex_desde_base(conn)
        conc = {c["material_id"]: c for c in conciliacion_cmd(conn, kardex)}
        conn.close()
    ventas = sum(e["ventas_netas_soles"] for e in estado)
    margen = sum(e["margen_operativo"] for e in estado)
    return {
        "archivo": excel.name,
        "periodo_id": periodo["periodo_id"],
        "periodo": periodo["periodo"],
        "ventas": ventas,
        "costo_total": sum(h["costo_total_ot"] for h in hoja),
        "cmd": sum(h["cmd_ot"] for h in hoja),
        "produccion": sum(h["costo_produccion_ot"] for h in hoja),
        "soporte": sum(h["costo_soporte_ot"] for h in hoja),
        "margen_bruto": sum(e["margen_bruto"] for e in estado),
        "margen_operativo": margen,
        "rentabilidad": margen / ventas if ventas else 0.0,
        "unidades": sum(h["cantidad_producida"] for h in hoja),
        "por_ot": [{"codigo_ot": e["codigo_ot"], "codigo_producto": e["codigo_producto"], "unidades": e["cantidad_producida"],
                    "ventas": e["ventas_netas_soles"], "margen_operativo": e["margen_operativo"],
                    "rentabilidad": e["rentabilidad_operativa"]} for e in estado],
        "kardex": [{"item": k.item_codigo, "descripcion": k.item_descripcion, "si_q": k.saldo_inicial_cantidad,
                    "si_v": k.saldo_inicial_total, "compras": k.ingresos_total, "consumos": k.salidas_total,
                    "sf_q": k.saldo_final_cantidad, "sf_v": k.saldo_final_total, "cuadra": k.cuadra,
                    "codigo_ot": conc[k.item_codigo]["codigo_ot"] if k.item_codigo in conc else None,
                    "cmd": conc[k.item_codigo]["cmd_soles"] if k.item_codigo in conc else None,
                    "coincide_cmd": conc[k.item_codigo]["cuadra"] if k.item_codigo in conc else False} for k in kardex],
        "advertencias_validacion": [i.message for i in validacion.issues],
        "avisos_importacion": avisos,
    }


def todos() -> list[dict]:
    return [resultados_mes(p) for p in sorted(CARPETA.glob("frenosa_datos_entrada_abc_2010_*.xlsx"))]


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--json")
    args = parser.parse_args()
    res = todos()
    for r in res:
        print(f"{r['periodo']:<16} ventas {r['ventas']:>12,.2f}  costo {r['costo_total']:>12,.2f}  "
              f"margen op. {r['margen_operativo']:>12,.2f}  rentab. {r['rentabilidad'] * 100:7.2f} %")
    if args.json:
        Path(args.json).write_text(json.dumps(res, ensure_ascii=False, indent=1, default=str))
