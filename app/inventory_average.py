from __future__ import annotations

import sqlite3


def calculate_average_unit_cost(total_cost: float, total_quantity: float) -> float:
    if total_quantity == 0:
        return 0.0
    return total_cost / total_quantity


def build_kardex_from_compras(conn: sqlite3.Connection, run_id: int) -> int:
    """Resumen del kardex por promedio ponderado (hoja Compras) por material y OT de destino."""
    from .kardex_compras import kardex_desde_base, mapa_ot, salidas_por_ot

    conn.execute("DELETE FROM kardex_promedio")
    items = kardex_desde_base(conn)
    acumulado, _avisos = salidas_por_ot(items, mapa_ot(conn))
    por_material = {item.item_codigo: item for item in items}
    count = 0
    for (ot_id, material_id), dato in sorted(acumulado.items()):
        item = por_material[material_id]
        entrada_q = item.ingresos_cantidad
        entrada_t = item.ingresos_total
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
                material_id,
                ot_id,
                dato["material"],
                entrada_q,
                calculate_average_unit_cost(entrada_t, entrada_q),
                calculate_average_unit_cost(dato["total"], dato["cantidad"]),
                dato["cantidad"],
                dato["total"],
                "Costo promedio = valor del saldo / cantidad del saldo; salida = cantidad x costo promedio vigente",
                f"Saldo inicial {item.saldo_inicial_cantidad:,.2f} {item.unidad} S/ {item.saldo_inicial_total:,.2f}; "
                f"saldo final {item.saldo_final_cantidad:,.2f} {item.unidad} S/ {item.saldo_final_total:,.2f}.",
            ),
        )
        count += 1
    return count


# compatibilidad con el nombre anterior
build_kardex_from_materiales = build_kardex_from_compras
