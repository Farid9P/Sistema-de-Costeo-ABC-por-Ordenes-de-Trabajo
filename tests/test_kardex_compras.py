"""Kardex por promedio ponderado (hoja Compras).

La primera prueba reproduce el ejemplo de referencia "Valorización de inventario
de repuestos" (repuesto RP154K, setiembre 2020): saldo final 68 unid S/ 12,845.17.

    python -m unittest tests.test_kardex_compras -v
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from app.kardex_compras import valorizar  # noqa: E402

EJEMPLO = [
    ("2020-09-01", "Saldo inicial", 45, None, "", {"importe_soles": 8394}),
    ("2020-09-02", "Compra", 10, 217.8, "F/ 01-1250", {}),
    ("2020-09-06", "Consumo", 8, None, "", {}),
    ("2020-09-07", "Consumo", 6, None, "", {}),
    ("2020-09-09", "Compra", 18, 208.8, "F/ 01-1254", {}),
    ("2020-09-10", "Devolucion a proveedor", 2, 208.8, "F/ 01-1254", {}),
    ("2020-09-15", "Consumo", 3, None, "", {}),
    ("2020-09-18", "Consumo", 3, None, "", {}),
    ("2020-09-20", "Consumo", 14, None, "", {}),
    ("2020-09-20", "Devolucion de consumo", 10, None, "", {}),
    ("2020-09-22", "Consumo", 1, None, "", {}),
    ("2020-09-25", "Compra", 28, 233.2, "F/ 01-1256", {}),
    ("2020-09-28", "Consumo", 6, None, "", {}),
]


def movimientos():
    salida = []
    for n, (fecha, tipo, cantidad, precio, doc, extra) in enumerate(EJEMPLO, start=1):
        salida.append({"compra_id": f"M{n:02d}", "item_codigo": "RP154K", "fecha": fecha, "tipo_movimiento": tipo,
                       "cantidad": cantidad, "precio_unitario_con_igv": precio, "igv_pct": 0.18, "documento": doc, **extra})
    return salida


class TestKardexCompras(unittest.TestCase):
    def test_reproduce_ejemplo_de_referencia(self):
        k = valorizar(movimientos())[0]
        saldos = [(round(f.saldo_cantidad, 2), round(f.saldo_costo_unitario, 2), round(f.saldo_total, 2)) for f in k.filas]
        self.assertEqual(saldos, [
            (45, 186.53, 8394.00), (55, 186.18, 10239.76), (47, 186.18, 8750.34), (41, 186.18, 7633.28),
            (59, 183.36, 10818.36), (57, 183.59, 10464.46), (54, 183.59, 9913.70), (51, 183.59, 9362.94),
            (37, 183.59, 6792.72), (47, 183.59, 8628.59), (46, 183.59, 8445.01), (74, 188.90, 13978.57),
            (68, 188.90, 12845.17),
        ])
        self.assertAlmostEqual(k.ingresos_cantidad, 54)
        self.assertAlmostEqual(round(k.ingresos_total, 2), 10210.51)
        self.assertAlmostEqual(k.salidas_cantidad, 31)
        self.assertAlmostEqual(round(k.salidas_total, 2), 5759.34)
        comp = {c: (q, round(t, 2)) for c, q, t in k.comprobacion()}
        self.assertEqual(comp["Disponible"], (99, 18604.51))
        self.assertEqual(comp["Saldo final"], (68, 12845.17))
        self.assertTrue(k.cuadra)
        self.assertEqual(k.avisos, [])

    def test_devolucion_a_proveedor_al_costo_de_la_compra(self):
        k = valorizar(movimientos())[0]
        dev = k.filas[5]
        self.assertAlmostEqual(dev.ing_cantidad, -2)
        self.assertAlmostEqual(round(dev.ing_costo_unitario, 2), 176.95)
        self.assertAlmostEqual(round(dev.ing_total, 2), -353.90)

    def test_no_permite_stock_negativo(self):
        movs = [{"compra_id": "A", "item_codigo": "X", "fecha": "2010-07-01", "tipo_movimiento": "Saldo inicial",
                 "cantidad": 5, "importe_soles": 50},
                {"compra_id": "B", "item_codigo": "X", "fecha": "2010-07-02", "tipo_movimiento": "Consumo", "cantidad": 6}]
        k = valorizar(movs)[0]
        self.assertEqual(len(k.filas), 1)
        self.assertTrue(any("stock insuficiente" in a for a in k.avisos))


if __name__ == "__main__":
    unittest.main()
