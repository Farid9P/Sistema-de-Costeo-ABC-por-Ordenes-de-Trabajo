"""Pruebas del Agente IA de Finanzas y Contabilidad.

    python -m unittest tests.test_agente_ia -v
"""

from __future__ import annotations

import hashlib
import shutil
import sqlite3
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

from app.agente_ia import AgenteFinanzas  # noqa: E402
from app.agente_ia.kardex import KardexPromedioPonderado, StockInsuficiente  # noqa: E402
from app.agente_ia.lectura import LectorSoloLectura, OperacionNoPermitida  # noqa: E402
from parafrasis_agente import PARAFRASIS  # noqa: E402

DB_ORIGEN = ROOT / "costos_abc_frenosa.sqlite"


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


@unittest.skipUnless((ROOT / "data" / "frenosa_datos_entrada_abc.xlsx").exists(), "Se necesita el Excel maestro")
class TestAgente(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = Path(tempfile.mkdtemp())
        cls.db = cls.tmp / "copia.sqlite"
        # La app empieza vacia: la prueba importa y calcula el Excel maestro (julio 2010) en una base propia.
        from app.costing_abc import run_model_1_calculation
        from app.importers import import_excel_to_sqlite

        import_excel_to_sqlite(ROOT / "data" / "frenosa_datos_entrada_abc.xlsx", cls.db)
        run_model_1_calculation(cls.db)
        cls.agente = AgenteFinanzas(cls.db)

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def test_ejemplos_del_usuario(self):
        r = self.agente.consultar("cual es producto con mas ventas?")
        self.assertEqual(r.interpretacion.interpretada, "¿Cuál es el objeto de costo que genera más ingresos?")
        self.assertIn("FEF0014P_FRE9000", r.respuesta.resumen)
        self.assertIn("S/ 453,602.00", r.respuesta.resumen)
        r = self.agente.consultar("Que producto genero mas costos?")
        self.assertEqual(r.interpretacion.interpretada, "¿Cuál objeto de costo generó más costo de ventas?")
        self.assertIn("S/ 341,687.69", r.respuesta.resumen)

    def test_parafrasis_del_catalogo(self):
        fallas = []
        for pregunta, esperada in PARAFRASIS:
            i = self.agente.interpretar(pregunta)
            obtenida = i.pregunta.codigo if i.pregunta else None
            if obtenida != esperada:
                fallas.append(f"{pregunta!r}: esperada {esperada}, obtenida {obtenida}")
        self.assertEqual(fallas, [])

    def test_cada_pregunta_del_txt_se_reconoce_y_responde(self):
        for base in self.agente.catalogo:
            r = self.agente.consultar(base.texto.replace("{", "").replace("}", ""))
            self.assertEqual(r.interpretacion.pregunta.codigo, base.codigo, base.texto)
            self.assertIsNotNone(r.respuesta, base.texto)
            self.assertNotEqual(r.respuesta.titulo, "Error", f"{base.codigo}: {r.respuesta.resumen}")

    def test_no_modifica_la_base(self):
        antes = sha256(self.db)
        for base in self.agente.catalogo:
            self.agente.consultar(base.texto)
        self.agente.propuestas()
        self.agente.consultar("borra las ventas de sw50")
        self.assertEqual(antes, sha256(self.db))

    def test_conexion_es_solo_lectura(self):
        with LectorSoloLectura(self.db) as lector:
            with self.assertRaises(OperacionNoPermitida):
                lector.filas("DELETE FROM ventas")
            with self.assertRaises(OperacionNoPermitida):
                lector.filas("SELECT 1; DROP TABLE ventas")
            with self.assertRaises(sqlite3.OperationalError):
                lector.conectar().execute("INSERT INTO ventas (venta_id) VALUES ('X')")

    def test_pedidos_de_cambio(self):
        for texto in ["modifica las ventas de sw50", "borra el kardex", "puedes cambiar el precio de fre9000",
                      "genera los reportes excel", "sube el costo de mantenimiento 10%"]:
            self.assertTrue(self.agente.interpretar(texto).es_pedido_de_cambio, texto)
        for texto in ["como reducir costos", "que alertas genera el dashboard", "que reportes puede generar el aplicativo",
                      "cual es el producto mas rentable"]:
            self.assertFalse(self.agente.interpretar(texto).es_pedido_de_cambio, texto)
        r = self.agente.consultar("modifica las ventas de sw50")
        self.assertIn("solo lectura", r.texto())

    def test_kardex_concilia_con_hoja_de_costos(self):
        r = self.agente.consultar("muestrame el kardex")
        self.assertEqual(r.interpretacion.pregunta.codigo, "E2")
        self.assertIn("S/ 443,993.26", r.respuesta.resumen)
        r = self.agente.consultar("costo unitario promedio del kardex")
        self.assertIn("diferencia S/ 0.00", r.texto())

    def test_propuestas_cuantificadas(self):
        props = self.agente.propuestas()
        self.assertGreaterEqual(len(props), 5)
        self.assertIn("FEF0011P_FRE9002", props[0].titulo)
        self.assertAlmostEqual(props[0].impacto_soles, 4014.35, places=2)


class TestKardex(unittest.TestCase):
    def test_promedio_ponderado_movil(self):
        k = KardexPromedioPonderado("Material X")
        k.saldo_inicial(100, 10)
        k.entrada(200, 13)
        self.assertAlmostEqual(k.costo_promedio, 12.0)
        s = k.salida(150, referencia="OT-1")
        self.assertAlmostEqual(s.sal_total, 1800.0)
        k.entrada(50, 13, flete=50)  # C.U. de entrada = (50 x 13 + 50) / 50 = 14
        self.assertAlmostEqual(k.filas[-1].ent_costo_unitario, 14.0)
        self.assertAlmostEqual(k.costo_promedio, 12.5)
        ec = k.ecuacion_almacen()
        self.assertTrue(ec["cuadra"])
        self.assertAlmostEqual(ec["consumo"], 1800.0)

    def test_no_permite_stock_negativo(self):
        k = KardexPromedioPonderado("Material Y")
        k.entrada(10, 5)
        with self.assertRaises(StockInsuficiente):
            k.salida(11)


if __name__ == "__main__":
    unittest.main()
