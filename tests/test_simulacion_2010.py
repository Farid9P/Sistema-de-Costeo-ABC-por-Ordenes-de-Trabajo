"""Los 12 Excel simulados de 2010 importan sin advertencias, julio reproduce el caso, el kardex es continuo
y sus salidas igualan al costo de material directo de cada OT.

    python -m unittest tests.test_simulacion_2010 -v
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tools"))

from resultados_simulacion_2010 import CARPETA, todos  # noqa: E402


@unittest.skipUnless(len(list(CARPETA.glob("frenosa_datos_entrada_abc_2010_*.xlsx"))) == 12, "Faltan los Excel simulados")
class TestSimulacion2010(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.res = todos()

    def test_doce_meses_sin_advertencias(self):
        self.assertEqual(len(self.res), 12)
        for r in self.res:
            self.assertEqual(r["advertencias_validacion"], [], r["archivo"])
            self.assertEqual(r["avisos_importacion"], [], r["archivo"])

    def test_julio_es_el_caso_real(self):
        julio = next(r for r in self.res if r["periodo_id"] == "PER2010-07")
        self.assertAlmostEqual(julio["ventas"], 888730.00, places=2)
        self.assertAlmostEqual(julio["costo_total"], 640276.0764, places=3)
        self.assertAlmostEqual(julio["margen_operativo"], 248453.9236, places=3)

    def test_meses_negativos_y_mejores_que_julio(self):
        rent = {r["periodo_id"][-2:]: r["rentabilidad"] for r in self.res}
        self.assertTrue({"01", "02", "11"} <= {m for m, v in rent.items() if v < 0})
        self.assertGreaterEqual(sum(1 for v in rent.values() if v > rent["07"]), 3)

    def test_kardex_continuo(self):
        previo = {}
        for r in self.res:
            for k in r["kardex"]:
                self.assertTrue(k["cuadra"], (r["archivo"], k["item"]))
                if k["item"] in previo:
                    q, v = previo[k["item"]]
                    self.assertAlmostEqual(k["si_q"], q, places=6)
                    self.assertAlmostEqual(k["si_v"], v, delta=0.006)
                previo[k["item"]] = (k["sf_q"], k["sf_v"])

    def test_salidas_kardex_igualan_cmd(self):
        for r in self.res:
            self.assertEqual({k["item"] for k in r["kardex"]}, {"MAT001", "MAT002", "MAT003"}, r["archivo"])
            for k in r["kardex"]:
                self.assertTrue(k["coincide_cmd"], (r["archivo"], k["item"], k["consumos"], k["cmd"]))
                self.assertEqual(round(k["consumos"], 2), round(k["cmd"], 2))
            self.assertAlmostEqual(sum(k["consumos"] for k in r["kardex"]), r["cmd"], delta=0.015)


if __name__ == "__main__":
    unittest.main()
