"""Analisis por periodos: un mes reproduce el calculo de un periodo; varios meses suman.

    python -m unittest tests.test_analisis_periodos -v
"""

from __future__ import annotations

import shutil
import sqlite3
import sys
import tempfile
import unittest
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from app import analisis_periodos as ap  # noqa: E402
from app.costing_abc import run_model_1_calculation  # noqa: E402
from app.database import initialize_database  # noqa: E402
from app.importers import import_excel_to_sqlite  # noqa: E402

MAESTRO = ROOT / "data" / "frenosa_datos_entrada_abc.xlsx"
CARPETA = ROOT / "data" / "simulacion_2010"
MESES_SIM = sorted(CARPETA.glob("frenosa_datos_entrada_abc_2010_*.xlsx"))


@unittest.skipUnless(len(MESES_SIM) == 12 and MAESTRO.exists(), "Faltan los Excel de 2010")
class TestAnalisisPeriodos(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = Path(tempfile.mkdtemp(prefix="test_periodos_"))
        cls.db = cls.tmp / "mp.sqlite"
        initialize_database(cls.db)
        cls.hechos, cls.errores = ap.procesar_varios(MESES_SIM, cls.db)
        # periodo de un solo mes calculado con el flujo normal (julio, Excel maestro)
        cls.base = cls.tmp / "julio.sqlite"
        import_excel_to_sqlite(MAESTRO, cls.base)
        cls.run_id = run_model_1_calculation(cls.base)

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def test_doce_meses_procesados(self):
        self.assertEqual(self.errores, [])
        self.assertEqual([(h.anio, h.mes) for h in self.hechos], [(2010, m) for m in range(1, 13)])
        self.assertEqual(len(ap.periodos_procesados(self.db)), 12)

    def test_un_mes_reproduce_el_calculo_del_periodo(self):
        db = self.tmp / "solo_julio.sqlite"
        initialize_database(db)
        ap.guardar_resultados(self.base, db, archivo=MAESTRO.name)
        p = ap.analizar(db, [(2010, 7)], "Mensual")
        conn = sqlite3.connect(self.base)
        conn.row_factory = sqlite3.Row
        comparados = 0
        for tabla, clave in (("hoja_costos_ot", "hoja"), ("estado_resultados_ot", "estado")):
            for fila in conn.execute(f"SELECT * FROM {tabla} WHERE run_id = ?", (self.run_id,)):
                fila = dict(fila)
                agregado = next(x for x in p[clave] if x["codigo_producto"] == fila["codigo_producto"])
                for campo, valor in fila.items():
                    if isinstance(valor, float) and campo in agregado:
                        self.assertEqual(agregado[campo], valor, (tabla, campo))
                        comparados += 1
        for fila in conn.execute("SELECT * FROM actividad_total WHERE run_id = ?", (self.run_id,)):
            agregado = next(x for x in p["costos_actividad"] if x["actividad_id"] == fila["actividad_id"])
            for campo in ap.CAMPOS_ACTIVIDAD:
                self.assertEqual(agregado[campo], fila[campo])
                comparados += 1
        conn.close()
        self.assertGreater(comparados, 60)
        self.assertAlmostEqual(p["kpis"]["ventas"], 888730.00, places=2)
        self.assertAlmostEqual(p["kpis"]["margen_operativo"], 248453.9236, places=4)

    def test_trimestre_suma_sus_meses(self):
        etiqueta, meses = ap.bloques("Trimestral", 2010)[0]
        self.assertEqual(etiqueta, "Trimestre 1: Enero–marzo 2010")
        t = ap.analizar(self.db, meses, "Trimestral", etiqueta)
        self.assertTrue(t["completo"])
        mensuales = [ap.analizar(self.db, [m], "Mensual") for m in meses]
        for campo in ("ventas", "costo_total", "margen_operativo", "unidades"):
            self.assertAlmostEqual(t["kpis"][campo], sum(m["kpis"][campo] for m in mensuales), places=6)
        for fila in t["hoja"]:
            suma = sum(next(x for x in m["hoja"] if x["codigo_producto"] == fila["codigo_producto"])["costo_total_ot"]
                       for m in mensuales)
            self.assertAlmostEqual(fila["costo_total_ot"], suma, places=6)
            self.assertAlmostEqual(fila["costo_unitario_ot"], fila["costo_total_ot"] / fila["cantidad_producida"], places=12)
            self.assertEqual(fila["codigo_ot"], fila["codigo_ot"].split("-2010-")[0])
        for fila in t["estado"]:
            self.assertAlmostEqual(fila["rentabilidad_operativa"], fila["margen_operativo"] / fila["ventas_netas_soles"], places=12)

    def test_anual_y_bloques(self):
        anual = ap.analizar(self.db, ap.bloques("Anual", 2010)[0][1], "Anual", "Año 2010")
        self.assertAlmostEqual(anual["kpis"]["rentabilidad_global"], 0.2264, places=4)
        for tipo, n in (("Mensual", 12), ("Bimestral", 6), ("Trimestral", 4), ("Cuatrimestral", 3), ("Semestral", 2)):
            comp = ap.comparar_bloques(self.db, tipo, 2010)
            self.assertEqual(len(comp), n)
            self.assertAlmostEqual(sum(b["ventas_netas_soles"] for b in comp), anual["kpis"]["ventas"], places=4)
            self.assertTrue(all(b["estado"] == "Completo" for b in comp))

    def test_kardex_del_rango(self):
        anual = ap.analizar(self.db, ap.meses_entre((2010, 1), (2010, 12)), "Anual")
        self.assertEqual([k["material_id"] for k in anual["kardex"]], ["MAT001", "MAT002", "MAT003"])
        for k in anual["kardex"]:
            self.assertEqual(k["continuidad"], "OK")
            self.assertEqual(k["salidas_vs_cmd"], "coincide")
            self.assertLessEqual(abs(k["ajuste_redondeo"]), 0.02)

    def test_meses_faltantes_y_personalizado(self):
        meses = ap.meses_entre((2010, 11), (2011, 2))
        p = ap.analizar(self.db, meses)
        self.assertEqual(p["titulo"], "Noviembre 2010 – Febrero 2011")
        self.assertEqual(p["meses_con_datos"], [(2010, 11), (2010, 12)])
        self.assertEqual(p["meses_faltantes"], [(2011, 1), (2011, 2)])
        self.assertFalse(p["completo"])
        self.assertTrue(any(a["tipo"] == "Datos incompletos" for a in p["alertas"]))

    def test_rechaza_periodo_no_mensual(self):
        copia = self.tmp / "bimestral.xlsx"
        wb = openpyxl.load_workbook(MESES_SIM[0])
        ws = wb["Periodo"]
        encabezado = [c.value for c in ws[1]]
        ws.cell(2, encabezado.index("fecha_fin") + 1, "2010-02-28")
        wb.save(copia)
        with self.assertRaises(ValueError):
            ap.procesar_excel(copia, self.tmp / "otra.sqlite")

    def test_reporte_excel(self):
        from app.reporte_periodos import generar_reporte_periodos

        etiqueta, meses = ap.bloques("Semestral", 2010)[1]
        ruta = generar_reporte_periodos(self.db, meses, "Semestral", etiqueta, self.tmp)
        wb = openpyxl.load_workbook(ruta)
        for hoja in ("Resumen", "Hoja_costos", "Estado_resultados", "Proceso1_EC_AC", "Proceso2_Produccion",
                     "Proceso2_Soporte", "Evolucion_mensual", "Comparativo", "Kardex_materiales", "Detalle_mensual_OT"):
            self.assertIn(hoja, wb.sheetnames)
        self.assertTrue(str(wb["Hoja_costos"]["E5"].value).startswith("=SUMIFS("))


if __name__ == "__main__":
    unittest.main()


@unittest.skipUnless(len(MESES_SIM) == 12, "Faltan los Excel de 2010")
class TestCargaYSimulacion(unittest.TestCase):
    """Flujo de la aplicacion: carga de varios Excel, opciones por meses consecutivos y simulacion."""

    @classmethod
    def setUpClass(cls):
        cls.tmp = Path(tempfile.mkdtemp(prefix="test_sim_"))
        cls.db = cls.tmp / "app.sqlite"
        ap.iniciar_sesion_vacia(cls.db)
        cls.hechos, cls.errores = ap.procesar_varios([MESES_SIM[0], MESES_SIM[1], MESES_SIM[2], MESES_SIM[4]], cls.db)

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def setUp(self):
        ap.borrar_ajustes(self.db)

    def test_identifica_el_periodo_de_cada_excel(self):
        self.assertEqual(self.errores, [])
        periodo = ap.leer_periodo_excel(MESES_SIM[4])
        self.assertEqual((periodo["anio"], periodo["mes"], periodo["periodo_id"]), (2010, 5, "PER2010-05"))

    def test_opciones_segun_meses_consecutivos(self):
        v = ap.ventanas([(p["anio"], p["mes"]) for p in ap.periodos_procesados(self.db)])
        self.assertEqual(list(v), ["Mensual", "Bimestral", "Trimestral"])  # ene-mar consecutivos; mayo suelto
        self.assertEqual(len(v["Mensual"]), 4)
        self.assertEqual([m for _e, m in v["Bimestral"]], [[(2010, 1), (2010, 2)], [(2010, 2), (2010, 3)]])
        self.assertEqual(v["Trimestral"][0][0], "Enero–marzo 2010")
        self.assertNotIn("Semestral", v)

    def test_sesion_vacia(self):
        db = self.tmp / "vacia.sqlite"
        ap.procesar_excel(MESES_SIM[0], db)
        ap.iniciar_sesion_vacia(db)
        self.assertEqual(ap.periodos_procesados(db), [])
        self.assertEqual(sqlite3.connect(db).execute("SELECT COUNT(*) FROM ordenes_trabajo").fetchone()[0], 0)

    def test_mes_en_base_reproduce_lo_importado(self):
        ap.calcular_mes_en_base(self.db, 2010, 1)
        with sqlite3.connect(self.db) as conn:
            conn.row_factory = sqlite3.Row
            run = conn.execute("SELECT MAX(run_id) FROM calculation_runs").fetchone()[0]
            total = conn.execute("SELECT SUM(costo_total_ot) FROM hoja_costos_ot WHERE run_id = ?", (run,)).fetchone()[0]
        self.assertAlmostEqual(total, self.hechos[0].costo_total, places=9)

    def test_consolidado_en_base_igual_al_analisis(self):
        meses = [(2010, 1), (2010, 2), (2010, 3)]
        resultados = ap.resultados_meses(self.db, meses)
        p = ap.analizar(self.db, meses, "Trimestral", resultados=resultados)
        run = ap.consolidar_en_base(self.db, p, resultados)
        with sqlite3.connect(self.db) as conn:
            conn.row_factory = sqlite3.Row
            for fila in conn.execute("SELECT * FROM estado_resultados_ot WHERE run_id = ?", (run,)):
                esperado = next(x for x in p["estado"] if x["codigo_producto"] == fila["codigo_producto"])
                self.assertAlmostEqual(fila["margen_operativo"], esperado["margen_operativo"], places=6)
                self.assertAlmostEqual(fila["rentabilidad_operativa"], esperado["rentabilidad_operativa"], places=12)

    def test_simulacion_dato_resultado_y_reestablecer(self):
        meses = [(2010, 1), (2010, 2)]
        inicial = ap.analizar(self.db, meses, "Bimestral")
        datos = ap.datos_mes(self.db, 2010, 1, con_ajustes=False)
        i = next(i for i, f in enumerate(datos["ventas"]) if f["ot_id"] == "OT001")
        original = datos["ventas"][i]["ventas_netas_soles"]
        ap.registrar_ajuste(self.db, 2010, 1, "dato", "ventas", "ventas_netas_soles", original, original + 10000, fila=i)
        sim = ap.analizar(self.db, meses, "Bimestral", resultados=ap.resultados_meses(self.db, meses))
        self.assertAlmostEqual(sim["kpis"]["ventas"] - inicial["kpis"]["ventas"], 10000, places=6)
        self.assertAlmostEqual(sim["kpis"]["margen_operativo"] - inicial["kpis"]["margen_operativo"], 10000, places=6)
        # valor calculado fijado: soporte de una actividad a cero en febrero
        res = ap.resultados_meses(self.db, [(2010, 2)], simulados=False)[(2010, 2)]
        act = res["actividad_total"][0]
        ap.registrar_ajuste(self.db, 2010, 2, "resultado", "actividad_total", "costo_soporte_ac", act["costo_soporte_ac"], 0.0,
                            clave=act["actividad_id"])
        sim2 = ap.analizar(self.db, meses, "Bimestral", resultados=ap.resultados_meses(self.db, meses))
        self.assertAlmostEqual(sim2["kpis"]["margen_operativo"] - sim["kpis"]["margen_operativo"], act["costo_soporte_ac"], places=6)
        # volver a escribir el valor inicial elimina el ajuste; reestablecer borra todo
        ap.registrar_ajuste(self.db, 2010, 1, "dato", "ventas", "ventas_netas_soles", original, original, fila=i)
        self.assertEqual(len(ap.ajustes(self.db)), 1)
        ap.borrar_ajustes(self.db)
        final = ap.analizar(self.db, meses, "Bimestral", resultados=ap.resultados_meses(self.db, meses))
        self.assertEqual(final["kpis"], inicial["kpis"])
