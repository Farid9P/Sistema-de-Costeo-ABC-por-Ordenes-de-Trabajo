"""Crea data/simulacion_2010/Resumen_simulacion_2010.xlsx con los resultados que calcula el sistema
para cada Excel mensual (importación + cálculo ABC en una base temporal)."""

from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tools"))

from resultados_simulacion_2010 import CARPETA, todos  # noqa: E402
from simular_periodos_2010 import ESCENARIOS  # noqa: E402

F = "Arial"
TIT = Font(name=F, size=12, bold=True)
NEG = Font(name=F, size=10, bold=True)
NOR = Font(name=F, size=10)
CAB = PatternFill("solid", fgColor="DCE6F1")
TOT = PatternFill("solid", fgColor="FFF2CC")
NUM = "#,##0.00;[Red]-#,##0.00;-"
PCT = "0.00%;[Red]-0.00%;-"
CEN = Alignment(horizontal="center", vertical="center", wrap_text=True)


def cabecera(ws, fila, titulos):
    for c, t in enumerate(titulos, start=1):
        cell = ws.cell(fila, c, t)
        cell.font, cell.fill, cell.alignment = NEG, CAB, CEN


def crear(destino: Path | None = None) -> Path:
    res = todos()
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumen_mensual"
    ws["A1"] = "FRENOSA — SIMULACIÓN 2010: RESULTADOS MENSUALES DEL MODELO ABC POR ÓRDENES DE TRABAJO"
    ws["A1"].font = TIT
    ws["A2"] = ("Cifras calculadas por el sistema al importar cada Excel mensual. Julio = datos reales del caso; "
                "los demás meses son simulados (ficticios).")
    ws["A2"].font = NOR
    titulos = ["Mes", "Archivo", "Unidades producidas", "Ventas netas S/", "CMD S/", "Producción (MOD + CIF) S/",
               "Soporte S/", "Costo total S/", "Margen bruto S/", "Margen operativo S/", "Rentabilidad operativa",
               "Escenario"]
    cabecera(ws, 4, titulos)
    r0 = 5
    for i, r in enumerate(res):
        f = r0 + i
        mes = int(r["periodo_id"][-2:])
        valores = [r["periodo"], r["archivo"], r["unidades"], r["ventas"], r["cmd"], r["produccion"], r["soporte"],
                   r["costo_total"], r["margen_bruto"], r["margen_operativo"], f"=IFERROR(J{f}/D{f},0)",
                   ESCENARIOS[mes]["texto"]]
        for c, v in enumerate(valores, start=1):
            cell = ws.cell(f, c, v)
            cell.font = NOR
            if 3 <= c <= 10:
                cell.number_format = "#,##0" if c == 3 else NUM
        ws.cell(f, 11).number_format = PCT
    ft = r0 + len(res)
    ws.cell(ft, 1, "TOTAL 2010").font = NEG
    for c in range(3, 11):
        letra = "CDEFGHIJ"[c - 3]
        cell = ws.cell(ft, c, f"=SUM({letra}{r0}:{letra}{ft - 1})")
        cell.font, cell.fill, cell.number_format = NEG, TOT, "#,##0" if c == 3 else NUM
    cell = ws.cell(ft, 11, f"=IFERROR(J{ft}/D{ft},0)")
    cell.font, cell.fill, cell.number_format = NEG, TOT, PCT
    ws.cell(ft + 2, 1, "Meses con pérdida operativa:").font = NEG
    ws.cell(ft + 2, 3, f'=COUNTIF(J{r0}:J{ft - 1},"<0")').font = NEG
    ws.cell(ft + 3, 1, "Meses con mayor rentabilidad que julio:").font = NEG
    fila_julio = r0 + next(i for i, r in enumerate(res) if r["periodo_id"].endswith("07"))
    ws.cell(ft + 3, 3, f'=COUNTIF(K{r0}:K{ft - 1},">"&K{fila_julio})').font = NEG
    for c, w in zip("ABCDEFGHIJKL", [15, 38, 12, 16, 15, 16, 13, 16, 16, 17, 13, 90]):
        ws.column_dimensions[c].width = w
    ws.freeze_panes = "C5"

    graf = BarChart()
    graf.title = "Rentabilidad operativa por mes"
    graf.y_axis.title = "Rentabilidad"
    graf.y_axis.number_format = "0%"
    graf.add_data(Reference(ws, min_col=11, min_row=4, max_row=ft - 1), titles_from_data=True)
    graf.set_categories(Reference(ws, min_col=1, min_row=r0, max_row=ft - 1))
    graf.roundedCorners = False
    graf.legend = None
    graf.x_axis.delete = False
    graf.y_axis.delete = False
    graf.title.overlay = False
    graf.height, graf.width = 8, 22
    ws.add_chart(graf, f"A{ft + 6}")

    # Por OT
    wo = wb.create_sheet("Rentabilidad_por_OT")
    wo["A1"] = "RENTABILIDAD OPERATIVA POR OBJETO DE COSTO (OT) Y MES"
    wo["A1"].font = TIT
    productos = [o["codigo_producto"] for o in res[0]["por_ot"]]
    cabecera(wo, 3, ["Mes"] + [f"{p} unidades" for p in productos] + [f"{p} ventas S/" for p in productos]
             + [f"{p} margen op. S/" for p in productos] + [f"{p} rentabilidad" for p in productos])
    for i, r in enumerate(res):
        f = 4 + i
        wo.cell(f, 1, r["periodo"]).font = NOR
        n = len(productos)
        for j, o in enumerate(r["por_ot"]):
            for k, (v, fmt) in enumerate([(o["unidades"], "#,##0"), (o["ventas"], NUM), (o["margen_operativo"], NUM)]):
                cell = wo.cell(f, 2 + k * n + j, v)
                cell.font, cell.number_format = NOR, fmt
            ventas_col = 2 + n + j
            margen_col = 2 + 2 * n + j
            from openpyxl.utils import get_column_letter as L
            cell = wo.cell(f, 2 + 3 * n + j, f"=IFERROR({L(margen_col)}{f}/{L(ventas_col)}{f},0)")
            cell.font, cell.number_format = NOR, PCT
    for c in range(1, 2 + 4 * len(productos)):
        from openpyxl.utils import get_column_letter as L
        wo.column_dimensions[L(c)].width = 16 if c > 1 else 15
    wo.freeze_panes = "B4"

    # Kardex
    wk = wb.create_sheet("Kardex_compras")
    wk["A1"] = "KARDEX DE MATERIALES DIRECTOS (PROMEDIO PONDERADO) — SALDOS POR MES"
    wk["A1"].font = TIT
    wk["A2"] = ("El saldo final de cada mes es el saldo inicial del mes siguiente. "
                "Las salidas del mes son el CMD de cada OT.")
    wk["A2"].font = NOR
    cabecera(wk, 4, ["Mes", "Ítem", "Descripción", "Saldo inicial cant.", "Saldo inicial S/", "Compras netas S/",
                     "Salidas a OT S/", "Saldo final cant.", "Saldo final S/", "Comprobación S/", "Continuidad",
                     "OT", "CMD S/", "Salidas − CMD S/"])
    f = 5
    previo: dict[str, int] = {}
    for r in res:
        for k in r["kardex"]:
            valores = [r["periodo"], k["item"], k["descripcion"], k["si_q"], k["si_v"], k["compras"], k["consumos"],
                       k["sf_q"], k["sf_v"], f"=E{f}+F{f}-G{f}"]
            for c, v in enumerate(valores, start=1):
                cell = wk.cell(f, c, v)
                cell.font = NOR
                if c >= 4:
                    cell.number_format = "#,##0.00" if c in (4, 8) else NUM
            if k["item"] in previo:
                p = previo[k["item"]]
                wk.cell(f, 11, f'=IF(AND(D{f}=H{p},ROUND(E{f}-I{p},2)=0),"OK","Revisar")').font = NOR
            else:
                wk.cell(f, 11, "Inicio del año").font = NOR
            previo[k["item"]] = f
            wk.cell(f, 12, k["codigo_ot"]).font = NOR
            for c, v in ((13, k["cmd"]), (14, f"=ROUND(G{f}-M{f},2)")):
                cell = wk.cell(f, c, v)
                cell.font, cell.number_format = NOR, NUM
            f += 1
    for c, w in zip("ABCDEFGHIJKLMN", [15, 11, 36, 13, 14, 15, 15, 13, 14, 15, 13, 20, 14, 14]):
        wk.column_dimensions[c].width = w
    wk.freeze_panes = "D5"

    destino = destino or CARPETA / "Resumen_simulacion_2010.xlsx"
    wb.save(destino)
    return destino


if __name__ == "__main__":
    print(crear())
