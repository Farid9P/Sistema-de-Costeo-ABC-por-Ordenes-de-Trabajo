"""Reporte Excel del analisis por periodos (consolidado de varios meses) con formulas.

Hojas de detalle (datos mensuales calculados por el sistema, en azul) y hojas
consolidadas que los suman con SUMIFS, de modo que el Excel recalcula si se
cambia un dato mensual. Las razones (costo unitario, rentabilidad, peso AC-OT)
se calculan sobre los totales del rango.
"""

from __future__ import annotations

import re
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .analisis_periodos import (
    TIPOS_ANALISIS, _cargar_meses, _ot_por_mes, analizar, bloques, comparar_bloques, nombre_mes,
)
from .constants import REPORTS_DIR

FUENTE = "Arial"
F_TIT = Font(name=FUENTE, size=13, bold=True, color="1F4D78")
F_SUB = Font(name=FUENTE, size=10, italic=True, color="334155")
F_NOR = Font(name=FUENTE, size=10)
F_BOLD = Font(name=FUENTE, size=10, bold=True)
F_IN = Font(name=FUENTE, size=10, color="0000FF")
F_LINK = Font(name=FUENTE, size=10, color="008000")
F_HDR = Font(name=FUENTE, size=10, bold=True, color="FFFFFF")
FILL_HDR = PatternFill("solid", fgColor="1F4D78")
FILL_TOT = PatternFill("solid", fgColor="DCE6F1")
NUM = "#,##0.00;[Red](#,##0.00);-"
CU = "#,##0.0000;[Red](#,##0.0000);-"
PCT = "0.00%;[Red]-0.00%;-"
CANT = "#,##0;[Red](#,##0);-"
CENTRO = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _cabecera(ws, fila: int, titulos: list[str], anchos: list[int] | None = None) -> None:
    for col, titulo in enumerate(titulos, start=1):
        c = ws.cell(fila, col, titulo)
        c.font, c.fill, c.alignment = F_HDR, FILL_HDR, CENTRO
    ws.row_dimensions[fila].height = 30
    for col, ancho in enumerate(anchos or [], start=1):
        ws.column_dimensions[get_column_letter(col)].width = ancho


def _titulo(ws, texto: str, sub: str) -> None:
    ws["A1"] = texto
    ws["A1"].font = F_TIT
    ws["A2"] = sub
    ws["A2"].font = F_SUB


def _rango(hoja: str, col: str, n: int) -> str:
    return f"'{hoja}'!${col}$5:${col}${n}"


def generar_reporte_periodos(db_path: str | Path, meses: list[tuple[int, int]], tipo: str, titulo: str,
                             output_dir: str | Path = REPORTS_DIR, resultados: dict | None = None) -> Path:
    """``resultados``: resultados por mes a usar (p. ej. los de una simulacion); por defecto, los iniciales."""
    payload = analizar(db_path, meses, tipo, titulo, resultados=resultados)
    if not payload.get("meses_con_datos"):
        raise ValueError("No hay meses procesados en el rango elegido.")
    en_rango = set(payload["meses_con_datos"])
    tamano = TIPOS_ANALISIS.get(tipo)
    anio_comp = payload["meses_con_datos"][-1][0]
    tipo_comp = tipo if tamano and 1 < tamano < 12 else "Trimestral"
    base = _cargar_meses(db_path, [(anio_comp, m) for m in range(1, 13)])
    if resultados:
        base.update(resultados)
    comparativo = comparar_bloques(db_path, tipo_comp, anio_comp, resultados=base)
    meses_comp = {m for b in comparativo for m in b["_meses"] if m in base}
    datos = {m: base[m] for m in sorted(en_rango | meses_comp) if m in base}
    meses_detalle = sorted(datos)

    wb = Workbook()
    resumen = wb.active
    resumen.title = "Resumen"
    sub = (f"Periodo analizado: {payload['titulo']} | Tipo: {tipo} | Meses con datos: {payload['texto_meses']}"
           + (" | Incluye valores simulados" if resultados else ""))

    # ---------------------------------------------------------------- Detalle mensual OT
    dm = wb.create_sheet("Detalle_mensual_OT")
    _titulo(dm, "DETALLE MENSUAL POR OT (resultados del calculo ABC de cada mes)", "Azul = dato calculado por el sistema en el mes; negro = formula.")
    _cabecera(dm, 4, ["Año", "Mes", "Mes (texto)", "En rango", "Periodo", "OT del mes", "Producto", "Cantidad producida",
                      "CMD", "MOD (PD)", "Otros recursos", "Costo produccion", "Costo soporte", "CSD", "Costo total OT",
                      "Costo unitario", "Ventas netas", "Margen bruto", "Margen operativo", "Rentabilidad operativa"],
              [7, 6, 14, 9, 13, 26, 20, 13, 14, 13, 13, 14, 13, 9, 15, 11, 15, 14, 15, 11])
    r = 5
    for mes in meses_detalle:
        tablas = datos[mes]
        ventas = {e["ot_id"]: e.get("ventas_netas_soles") for e in tablas.get("estado_resultados_ot", [])}
        periodo_id = (tablas.get("periodo") or [{}])[0].get("periodo_id")
        for h in tablas.get("hoja_costos_ot", []):
            valores = [mes[0], mes[1], nombre_mes(*mes), "Sí" if mes in en_rango else "No", periodo_id, h.get("codigo_ot"),
                       h.get("codigo_producto"), h.get("cantidad_producida"), h.get("cmd_ot"), h.get("mod_ot"),
                       h.get("otros_recursos_ot"), h.get("costo_produccion_ot"), h.get("costo_soporte_ot"), h.get("csd_ot")]
            for col, v in enumerate(valores, start=1):
                c = dm.cell(r, col, v)
                c.font = F_IN if col >= 8 else F_NOR
            dm.cell(r, 15, f"=I{r}+L{r}+M{r}+N{r}")
            dm.cell(r, 16, f"=IFERROR(O{r}/H{r},0)")
            dm.cell(r, 17, ventas.get(h.get("ot_id"))).font = F_IN
            dm.cell(r, 18, f"=Q{r}-I{r}-L{r}")
            dm.cell(r, 19, f"=R{r}-M{r}")
            dm.cell(r, 20, f"=IFERROR(S{r}/Q{r},0)")
            for col in range(8, 21):
                c = dm.cell(r, col)
                if col >= 15 and col != 17:
                    c.font = F_NOR
                c.number_format = CANT if col == 8 else CU if col == 16 else PCT if col == 20 else NUM
            r += 1
    fin_dm = r - 1
    dm.freeze_panes = "H5"

    # ---------------------------------------------------------------- Detalle actividades
    da = wb.create_sheet("Detalle_actividades")
    _titulo(da, "DETALLE MENSUAL POR ACTIVIDAD (Proceso 1 EC-AC de cada mes)", "Azul = dato calculado por el sistema; negro = formula.")
    _cabecera(da, 4, ["Año", "Mes", "Mes (texto)", "En rango", "Actividad", "Codigo", "Nombre actividad", "PD (MOD) AC",
                      "Otros recursos AC", "Costo conversion AC", "Soporte AC", "Costo total AC"],
              [7, 6, 14, 9, 10, 10, 30, 14, 15, 15, 14, 15])
    r = 5
    for mes in meses_detalle:
        for a in datos[mes].get("actividad_total", []):
            valores = [mes[0], mes[1], nombre_mes(*mes), "Sí" if mes in en_rango else "No", a.get("actividad_id"),
                       a.get("codigo_actividad"), a.get("nombre_actividad"), a.get("costo_mod_ac"),
                       a.get("costo_otros_recursos_ac")]
            for col, v in enumerate(valores, start=1):
                da.cell(r, col, v).font = F_IN if col >= 8 else F_NOR
            da.cell(r, 10, f"=H{r}+I{r}").font = F_NOR
            da.cell(r, 11, a.get("costo_soporte_ac")).font = F_IN
            da.cell(r, 12, f"=J{r}+K{r}").font = F_NOR
            for col in range(8, 13):
                da.cell(r, col).number_format = NUM
            r += 1
    fin_da = r - 1
    da.freeze_panes = "H5"

    # ---------------------------------------------------------------- Detalle AC-OT
    dx = wb.create_sheet("Detalle_AC_OT")
    _titulo(dx, "DETALLE MENSUAL AC → OT (Proceso 2 de cada mes)", "Azul = dato calculado por el sistema.")
    _cabecera(dx, 4, ["Año", "Mes", "Mes (texto)", "En rango", "Actividad", "Nombre actividad", "OT del mes", "Producto",
                      "Costo conversion AC", "Costo conversion OT", "Costo soporte AC", "Costo soporte OT"],
              [7, 6, 14, 9, 10, 30, 26, 20, 15, 15, 15, 15])
    r = 5
    for mes in meses_detalle:
        tablas = datos[mes]
        ots = _ot_por_mes(tablas)
        soporte = {(f["ot_id"], f["actividad_id"]): f for f in tablas.get("soporte_ot_detalle", [])}
        claves = [(f["ot_id"], f["actividad_id"]) for f in tablas.get("produccion_ot_detalle", [])]
        claves += [k for k in soporte if k not in claves]
        produccion = {(f["ot_id"], f["actividad_id"]): f for f in tablas.get("produccion_ot_detalle", [])}
        for clave in claves:
            p, s = produccion.get(clave, {}), soporte.get(clave, {})
            base = p or s
            info = ots.get(clave[0], {})
            valores = [mes[0], mes[1], nombre_mes(*mes), "Sí" if mes in en_rango else "No", clave[1],
                       base.get("nombre_actividad"), info.get("codigo_ot"), base.get("codigo_producto") or info.get("codigo_producto"),
                       p.get("costo_conversion_ac"), p.get("costo_conversion_ot"), s.get("costo_soporte_ac"), s.get("costo_soporte_ot")]
            for col, v in enumerate(valores, start=1):
                c = dx.cell(r, col, v)
                c.font = F_IN if col >= 9 else F_NOR
                if col >= 9:
                    c.number_format = NUM
            r += 1
    fin_dx = r - 1
    dx.freeze_panes = "I5"

    def sumifs(hoja, col, fin, criterios: list[tuple[str, str]]) -> str:
        partes = [_rango(hoja, col, fin)]
        for c_col, valor in criterios:
            partes += [_rango(hoja, c_col, fin), valor]
        return "SUMIFS(" + ",".join(partes) + ")"

    # ---------------------------------------------------------------- Hoja de costos consolidada
    hc = wb.create_sheet("Hoja_costos")
    _titulo(hc, f"HOJA DE COSTOS POR OT — {payload['titulo']}", "Suma de las OT mensuales de cada producto (SUMIFS sobre Detalle_mensual_OT, En rango = Sí).")
    _cabecera(hc, 4, ["OT (linea)", "Producto", "OT incluidas", "Cantidad producida", "CMD", "MOD (PD)", "Otros recursos",
                      "Costo produccion", "Costo soporte", "CSD", "Costo total OT", "Costo unitario OT"],
              [22, 20, 48, 14, 15, 14, 14, 15, 14, 10, 16, 13])
    cols_dm = {"D": "H", "E": "I", "F": "J", "G": "K", "H": "L", "I": "M", "J": "N"}
    r = 5
    for fila in payload["hoja"]:
        hc.cell(r, 1, fila["codigo_ot"]).font = F_NOR
        hc.cell(r, 2, fila["codigo_producto"]).font = F_NOR
        hc.cell(r, 3, fila["ot_incluidas"]).font = F_NOR
        for destino, origen in cols_dm.items():
            c = hc[f"{destino}{r}"]
            c.value = "=" + sumifs("Detalle_mensual_OT", origen, fin_dm, [("G", f"$B{r}"), ("D", '"Sí"')])
            c.font, c.number_format = F_LINK, CANT if destino == "D" else NUM
        hc[f"K{r}"] = f"=E{r}+H{r}+I{r}+J{r}"
        hc[f"L{r}"] = f"=IFERROR(K{r}/D{r},0)"
        hc[f"K{r}"].number_format, hc[f"L{r}"].number_format = NUM, CU
        r += 1
    fin_hc = r - 1
    hc.cell(r, 1, "TOTAL").font = F_BOLD
    for col in "DEFGHIJK":
        c = hc[f"{col}{r}"]
        c.value = f"=SUM({col}5:{col}{fin_hc})"
        c.font, c.fill, c.number_format = F_BOLD, FILL_TOT, CANT if col == "D" else NUM
    hc[f"L{r}"] = f"=IFERROR(K{r}/D{r},0)"
    hc[f"L{r}"].font, hc[f"L{r}"].fill, hc[f"L{r}"].number_format = F_BOLD, FILL_TOT, CU
    tot_hc = r
    hc.cell(r + 2, 1, "Costo unitario OT = costo total del rango / cantidad producida del rango.").font = F_SUB
    hc.freeze_panes = "D5"

    # ---------------------------------------------------------------- Estado de resultados consolidado
    er = wb.create_sheet("Estado_resultados")
    _titulo(er, f"ESTADO DE RESULTADOS POR OT — {payload['titulo']}", "Suma de los meses en rango; rentabilidad = margen operativo / ventas del rango.")
    _cabecera(er, 4, ["OT (linea)", "Producto", "Ventas netas", "Materia prima directa", "Costos de conversion", "Margen bruto",
                      "Procesos de soporte", "Margen operativo", "Rentabilidad operativa"],
              [22, 20, 16, 16, 16, 15, 15, 16, 13])
    r = 5
    for fila in payload["estado"]:
        er.cell(r, 1, fila["codigo_ot"]).font = F_NOR
        er.cell(r, 2, fila["codigo_producto"]).font = F_NOR
        for destino, origen in (("C", "Q"), ("D", "I"), ("E", "L"), ("G", "M")):
            c = er[f"{destino}{r}"]
            c.value = "=" + sumifs("Detalle_mensual_OT", origen, fin_dm, [("G", f"$B{r}"), ("D", '"Sí"')])
            c.font, c.number_format = F_LINK, NUM
        er[f"F{r}"] = f"=C{r}-D{r}-E{r}"
        er[f"H{r}"] = f"=F{r}-G{r}"
        er[f"I{r}"] = f"=IFERROR(H{r}/C{r},0)"
        er[f"F{r}"].number_format = er[f"H{r}"].number_format = NUM
        er[f"I{r}"].number_format = PCT
        r += 1
    fin_er = r - 1
    er.cell(r, 1, "TOTAL").font = F_BOLD
    for col in "CDEFGH":
        c = er[f"{col}{r}"]
        c.value = f"=SUM({col}5:{col}{fin_er})"
        c.font, c.fill, c.number_format = F_BOLD, FILL_TOT, NUM
    er[f"I{r}"] = f"=IFERROR(H{r}/C{r},0)"
    er[f"I{r}"].font, er[f"I{r}"].fill, er[f"I{r}"].number_format = F_BOLD, FILL_TOT, PCT
    tot_er = r
    er.freeze_panes = "C5"

    # ---------------------------------------------------------------- Proceso 1
    p1 = wb.create_sheet("Proceso1_EC_AC")
    _titulo(p1, f"PROCESO 1: EC → AC — {payload['titulo']}", "Suma mensual por actividad (SUMIFS sobre Detalle_actividades).")
    _cabecera(p1, 4, ["Actividad", "Codigo", "Nombre actividad", "PD (MOD) AC", "Otros recursos AC", "Costo conversion AC",
                      "Soporte AC", "Costo total AC", "% del total"], [10, 10, 30, 15, 15, 16, 15, 16, 10])
    r = 5
    for fila in payload["proceso1"]:
        for col, v in enumerate([fila["actividad_id"], fila["codigo_actividad"], fila["nombre_actividad"]], start=1):
            p1.cell(r, col, v).font = F_NOR
        for destino, origen in (("D", "H"), ("E", "I"), ("G", "K")):
            c = p1[f"{destino}{r}"]
            c.value = "=" + sumifs("Detalle_actividades", origen, fin_da, [("E", f"$A{r}"), ("D", '"Sí"')])
            c.font, c.number_format = F_LINK, NUM
        p1[f"F{r}"] = f"=D{r}+E{r}"
        p1[f"H{r}"] = f"=F{r}+G{r}"
        p1[f"F{r}"].number_format = p1[f"H{r}"].number_format = NUM
        r += 1
    fin_p1 = r - 1
    p1.cell(r, 3, "TOTAL").font = F_BOLD
    for col in "DEFGH":
        c = p1[f"{col}{r}"]
        c.value = f"=SUM({col}5:{col}{fin_p1})"
        c.font, c.fill, c.number_format = F_BOLD, FILL_TOT, NUM
    for i in range(5, fin_p1 + 1):
        p1[f"I{i}"] = f"=IFERROR(H{i}/$H${r},0)"
        p1[f"I{i}"].number_format = PCT

    # ---------------------------------------------------------------- Proceso 2
    for nombre, filas, campos, titulo2 in (
        ("Proceso2_Produccion", payload["proceso2_produccion"], ("I", "J"), "PROCESO 2: AC → OT (PRODUCCION)"),
        ("Proceso2_Soporte", payload["proceso2_soporte"], ("K", "L"), "PROCESO 2: AC → OT (SOPORTE)"),
    ):
        ws = wb.create_sheet(nombre)
        _titulo(ws, f"{titulo2} — {payload['titulo']}", "Suma mensual por actividad y producto; peso efectivo = costo asignado a la OT / costo de la actividad.")
        _cabecera(ws, 4, ["Actividad", "Nombre actividad", "OT (linea)", "Producto", "Costo AC", "Costo asignado a OT", "Peso efectivo AC-OT"],
                  [10, 30, 22, 20, 16, 17, 13])
        r = 5
        for fila in filas:
            for col, v in enumerate([fila["actividad_id"], fila["nombre_actividad"], fila["codigo_ot"], fila["codigo_producto"]], start=1):
                ws.cell(r, col, v).font = F_NOR
            for destino, origen in (("E", campos[0]), ("F", campos[1])):
                c = ws[f"{destino}{r}"]
                c.value = "=" + sumifs("Detalle_AC_OT", origen, fin_dx, [("E", f"$A{r}"), ("H", f"$D{r}"), ("D", '"Sí"')])
                c.font, c.number_format = F_LINK, NUM
            ws[f"G{r}"] = f"=IFERROR(F{r}/E{r},0)"
            ws[f"G{r}"].number_format = "0.0000"
            r += 1
        ws.cell(r, 2, "TOTAL asignado a OT").font = F_BOLD
        c = ws[f"F{r}"]
        c.value = f"=SUM(F5:F{r - 1})"
        c.font, c.fill, c.number_format = F_BOLD, FILL_TOT, NUM
        ws.freeze_panes = "E5"

    # ---------------------------------------------------------------- Evolucion mensual
    ev = wb.create_sheet("Evolucion_mensual")
    _titulo(ev, f"EVOLUCION MENSUAL — {payload['titulo']}", "SUMIFS por año y mes sobre Detalle_mensual_OT.")
    _cabecera(ev, 4, ["Año", "Mes", "Mes (texto)", "Unidades", "Ventas netas", "CMD", "Costo produccion", "Costo soporte",
                      "Costo total", "Margen bruto", "Margen operativo", "Rentabilidad operativa", "Costo unitario medio"],
              [7, 6, 15, 12, 15, 14, 15, 14, 15, 14, 15, 12, 12])
    r = 5
    for mes in payload["meses_con_datos"]:
        ev.cell(r, 1, mes[0]).font = F_NOR
        ev.cell(r, 2, mes[1]).font = F_NOR
        ev.cell(r, 3, nombre_mes(*mes)).font = F_NOR
        for destino, origen in (("D", "H"), ("E", "Q"), ("F", "I"), ("G", "L"), ("H", "M"), ("I", "O")):
            c = ev[f"{destino}{r}"]
            c.value = "=" + sumifs("Detalle_mensual_OT", origen, fin_dm, [("A", f"$A{r}"), ("B", f"$B{r}")])
            c.font, c.number_format = F_LINK, CANT if destino == "D" else NUM
        ev[f"J{r}"] = f"=E{r}-F{r}-G{r}"
        ev[f"K{r}"] = f"=J{r}-H{r}"
        ev[f"L{r}"] = f"=IFERROR(K{r}/E{r},0)"
        ev[f"M{r}"] = f"=IFERROR(I{r}/D{r},0)"
        for col, fmt in (("J", NUM), ("K", NUM), ("L", PCT), ("M", CU)):
            ev[f"{col}{r}"].number_format = fmt
        r += 1
    fin_ev = r - 1
    ev.cell(r, 3, "TOTAL").font = F_BOLD
    for col in "DEFGHIJK":
        c = ev[f"{col}{r}"]
        c.value = f"=SUM({col}5:{col}{fin_ev})"
        c.font, c.fill, c.number_format = F_BOLD, FILL_TOT, CANT if col == "D" else NUM
    ev[f"L{r}"] = f"=IFERROR(K{r}/E{r},0)"
    ev[f"M{r}"] = f"=IFERROR(I{r}/D{r},0)"
    for col, fmt in (("L", PCT), ("M", CU)):
        ev[f"{col}{r}"].font, ev[f"{col}{r}"].fill, ev[f"{col}{r}"].number_format = F_BOLD, FILL_TOT, fmt
    _grafico_evolucion(ev, fin_ev)

    # ---------------------------------------------------------------- Comparativo
    if comparativo:
        cp = wb.create_sheet("Comparativo")
        _titulo(cp, f"COMPARATIVO {tipo_comp.upper()} {anio_comp}",
                "SUMIFS por año y rango de meses sobre Detalle_mensual_OT (todos los meses procesados del año).")
        _cabecera(cp, 4, ["Subperiodo", "Año", "Mes inicial", "Mes final", "Meses con datos", "Unidades", "Ventas netas",
                          "Costo total", "Margen bruto", "Margen operativo", "Rentabilidad operativa", "Estado"],
                  [40, 7, 9, 9, 10, 12, 16, 16, 15, 16, 12, 11])
        r = 5
        for b in comparativo:
            ini, fin = b["_meses"][0][1], b["_meses"][-1][1]
            for col, v in enumerate([b["subperiodo"], anio_comp, ini, fin, b["meses_con_datos"]], start=1):
                cp.cell(r, col, v).font = F_NOR
            crit = [("A", f"$B{r}"), ("B", f'">="&$C{r}'), ("B", f'"<="&$D{r}')]
            for destino, origen in (("F", "H"), ("G", "Q"), ("H", "O")):
                c = cp[f"{destino}{r}"]
                c.value = "=" + sumifs("Detalle_mensual_OT", origen, fin_dm, crit)
                c.font, c.number_format = F_LINK, CANT if destino == "F" else NUM
            cp[f"I{r}"] = "=" + sumifs("Detalle_mensual_OT", "R", fin_dm, crit)
            cp[f"J{r}"] = "=" + sumifs("Detalle_mensual_OT", "S", fin_dm, crit)
            cp[f"K{r}"] = f"=IFERROR(J{r}/G{r},0)"
            cp[f"L{r}"] = b["estado"]
            for col, fmt in (("I", NUM), ("J", NUM), ("K", PCT)):
                cp[f"{col}{r}"].number_format = fmt
            r += 1
        _grafico_comparativo(cp, r - 1)

    # ---------------------------------------------------------------- Kardex
    kx = wb.create_sheet("Kardex_materiales")
    _titulo(kx, f"KARDEX DE MATERIALES DIRECTOS (PROMEDIO PONDERADO) — {payload['titulo']}",
            "Saldo inicial del primer mes, compras y salidas sumadas (al centimo de cada mes), saldo final del ultimo mes. "
            "Ajuste por redondeo = SI + compras − salidas − SF (centimos por registrar cada mes al centimo).")
    _cabecera(kx, 4, ["Material", "Descripcion", "Unidad", "Saldo inicial cant.", "Saldo inicial S/", "Compras cant.", "Compras S/",
                      "Salidas a OT cant.", "Salidas a OT S/", "Saldo final cant.", "Saldo final S/", "Ajuste por redondeo S/",
                      "CMD del rango S/", "Salidas − CMD S/", "Continuidad"],
              [10, 20, 7, 13, 14, 12, 15, 13, 15, 13, 14, 13, 15, 13, 12])
    r = 5
    for k in payload.get("kardex", []):
        valores = [k["material_id"], k["material"], k["unidad"], k["saldo_inicial_cantidad"], k["saldo_inicial_soles"],
                   k["compras_cantidad"], k["compras_soles"], k["salidas_cantidad"], k["salidas_soles"],
                   k.get("saldo_final_cantidad"), k.get("saldo_final_soles")]
        for col, v in enumerate(valores, start=1):
            c = kx.cell(r, col, v)
            c.font = F_IN if col >= 4 else F_NOR
            if col >= 4:
                c.number_format = NUM
        kx[f"L{r}"] = f"=ROUND(E{r}+G{r}-I{r}-K{r},2)"
        kx[f"M{r}"] = k["cmd_soles"]
        kx[f"M{r}"].font = F_IN
        kx[f"N{r}"] = f"=ROUND(I{r}-M{r},2)"
        kx[f"O{r}"] = k["continuidad"]
        for col in "LMN":
            kx[f"{col}{r}"].number_format = NUM
        r += 1
    if r == 5:
        kx["A5"] = "Los meses procesados no traen hoja Compras."

    # ---------------------------------------------------------------- Resumen
    _titulo(resumen, "ANALISIS POR PERIODOS — COSTEO ABC POR ORDENES DE TRABAJO (FRENOSA)", sub)
    filas = [
        ("Unidades producidas", f"='Hoja_costos'!D{tot_hc}", CANT),
        ("Ventas netas (S/)", f"='Estado_resultados'!C{tot_er}", NUM),
        ("Costo total OT (S/)", f"='Hoja_costos'!K{tot_hc}", NUM),
        ("Costo unitario medio (S/)", f"=IFERROR(B6/B4,0)", CU),
        ("Costo unitario promedio de las OT (S/)", f"=IFERROR(AVERAGE('Hoja_costos'!L5:L{fin_hc}),0)", CU),
        ("Margen bruto (S/)", f"='Estado_resultados'!F{tot_er}", NUM),
        ("Margen operativo (S/)", f"='Estado_resultados'!H{tot_er}", NUM),
        ("Rentabilidad operativa global", f"=IFERROR(B10/B5,0)", PCT),
        ("Rentabilidad promedio de las OT", f"=IFERROR(AVERAGE('Estado_resultados'!I5:I{fin_er}),0)", PCT),
    ]
    for i, (concepto, formula, fmt) in enumerate(filas, start=4):
        resumen.cell(i, 1, concepto).font = F_NOR
        c = resumen.cell(i, 2, formula)
        c.font, c.number_format = F_LINK, fmt
    resumen.column_dimensions["A"].width = 42
    resumen.column_dimensions["B"].width = 20
    fila = 15
    resumen.cell(fila, 1, "Meses incluidos").font = F_BOLD
    resumen.cell(fila, 2, payload["texto_meses"]).font = F_NOR
    if payload.get("meses_faltantes"):
        resumen.cell(fila + 1, 1, "Meses faltantes (sin procesar)").font = F_BOLD
        resumen.cell(fila + 1, 2, ", ".join(nombre_mes(a, m) for a, m in payload["meses_faltantes"])).font = Font(name=FUENTE, size=10, color="FF0000")
    fila += 3
    resumen.cell(fila, 1, "Alertas").font = F_BOLD
    for i, alerta in enumerate(payload.get("alertas", []), start=fila + 1):
        resumen.cell(i, 1, f"[{alerta['severidad']}] {alerta['tipo']}").font = F_NOR
        resumen.cell(i, 2, alerta["mensaje"]).font = F_NOR
    fila = fila + len(payload.get("alertas", [])) + 2
    notas = [
        "Notas:",
        "Las hojas Detalle_* tienen los resultados del calculo ABC de cada mes (azul); las demas hojas los suman con formulas.",
        "Las OT de cada mes se agrupan por producto: la OT (linea) del rango es la suma de las OT mensuales de ese producto.",
        "Costo unitario y rentabilidad se calculan con los totales del rango (no son promedio de los meses).",
        "Fuente: " + payload.get("fuente", ""),
    ]
    for i, texto in enumerate(notas, start=fila):
        resumen.cell(i, 1, texto).font = F_BOLD if texto == "Notas:" else F_NOR

    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = ws.title.startswith("Detalle") or ws.title == "Resumen"
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    slug = re.sub(r"[^A-Za-z0-9]+", "_", _sin_tildes(payload["titulo"])).strip("_")
    ruta = output_dir / f"Reporte_Analisis_Periodos_{slug}.xlsx"
    wb.save(ruta)
    return ruta


def _sin_tildes(texto: str) -> str:
    import unicodedata

    return "".join(c for c in unicodedata.normalize("NFKD", texto) if not unicodedata.combining(c))


def _grafico_evolucion(ws, fin: int) -> None:
    if fin < 5:
        return
    from openpyxl.chart import BarChart, LineChart, Reference

    barras = BarChart()
    barras.title = "Ventas, costo total y margen operativo por mes"
    barras.y_axis.title = "Soles"
    for col in (5, 9, 11):
        barras.add_data(Reference(ws, min_col=col, min_row=4, max_row=fin), titles_from_data=True)
    barras.set_categories(Reference(ws, min_col=3, min_row=5, max_row=fin))
    _ajustar_grafico(barras)
    barras.width, barras.height = 24, 9
    ws.add_chart(barras, f"B{fin + 4}")
    linea = LineChart()
    linea.title = "Rentabilidad operativa por mes"
    linea.add_data(Reference(ws, min_col=12, min_row=4, max_row=fin), titles_from_data=True)
    linea.set_categories(Reference(ws, min_col=3, min_row=5, max_row=fin))
    linea.y_axis.number_format = "0%"
    _ajustar_grafico(linea)
    linea.width, linea.height = 24, 8
    ws.add_chart(linea, f"B{fin + 23}")


def _grafico_comparativo(ws, fin: int) -> None:
    from openpyxl.chart import BarChart, Reference

    graf = BarChart()
    graf.title = "Ventas, costo total y margen operativo por subperiodo"
    for col in (7, 8, 10):
        graf.add_data(Reference(ws, min_col=col, min_row=4, max_row=fin), titles_from_data=True)
    graf.set_categories(Reference(ws, min_col=1, min_row=5, max_row=fin))
    _ajustar_grafico(graf)
    graf.width, graf.height = 24, 9
    ws.add_chart(graf, f"A{fin + 3}")


def _ajustar_grafico(graf) -> None:
    graf.roundedCorners = False
    for titulo in (graf.title, graf.x_axis.title, graf.y_axis.title):
        if titulo is not None and hasattr(titulo, "overlay"):
            titulo.overlay = False
    graf.x_axis.delete = False
    graf.y_axis.delete = False
    if graf.legend is not None:
        graf.legend.overlay = False
        graf.legend.position = "b"
