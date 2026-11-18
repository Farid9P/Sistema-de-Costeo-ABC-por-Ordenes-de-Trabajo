"""Genera Excel maestros simulados para enero a diciembre de 2010.

Toma como guía el Excel maestro del caso (julio 2010): misma estructura de
hojas y columnas. Julio conserva los datos reales del caso; los demás meses
son datos ficticios con escenarios (estacionalidad, precios, alza del material,
reparaciones) para que haya meses con más rentabilidad operativa que julio y
meses con rentabilidad negativa.

Uso:
    python tools/simular_periodos_2010.py                   # genera los 12 archivos
    python tools/simular_periodos_2010.py --indices precios.json

Los archivos se guardan en data/simulacion_2010/. La hoja Compras mantiene la
continuidad del kardex: el saldo final de cada mes es el saldo inicial del
siguiente (incluido junio → julio → agosto).
"""

from __future__ import annotations

import argparse
import json
import random
import sys
from copy import copy
from datetime import date
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from app.kardex_compras import valorizar  # noqa: E402

MAESTRO = ROOT / "data" / "frenosa_datos_entrada_abc.xlsx"
SALIDA = ROOT / "data" / "simulacion_2010"
MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre",
         "Noviembre", "Diciembre"]
ULTIMO_DIA = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
IGV = 0.19  # IGV vigente en el Perú durante 2010 (16 % + 3 % IPM)
FUENTE_SIM = ("Simulación académica 2010 (datos ficticios) elaborada con la estructura del caso Frenosa de julio 2010; "
              "Cherres Juárez (2010), Contabilidad y Negocios, PUCP.")

# Escenarios: volumen por OT (SW50, FRE9000, FRE9002) respecto de julio, índice del costo del material,
# operarios, eficiencia (horas reales / horas estándar), reparaciones extraordinarias y rentabilidad objetivo.
ESCENARIOS = {
    1: dict(vol=(0.37, 0.35, 0.38), mat=1.00, operarios=52, efic=0.97, extra={}, objetivo=-0.06,
            texto="Temporada baja por vacaciones: la demanda cae a cerca de un tercio de julio y se ofrecen descuentos; "
                  "la planilla y los demás costos fijos se mantienen, por eso hay pérdida operativa."),
    2: dict(vol=(0.34, 0.32, 0.35), mat=1.01, operarios=52, efic=1.03, extra={"MAQ004": 12500}, objetivo=-0.14,
            texto="Demanda mínima, descuentos y reparación mayor de la prensa hidráulica (S/ 12,500): el peor mes del año."),
    3: dict(vol=(0.62, 0.60, 0.64), mat=1.00, operarios=56, efic=0.98, extra={}, objetivo=0.13,
            texto="Recuperación de la demanda con precios todavía bajos y capacidad ociosa."),
    4: dict(vol=(0.82, 0.80, 0.86), mat=0.99, operarios=58, efic=0.97, extra={}, objetivo=0.22,
            texto="Demanda en crecimiento y material estable."),
    5: dict(vol=(0.94, 0.92, 0.95), mat=0.99, operarios=59, efic=0.96, extra={}, objetivo=0.25,
            texto="Producción cercana al nivel de julio."),
    6: dict(vol=(1.02, 1.00, 0.98), mat=1.00, operarios=59, efic=0.95, extra={}, objetivo=0.30,
            texto="Buen mes: volumen similar a julio con mejores precios."),
    7: dict(vol=(1.0, 1.0, 1.0), mat=1.0, operarios=59, efic=None, extra={}, objetivo=None,
            texto="Datos reales del caso Frenosa (julio 2010), sin cambios."),
    8: dict(vol=(1.08, 1.10, 1.02), mat=1.00, operarios=60, efic=0.95, extra={}, objetivo=0.32,
            texto="Mayor demanda de reposición de flotas: más volumen y precios firmes."),
    9: dict(vol=(1.15, 1.18, 1.05), mat=0.98, operarios=61, efic=0.94, extra={}, objetivo=0.36,
            texto="Mejor mes del año: volumen alto, material más barato y buenos precios."),
    10: dict(vol=(1.10, 1.12, 1.00), mat=1.02, operarios=61, efic=0.95, extra={}, objetivo=0.31,
             texto="Demanda alta; el material empieza a subir."),
    11: dict(vol=(0.80, 0.78, 0.85), mat=1.30, operarios=60, efic=1.00, extra={"MAQ006": 4200}, objetivo=-0.05,
             texto="Alza de 30 % en el material importado, guerra de precios y cambio de rodillos de la lijadora: "
                   "pérdida operativa."),
    12: dict(vol=(0.90, 0.88, 0.92), mat=1.12, operarios=59, efic=0.97, extra={}, objetivo=0.16,
             texto="Recuperación parcial de precios; el material sigue caro."),
}

# Comportamiento de los otros recursos: parte fija y parte variable con el volumen (kg)
VARIABILIDAD_RECURSOS = {"EC001": 0.0, "EC002": 0.6, "EC003": 0.7, "EC004": 0.1, "EC005": 0.0}


def r2(x: float) -> float:
    return float(Decimal(str(x)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def rh(x: float, d: int = 0) -> float:
    return float(Decimal(str(x)).quantize(Decimal(1).scaleb(-d), rounding=ROUND_HALF_UP))


def filas(ws):
    hdr = [c.value for c in ws[1]]
    return hdr, [row for row in ws.iter_rows(min_row=2) if row[0].value not in (None, "")]


def col(hdr, nombre):
    return hdr.index(nombre)


def mayor_residuo(valores: list[float], total: int = 100) -> list[float]:
    """Redondea porcentajes a 2 decimales manteniendo la suma en 1.00."""
    suma = sum(valores) or 1.0
    crudos = [v / suma * total for v in valores]
    base = [int(x) for x in crudos]
    faltan = total - sum(base)
    orden = sorted(range(len(crudos)), key=lambda i: crudos[i] - base[i], reverse=True)
    for i in orden[:faltan]:
        base[i] += 1
    return [b / 100 for b in base]


# ---------------------------------------------------------------------------
# Kardex de materiales directos (hoja Compras) con continuidad enero-diciembre
# ---------------------------------------------------------------------------
# Regla: las salidas del mes a cada OT (cantidad y costo) son el CMD de la OT:
# cantidad x costo unitario del material (Cuadro 1 del caso en julio; simulados en
# los demas meses, tools/materiales_directos_base_julio_2010.json). El kardex es la
# unica fuente del CMD (ya no existe la hoja Materiales_Directos); la diferencia con
# el CMD objetivo es menor a S/ 0.000001.
TOLERANCIA_CMD = 1e-6
MATERIALES_BASE = json.loads((ROOT / "tools" / "materiales_directos_base_julio_2010.json").read_text(encoding="utf-8"))
ALMACEN_MP = "Almacén de materias primas"
PROVEEDOR = {"MAT001": ("001", 5200), "MAT002": ("002", 3300), "MAT003": ("003", 1800)}
STOCK_FINAL = 0.15  # saldo final objetivo: 15 % del consumo del mes


def _p4(x: float) -> float:
    return float(Decimal(str(x)).quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP))


def _mov(mes, dia, mat, tipo, cantidad, precio=None, documento="", destino="", importe=None, detalle="", obs="",
         flete=None):
    return dict(periodo_id=f"PER2010-{mes:02d}", fecha=date(2010, mes, dia), item_codigo=mat["material_id"],
                item_descripcion=mat["material"], almacen=ALMACEN_MP, unidad="kg", tipo_movimiento=tipo,
                cantidad=float(cantidad), precio_unitario_con_igv=precio, igv_pct=IGV if precio else None,
                flete_soles=flete, importe_soles=importe, documento=documento, destino=destino, detalle=detalle,
                observacion=obs or None)


def _salidas(movs):
    k = valorizar(movs)[0]
    return k, sum(f.sal_total or 0 for f in k.filas)


def kardex_mes(mes: int, mat: dict, saldo: tuple[float, float], rnd: random.Random, vales: list[int],
               facturas: dict) -> list[dict]:
    """Movimientos de un material en el mes; las salidas suman exactamente el CMD de su OT."""
    c_total = int(mat["cantidad_kg"])
    cu = float(mat["costo_unitario_soles"])
    objetivo = c_total * cu
    ot = mat["codigo_ot"]
    ultimo = ULTIMO_DIA[mes - 1]
    q0, v0 = saldo
    c1 = int(round(c_total * rnd.uniform(0.30, 0.38)))
    c2 = int(round(c_total * rnd.uniform(0.30, 0.38)))
    c3 = c_total - c1 - c2
    qa = max(0, c1 - int(q0)) + int(round(c_total * rnd.uniform(0.12, 0.20)))
    dev = int(round(qa * rnd.uniform(0.01, 0.03))) if rnd.random() < 0.3 and qa > 100 else 0
    disponible_tras_c1 = q0 + qa - dev - c1
    qb_base = max(1, int(c2 + c3 - disponible_tras_c1 + round(c_total * rnd.uniform(0.03, 0.08))))
    pa = _p4(cu * rnd.uniform(0.98, 1.02) * (1 + IGV))
    serie, _ = PROVEEDOR[mat["material_id"]]

    def factura():
        facturas[mat["material_id"]] += rnd.randint(3, 12)
        return f"F/ {serie}-{facturas[mat['material_id']]}"

    def vale():
        vales[0] += 1
        return f"VS 2010-{vales[0]:03d}"

    doc_a, doc_b, doc_c = factura(), factura(), factura()
    vs = [vale(), vale(), vale()]

    flete_a = r2(qa * rnd.uniform(0.02, 0.04))

    def armar(qb, pb, c1_, c2_, flete_b):
        c3_ = c_total - c1_ - c2_
        movs = [_mov(mes, 1, mat, "Saldo inicial", q0, importe=v0, detalle=f"{q0:,.0f} kg a S/ {v0:,.2f}",
                     obs="Saldo final del mes anterior.")]
        movs.append(_mov(mes, 2, mat, "Compra", qa, pa, doc_a, flete=flete_a,
                         detalle=f"{qa:,.0f} kg, costo compra unitario S/ {pa:,.4f} con IGV + flete S/ {flete_a:,.2f}"))
        if dev:
            movs.append(_mov(mes, 4, mat, "Devolucion a proveedor", dev, pa, doc_a,
                             detalle=f"{dev:,.0f} kg, devolución al proveedor de la compra del 02/{mes:02d}/2010",
                             obs="Material fuera de especificación; se devuelve al costo unitario de esa compra."))
        movs.append(_mov(mes, 5, mat, "Consumo", c1_, documento=vs[0], destino=ot,
                         detalle=f"{c1_:,.0f} kg, salida a la {ot} ({vs[0]})"))
        movs.append(_mov(mes, 12, mat, "Compra", qb, pb, doc_b, flete=flete_b,
                         detalle=f"{qb:,.0f} kg, costo compra unitario S/ {pb:,.4f} con IGV + flete S/ {flete_b:,.2f}"))
        movs.append(_mov(mes, 15, mat, "Consumo", c2_, documento=vs[1], destino=ot,
                         detalle=f"{c2_:,.0f} kg, salida a la {ot} ({vs[1]})"))
        movs.append(_mov(mes, 22, mat, "Consumo", c3_, documento=vs[2], destino=ot,
                         detalle=f"{c3_:,.0f} kg, salida a la {ot} ({vs[2]})"))
        return movs

    def s_de(*args):
        return _salidas(armar(*args))[1]

    import math

    def reposicion(movs, k):
        q_fin = k.saldo_final_cantidad
        qc = max(0, int(round(c_total * STOCK_FINAL - q_fin)))
        if qc > 0:
            pc = _p4(cu * rnd.uniform(0.99, 1.03) * (1 + IGV))
            fc = r2(qc * rnd.uniform(0.02, 0.04))
            movs.append(_mov(mes, min(26, ultimo - 2), mat, "Compra", qc, pc, doc_c, flete=fc,
                             detalle=f"{qc:,.0f} kg, costo compra unitario S/ {pc:,.4f} con IGV + flete S/ {fc:,.2f}",
                             obs="Reposición para el mes siguiente."))
        return movs

    # S (salidas valorizadas) es lineal en el flete de la compra B: se elige el centimo que deja el
    # residuo mas chico y se prueban otras cantidades de compra/consumo hasta que sea < TOLERANCIA_CMD.
    for dq in [0] + [d for k in range(1, 1500) for d in (k, -k)]:
        qb = qb_base + dq
        if qb < 1:
            continue
        for dc in [0] + [d for k in range(1, 25) for d in (k, -k)]:
            c1_ = c1 + dc
            if c1_ < 1 or c_total - c1_ - c2 < 1:
                continue
            flete0 = r2(qb * 0.03)
            s1, s2 = s_de(qb, 1.0, c1_, c2, flete0), s_de(qb, 2.0, c1_, c2, flete0)
            pb_cont = 1.0 + (objetivo - s1) / (s2 - s1)
            pb = math.floor(pb_cont * 10000) / 10000
            if not (0.85 * cu * (1 + IGV) <= pb <= 1.15 * cu * (1 + IGV)):
                continue
            base = s_de(qb, pb, c1_, c2, flete0)
            pendiente = (s_de(qb, pb, c1_, c2, flete0 + 10) - base) / 10
            if pendiente <= 0:
                continue
            exacto = flete0 + (objetivo - base) / pendiente
            for fb in sorted({r2(exacto), r2(exacto + 0.01), r2(exacto - 0.01)}, key=lambda f: abs(f - exacto)):
                if fb < 0 or fb > qb * 0.08:
                    continue
                if abs(base + pendiente * (fb - flete0) - objetivo) > 50 * TOLERANCIA_CMD:
                    continue
                movs = armar(qb, pb, c1_, c2, fb)
                k, s_total = _salidas(movs)
                if abs(s_total - objetivo) < TOLERANCIA_CMD and not k.avisos:
                    return reposicion(movs, k)
    raise RuntimeError(f"No se encontró una compra que cuadre el CMD de {mat['material']} en el mes {mes}")


def simular_kardex_materiales(materiales_por_mes: dict[int, list[dict]]) -> dict[int, list[dict]]:
    rnd = random.Random(2010)
    facturas = {k: v[1] for k, v in PROVEEDOR.items()}
    vales = [0]
    saldos: dict[str, tuple[float, float]] = {}
    salida: dict[int, list[dict]] = {}
    for mes in range(1, 13):
        salida[mes] = []
        for mat in materiales_por_mes[mes]:
            mid = mat["material_id"]
            if mid not in saldos:
                q0 = int(round(float(mat["cantidad_kg"]) * 0.15))
                saldos[mid] = (float(q0), r2(q0 * float(mat["costo_unitario_soles"]) * 0.99))
            movs = kardex_mes(mes, mat, saldos[mid], rnd, vales, facturas)
            k = valorizar(movs)[0]
            saldos[mid] = (k.saldo_final_cantidad, r2(k.saldo_final_total))
            salida[mes].extend(movs)
    return salida


# ---------------------------------------------------------------------------
# Construcción de cada Excel
# ---------------------------------------------------------------------------

def construir_mes(mes: int, indice_precio: float, ruido: random.Random) -> tuple[Path, list[dict]]:
    esc = ESCENARIOS[mes]
    wb = openpyxl.load_workbook(MAESTRO)
    per_id = f"PER2010-{mes:02d}"
    nombre_mes = f"{MESES[mes - 1]} 2010"
    real = mes == 7

    # 1) periodo_id en todas las hojas
    for ws in wb:
        hdr = [c.value for c in ws[1]]
        if "periodo_id" in hdr:
            j = hdr.index("periodo_id")
            for row in ws.iter_rows(min_row=2):
                if row[j].value == "PER001":
                    row[j].value = per_id
    hdr, rows = filas(wb["Periodo"])
    p = rows[0]
    p[col(hdr, "periodo")].value = nombre_mes
    p[col(hdr, "fecha_inicio")].value = f"2010-{mes:02d}-01"
    p[col(hdr, "fecha_fin")].value = f"2010-{mes:02d}-{ULTIMO_DIA[mes - 1]:02d}"
    p[col(hdr, "descripcion")].value = (f"Periodo mensual del caso Frenosa (datos reales)." if real else
                                        f"Periodo simulado (datos ficticios). Escenario: {esc['texto']}")
    if not real:
        p[col(hdr, "fuente")].value = FUENTE_SIM

    # 2) OT y objetos de costo
    hdr, ots = filas(wb["Ordenes_Trabajo"])
    base_q = {}
    for i, row in enumerate(ots):
        ot = row[col(hdr, "ot_id")].value
        q0 = float(row[col(hdr, "cantidad_producida")].value)
        base_q[ot] = q0
        q = q0 if real else float(round(q0 * esc["vol"][i]))
        row[col(hdr, "cantidad_producida")].value = q
        cod = row[col(hdr, "codigo_ot")].value
        row[col(hdr, "codigo_ot")].value = f"{cod}-2010-{mes:02d}"
        row[col(hdr, "descripcion_ot")].value = f"{row[col(hdr, 'descripcion_ot')].value} - {nombre_mes}"
        if not real:
            row[col(hdr, "fuente")].value = FUENTE_SIM
    q_mes = {row[col(hdr, "ot_id")].value: float(row[col(hdr, "cantidad_producida")].value) for row in ots}
    codigos = {row[col(hdr, "ot_id")].value: row[col(hdr, "codigo_ot")].value for row in ots}
    ratio = {ot: q_mes[ot] / base_q[ot] for ot in q_mes}
    vol_total = sum(q_mes.values()) / sum(base_q.values())

    hdr, objs = filas(wb["Objetos_Costo"])
    for row in objs:
        ot = row[col(hdr, "objeto_costo_id")].value
        row[col(hdr, "cantidad_piezas")].value = q_mes[ot]
        desc = row[col(hdr, "descripcion_producto")].value
        row[col(hdr, "descripcion_producto")].value = desc.replace(desc.split(" - ")[0], codigos[ot], 1)
        if not real:
            row[col(hdr, "fuente")].value = FUENTE_SIM

    # 3) materiales directos del mes (objetivo de CMD del kardex; no se escriben en una hoja)
    materiales = []
    for base in MATERIALES_BASE:
        mat = dict(base)
        mat["codigo_ot"] = codigos[mat["ot_id"]]
        if not real:
            mat["cantidad_kg"] = q_mes[mat["ot_id"]]
            mat["costo_unitario_soles"] = round(float(base["costo_unitario_soles"]) * esc["mat"] * ruido.uniform(0.995, 1.005), 4)
        materiales.append(mat)
    if not real:

        # 4) personal directo (MOD)
        hdr, pds = filas(wb["Personal_directos"])
        kg = sum(q_mes.values())
        te = []
        for row in pds:
            ts = sum(float(row[col(hdr, c)].value) for c in ("tolerancia_personal_pct", "tolerancia_fatiga_pct",
                                                             "tolerancia_pie_pct", "tolerancia_encorvado_pct"))
            te.append(rh(float(row[col(hdr, "tiempo_normal_min_kg")].value) * (1 + ts), 2))
        horas_std = rh(sum(te) * kg / 60)
        horas_reales = rh(horas_std * esc["efic"])
        fd = rh(horas_reales / horas_std, 2)
        tiempo_total = sum(rh(t * kg * fd) for t in te)
        for row in pds:
            row[col(hdr, "kg_producidos_periodo")].value = kg
            row[col(hdr, "horas_reales_trabajadas_periodo")].value = horas_reales
            row[col(hdr, "horas_estandar_totales_periodo")].value = horas_std
            row[col(hdr, "tiempo_real_total_minutos_periodo")].value = tiempo_total
            row[col(hdr, "cantidad_operarios")].value = esc["operarios"]
            row[col(hdr, "fuente")].value = FUENTE_SIM

        # 5) otros recursos indirectos (fijos + variables con el volumen)
        hdr, recs = filas(wb["costo_otros_recursos(CI)"])
        for row in recs:
            ec = row[col(hdr, "elemento_costo_id")].value
            base = float(row[col(hdr, "importe_soles")].value)
            var = VARIABILIDAD_RECURSOS.get(ec, 0.3)
            row[col(hdr, "importe_soles")].value = rh(base * ((1 - var) + var * vol_total) * ruido.uniform(0.98, 1.02))
            row[col(hdr, "fuente")].value = FUENTE_SIM

        # 6) soporte: mantenimiento, control de calidad, ingeniería
        hdr, mant = filas(wb["Soporte_Mantenimiento(CI)"])
        for row in mant:
            base = float(row[col(hdr, "costo_equipo_soles")].value)
            extra = esc["extra"].get(row[col(hdr, "equipo_codigo")].value, 0)
            row[col(hdr, "costo_equipo_soles")].value = rh(base * (0.6 + 0.4 * vol_total) * ruido.uniform(0.95, 1.05) + extra)
            row[col(hdr, "fuente")].value = FUENTE_SIM
            if extra:
                row[col(hdr, "observacion")].value = f"Incluye reparación extraordinaria de S/ {extra:,.0f} (escenario simulado)."
        hdr, cc = filas(wb["Soporte_Control_Calidad(CI)"])
        for row in cc:
            base = float(row[col(hdr, "cantidad")].value)
            row[col(hdr, "cantidad")].value = max(1.0, float(round(base * vol_total)))
            row[col(hdr, "fuente")].value = FUENTE_SIM
        hdr, ing = filas(wb["Soporte_Ingenieria(CI)"])
        for i, row in enumerate(ing):
            base = float(row[col(hdr, "cantidad")].value)
            factor = (0.5 + 0.5 * vol_total) if i == 0 else ruido.uniform(0.6, 1.4)
            row[col(hdr, "cantidad")].value = max(1.0, float(round(base * factor)))
            row[col(hdr, "fuente")].value = FUENTE_SIM

        # 7) pesos actividad -> OT según la mezcla de producción del mes
        for hoja in ("AC_OT_Pesos_Produccion", "AC_OT_Pesos_Soporte"):
            hdr, pesos = filas(wb[hoja])
            por_ac: dict[str, list] = {}
            for row in pesos:
                por_ac.setdefault(row[col(hdr, "actividad_id")].value, []).append(row)
            for grupo in por_ac.values():
                crudos = [float(r[col(hdr, "porcentaje_inductor")].value) * ratio[r[col(hdr, "ot_id")].value] for r in grupo]
                for r, nuevo in zip(grupo, mayor_residuo(crudos)):
                    r[col(hdr, "porcentaje_inductor")].value = nuevo
                    r[col(hdr, "fuente")].value = FUENTE_SIM
                    r[col(hdr, "observacion")].value = ("Peso del caso ajustado a la mezcla de producción del mes "
                                                        "(peso de julio × variación del volumen de la OT, normalizado).")

        # 8) ventas: cantidad vendida = producida; valor de venta unitario según el escenario
        hdr, ventas = filas(wb["Ventas"])
        for row in ventas:
            ot = row[col(hdr, "ot_id")].value
            q = q_mes[ot]
            vvu_julio = r2(float(row[col(hdr, "ventas_netas_soles")].value) / base_q[ot])
            vvu = r2(vvu_julio * indice_precio * ruido.uniform(0.99, 1.01))
            row[col(hdr, "cantidad_vendida")].value = q
            row[col(hdr, "ventas_netas_soles")].value = r2(q * vvu)
            row[col(hdr, "fuente")].value = FUENTE_SIM
            row[col(hdr, "observacion")].value = (
                f"Simulado: cantidad vendida = cantidad producida; valor de venta unitario S/ {vvu:,.2f} "
                f"(índice de precio del mes {indice_precio:.3f} respecto de julio).")

    # 10) README
    rd = wb["README"]
    for row in rd.iter_rows(min_row=2):
        if row[0].value == "Archivo":
            row[1].value = f"frenosa_datos_entrada_abc_2010_{mes:02d}.xlsx"
    fila = next(r[0].row for r in rd.iter_rows(min_row=2) if r[0].value == "Objeto de costo") + 1
    rd.insert_rows(fila, 2)
    rd.cell(fila, 1, "Periodo del archivo")
    rd.cell(fila, 2, f"{nombre_mes} ({per_id}). " + ("Datos reales del caso; mismos resultados que el Excel maestro."
                                                     if real else "Datos simulados (ficticios) para análisis mensual."))
    rd.cell(fila + 1, 1, "Escenario del mes")
    rd.cell(fila + 1, 2, esc["texto"])
    for f in (fila, fila + 1):
        for c in (1, 2):
            rd.cell(f, c)._style = copy(rd.cell(fila - 1, c)._style)

    SALIDA.mkdir(parents=True, exist_ok=True)
    destino = SALIDA / f"frenosa_datos_entrada_abc_2010_{mes:02d}.xlsx"
    wb.save(destino)
    return destino, materiales


def quitar_hoja_materiales(wb) -> None:
    """Elimina la hoja Materiales_Directos y deja los textos del libro con el kardex como fuente del CMD."""
    if "Materiales_Directos" in wb.sheetnames:
        del wb["Materiales_Directos"]
    dic = wb["Diccionario"]
    for r in list(dic.iter_rows(min_row=2)):
        if r[0].value == "Materiales_Directos":
            dic.delete_rows(r[0].row)
            break
    for r in dic.iter_rows(min_row=2):
        if r[0].value == "Compras":
            r[1].value = "Entrada (fuente del CMD)"
            r[2].value = ("Kardex de materiales directos (SW-50, FRE-9000, FRE-9002) del almacén de materias primas: saldo inicial, "
                          "compras con IGV y flete (flete_soles), devoluciones a proveedor y salidas a cada OT (destino = codigo_ot). "
                          "El sistema lo valoriza por promedio ponderado; las salidas valorizadas a cada OT son su CMD.")
            r[3].value = "No incluye costo promedio, salidas valorizadas ni saldos: los calcula el sistema."
    fu = wb["Fuentes"]
    for r in fu.iter_rows(min_row=2):
        if r[0].value == "F001":
            r[3].value = ("Materiales directos: cantidades y costos unitarios de julio 2010; el kardex de la hoja Compras "
                          "reproduce ese CMD con sus salidas a cada OT.")
        if r[0].value == "F008":
            r[3].value = ("Kardex de materiales directos por promedio ponderado (fuente del CMD por OT); "
                          "cantidad vendida y valor de venta unitario por OT.")
    rc = wb["Reglas_Calculo"]
    for r in rc.iter_rows(min_row=2):
        if r[0].value == "R001":
            r[2].value = "Salidas del kardex (hoja Compras, tipo Consumo) a la OT × costo promedio ponderado vigente"
            r[4].value = ("El material directo se asigna a la OT de destino de cada salida del kardex; el sistema valoriza "
                          "el kardex por promedio ponderado y suma las salidas de cada OT.")
    rd = wb["README"]
    for r in rd.iter_rows(min_row=2):
        if r[0].value == "Datos ficticios":
            r[1].value = ("La hoja Compras (kardex de materiales directos SW-50, FRE-9000 y FRE-9002; IGV 19 %) y las columnas "
                          "cantidad_vendida y valor_venta_unitario_soles de Ventas son datos simulados. El CMD de cada OT sale de "
                          "las salidas del kardex; en julio reproducen las cantidades y costos unitarios del Cuadro 1 del caso.")


def leer_materiales(ruta: Path) -> list[dict]:
    wb = openpyxl.load_workbook(ruta, data_only=True)
    hdr, mats = filas(wb["Materiales_Directos"])
    _, ots = filas(wb["Ordenes_Trabajo"])
    codigos = {row[0].value: row[5].value for row in ots}
    salida = []
    for row in mats:
        d = {h: c.value for h, c in zip(hdr, row)}
        d["codigo_ot"] = codigos.get(d["ot_id"], d["ot_id"])
        salida.append(d)
    return salida


def escribir_compras(ruta: Path, movimientos: list[dict], periodo_id: str, prefijo: str,
                     quitar_sufijo: str | None = None) -> None:
    """Reemplaza la hoja Compras del archivo con los movimientos del kardex de materiales."""
    wb = openpyxl.load_workbook(ruta)
    ws = wb["Compras"]
    hdr = [c.value for c in ws[1]]
    if "flete_soles" not in hdr:
        pos = hdr.index("igv_pct") + 2
        ws.insert_cols(pos)
        ws.cell(1, pos, "flete_soles")._style = copy(ws.cell(1, pos - 1)._style)
        ws.column_dimensions[openpyxl.utils.get_column_letter(pos)].width = 12
        hdr = [c.value for c in ws[1]]
    estilo = [copy(c._style) for c in ws[2]]
    formatos = [c.number_format for c in ws[2]]
    ws.delete_rows(2, ws.max_row)
    for n, mov in enumerate(movimientos, start=1):
        v = dict(mov, compra_id=f"{prefijo}{n:03d}", periodo_id=periodo_id, fuente=FUENTE_KARDEX)
        if quitar_sufijo:
            for campo in ("destino", "detalle"):
                v[campo] = (v.get(campo) or "").replace(quitar_sufijo, "")
        ws.append([v.get(h) for h in hdr])
        for c, st, fmt in zip(ws[ws.max_row], estilo, formatos):
            c._style = copy(st)
            c.number_format = fmt
    for row in ws.iter_rows(min_row=2):
        row[hdr.index("cantidad")].number_format = "#,##0.00"
        row[hdr.index("precio_unitario_con_igv")].number_format = "#,##0.0000"
        row[hdr.index("flete_soles")].number_format = "#,##0.00"
    _textos_kardex(wb)
    quitar_hoja_materiales(wb)
    wb.save(ruta)


def _textos_kardex(wb) -> None:
    dic = wb["Diccionario"]
    for r in dic.iter_rows(min_row=2):
        if r[0].value == "Compras":
            r[2].value = ("Kardex de materiales directos (SW-50, FRE-9000, FRE-9002) del almacén de materias primas: saldo inicial, "
                          "compras con IGV y flete (flete_soles), devoluciones a proveedor y salidas a cada OT. El sistema lo valoriza por promedio "
                          "ponderado; las salidas a cada OT suman el CMD de Materiales_Directos.")
            r[3].value = "No incluye costo promedio ni saldos; no cambia el cálculo ABC (el CMD sigue saliendo de Materiales_Directos)."
    fu = wb["Fuentes"]
    for r in fu.iter_rows(min_row=2):
        if r[0].value == "F008":
            r[3].value = ("Kardex de materiales directos por promedio ponderado (salidas = CMD por OT); "
                          "cantidad vendida y valor de venta unitario por OT.")
    rd = wb["README"]
    for r in rd.iter_rows(min_row=2):
        if r[0].value == "Datos ficticios":
            r[1].value = ("La hoja Compras (kardex de materiales directos SW-50, FRE-9000 y FRE-9002; IGV 19 %) y las columnas "
                          "cantidad_vendida y valor_venta_unitario_soles de Ventas son datos simulados. Las salidas del kardex a "
                          "cada OT coinciden con el CMD, por eso no cambian los cálculos ni resultados del caso.")


FUENTE_KARDEX = ("Kardex simulado (datos ficticios) para fines académicos: compras y saldos inventados; las salidas a cada OT "
                 "son el CMD (en julio reproducen el Cuadro 1 del caso).")


def generar(indices: dict[int, float], actualizar_maestro: bool = True) -> list[Path]:
    rutas, materiales = [], {}
    for mes in range(1, 13):
        ruido = random.Random(1000 + mes)
        ruta, mats = construir_mes(mes, indices.get(mes, 1.0), ruido)
        rutas.append(ruta)
        materiales[mes] = mats
    kardex = simular_kardex_materiales(materiales)
    objetivos = {f"2010-{mes:02d}": [{k: m[k] for k in ("material_id", "material", "ot_id", "codigo_ot", "cantidad_kg",
                                                         "costo_unitario_soles")} for m in mats] for mes, mats in materiales.items()}
    (ROOT / "tools" / "cmd_objetivo_2010.json").write_text(json.dumps(objetivos, ensure_ascii=False, indent=1), encoding="utf-8")
    for mes, ruta in enumerate(rutas, start=1):
        escribir_compras(ruta, kardex[mes], f"PER2010-{mes:02d}", f"CMP{mes:02d}")
    if actualizar_maestro:
        escribir_compras(MAESTRO, kardex[7], "PER001", "CMP", quitar_sufijo="-2010-07")
    return rutas


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--indices", help="JSON {mes: indice_de_precio}")
    args = parser.parse_args()
    indices = {int(k): v for k, v in json.loads(Path(args.indices).read_text()).items()} if args.indices else {}
    for ruta in generar(indices):
        print(ruta)
    print(MAESTRO, "(hoja Compras de julio)")
