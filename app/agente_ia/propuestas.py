"""Propuestas de mejora del agente de Finanzas y Contabilidad.

Las propuestas se calculan leyendo los resultados del ultimo calculo. El agente
NO aplica ningun cambio: cada propuesta indica el area responsable y el dato
que habria que modificar en el Excel maestro si la empresa decide ejecutarla.
Los impactos son estimaciones ceteris paribus (todo lo demas constante).
"""

from __future__ import annotations

from dataclasses import dataclass

from .formato import num, pct, soles
from .kardex import kardex_desde_base, salidas_por_material


@dataclass
class Propuesta:
    prioridad: str
    titulo: str
    diagnostico: str
    accion: str
    impacto: str
    area: str
    impacto_soles: float = 0.0


def _etiquetas(lector) -> dict[str, str]:
    objetos = {r["objeto_costo_id"]: r for r in lector.tabla("objetos_costo")}
    salida = {}
    for ot in lector.tabla("ordenes_trabajo"):
        prod = objetos.get(ot.get("objeto_costo_id"), {}).get("codigo_producto") or ot.get("objeto_costo_id")
        salida[ot["ot_id"]] = f"{ot.get('codigo_ot') or ot['ot_id']} ({prod})"
    return salida


def priorizar_ots(lector) -> list[dict]:
    hoja = {f["ot_id"]: f for f in lector.calculado("hoja_costos_ot")}
    estado = lector.calculado("estado_resultados_ot")
    if not estado:
        return []
    promedio = sum(num(e.get("rentabilidad_operativa")) for e in estado) / len(estado)
    filas = []
    for e in estado:
        h = hoja.get(e["ot_id"], {})
        ventas = num(e.get("ventas_netas_soles"))
        costo = num(h.get("costo_total_ot"))
        rent = num(e.get("rentabilidad_operativa"))
        cmd_v = num(h.get("cmd_ot")) / ventas if ventas else 0
        sop_c = num(h.get("costo_soporte_ot")) / costo if costo else 0
        motivos = []
        if rent < promedio:
            motivos.append(f"rentabilidad operativa {pct(rent)} (promedio simple de las OT {pct(promedio)})")
        if cmd_v >= 0.6:
            motivos.append(f"el CMD absorbe {pct(cmd_v)} de sus ventas")
        if sop_c >= 0.04:
            motivos.append(f"soporte = {pct(sop_c)} de su costo, el mayor peso relativo")
        filas.append({"ot_id": e["ot_id"], "rentabilidad": rent, "cmd_sobre_ventas": cmd_v, "soporte_sobre_costo": sop_c,
                      "motivos": motivos})
    max_sop = max(f["soporte_sobre_costo"] for f in filas)
    for f in filas:
        if f["soporte_sobre_costo"] != max_sop:
            f["motivos"] = [m for m in f["motivos"] if not m.startswith("soporte")]
    filas.sort(key=lambda f: (f["rentabilidad"], -f["cmd_sobre_ventas"], -f["soporte_sobre_costo"]))
    return filas


def generar_propuestas(lector) -> list[Propuesta]:
    hoja = lector.calculado("hoja_costos_ot")
    estado = lector.calculado("estado_resultados_ot")
    actividades = lector.calculado("actividad_total")
    if not hoja or not estado:
        return []
    et = _etiquetas(lector)
    props: list[Propuesta] = []
    ventas_tot = sum(num(e.get("ventas_netas_soles")) for e in estado)
    mo_tot = sum(num(e.get("margen_operativo")) for e in estado)
    rent_glob = mo_tot / ventas_tot if ventas_tot else 0
    hoja_por_ot = {h["ot_id"]: h for h in hoja}

    # 1. OT con menor rentabilidad: material y precio
    bajo = min(estado, key=lambda e: num(e.get("rentabilidad_operativa")))
    alto = max(estado, key=lambda e: num(e.get("rentabilidad_operativa")))
    hb = hoja_por_ot.get(bajo["ot_id"], {})
    v_b = num(bajo.get("ventas_netas_soles"))
    cmd_b = num(hb.get("cmd_ot"))
    ahorro_cmd = 0.05 * cmd_b
    alza_precio = 0.03 * v_b
    nueva_rent = (num(bajo.get("margen_operativo")) + ahorro_cmd) / v_b if v_b else 0
    props.append(Propuesta(
        "Alta", f"Recuperar la rentabilidad de {et.get(bajo['ot_id'])}",
        f"Tiene la menor rentabilidad operativa ({pct(bajo.get('rentabilidad_operativa'))}); su CMD equivale a "
        f"{pct(cmd_b / v_b if v_b else 0)} de sus ventas netas.",
        "Negociar el costo del material con proveedores o evaluar un sustituto homologado por Ingeniería; en paralelo, "
        "revisar la lista de precios de este producto con Comercial.",
        f"−5 % en el costo del material: +{soles(ahorro_cmd)} de margen (rentabilidad {pct(bajo.get('rentabilidad_operativa'))} → "
        f"{pct(nueva_rent)}). +3 % en el precio: +{soles(alza_precio)}.",
        "Logística / Comercial", ahorro_cmd))

    # 2. Material directo (kardex) como componente dominante
    cmd_tot = sum(num(h.get("cmd_ot")) for h in hoja)
    costo_tot = sum(num(h.get("costo_total_ot")) for h in hoja)
    kardexs, _obs = kardex_desde_base(lector)
    if kardexs:
        mayor = max(kardexs, key=lambda k: k.consumo)
        ahorro = 0.03 * mayor.consumo
        q_may, t_may, _dest = salidas_por_material(mayor)
        un_ingreso = all(sum(1 for f in k.filas if f.tipo == "ENTRADA") <= 1 for k in kardexs)
        props.append(Propuesta(
            "Alta", "Gestionar el costo del material directo con el kardex",
            f"El CMD es {pct(cmd_tot / costo_tot if costo_tot else 0)} del costo total de las OT; el material de mayor consumo "
            f"valorizado es {mayor.material} ({soles(mayor.consumo)} a S/ {(t_may / q_may if q_may else 0):,.4f} promedio).",
            "Consolidar compras y negociar precio/flete del material de mayor consumo"
            + ("; registrar en la hoja Compras cada compra y consumo real para que el promedio ponderado refleje los precios "
               "del periodo (hoy hay un solo ingreso por material)." if un_ingreso else "."),
            f"−3 % en el costo promedio de {mayor.material}: +{soles(ahorro)} de margen operativo.",
            "Logística / Contabilidad de costos", ahorro))

    # 3. Actividad de mayor costo
    if actividades:
        total_ac = sum(num(a.get("costo_total_ac")) for a in actividades)
        top = max(actividades, key=lambda a: num(a.get("costo_total_ac")))
        t = num(top.get("costo_total_ac"))
        comps = {"MOD": num(top.get("costo_mod_ac")), "otros recursos": num(top.get("costo_otros_recursos_ac")),
                 "soporte": num(top.get("costo_soporte_ac"))}
        dom = max(comps, key=comps.get)
        props.append(Propuesta(
            "Media", f"Eficiencia en la actividad {top.get('nombre_actividad')}",
            f"Es la actividad más costosa: {soles(t)} ({pct(t / total_ac if total_ac else 0)} del costo de actividades); "
            f"su principal recurso es {dom} ({pct(comps[dom] / t if t else 0)}).",
            "Revisar el método de trabajo, la carga de máquina y los inductores de esta actividad con Producción.",
            f"−10 % de costo en la actividad: +{soles(0.10 * t)} de margen, repartido entre las OT según sus pesos AC-OT.",
            "Producción", 0.10 * t))

        # 4. Soporte intensivo
        def ratio(a):
            return num(a.get("costo_soporte_ac")) / num(a.get("costo_total_ac")) if num(a.get("costo_total_ac")) else 0

        sop = max(actividades, key=ratio)
        mant = {m["actividad_id"]: num(m.get("costo_mantenimiento_ac")) for m in lector.calculado("soporte_mantenimiento_ac")}
        equipos = [e for e in lector.tabla("Soporte_Mantenimiento(CI)") if e.get("actividad_id_destino") == sop.get("actividad_id")]
        equipo = max(equipos, key=lambda e: num(e.get("costo_equipo_soles"))) if equipos else None
        m = mant.get(sop.get("actividad_id"), 0)
        props.append(Propuesta(
            "Media", f"Mantenimiento preventivo en {sop.get('nombre_actividad')}",
            f"El soporte representa {pct(ratio(sop))} del costo de {sop.get('nombre_actividad')}; el mantenimiento suma {soles(m)}"
            + (f", sobre todo {equipo.get('equipo_descripcion')} ({soles(equipo.get('costo_equipo_soles'))})." if equipo else "."),
            "Programar mantenimiento preventivo y llevar control de fallas del equipo para reducir correctivos.",
            f"−15 % del mantenimiento de la actividad: +{soles(0.15 * m)} de margen.",
            "Mantenimiento", 0.15 * m))

    # 5. Mezcla de ventas
    if bajo["ot_id"] != alto["ot_id"]:
        dif = num(alto.get("rentabilidad_operativa")) - num(bajo.get("rentabilidad_operativa"))
        props.append(Propuesta(
            "Media", "Orientar la mezcla de ventas hacia el producto más rentable",
            f"{et.get(alto['ot_id'])} rinde {pct(alto.get('rentabilidad_operativa'))} y {et.get(bajo['ot_id'])} "
            f"{pct(bajo.get('rentabilidad_operativa'))} (global {pct(rent_glob)}).",
            "Priorizar en el plan comercial y de producción el producto de mayor rentabilidad, sin desatender la demanda del resto.",
            f"Cada S/ 1,000.00 de ventas que se desplacen hacia {et.get(alto['ot_id'])} aportan {soles(1000 * dif)} más de margen "
            "(supone márgenes constantes y capacidad disponible).",
            "Comercial / Planeamiento", 0.0))

    # 6. MOD y factor de distribucion
    mod = lector.calculado("mod_actividad")
    if mod:
        fd = num(mod[0].get("fd"))
        if fd < 1:
            diag = f"El Fd es {fd:,.2f}: las horas reales fueron {pct(1 - fd, 0)} menores que las estándar."
            acc = "Actualizar el estudio de tiempos: un estándar holgado distorsiona el reparto de la MOD entre actividades."
        elif fd > 1:
            diag = f"El Fd es {fd:,.2f}: las horas reales superaron en {pct(fd - 1, 0)} a las estándar."
            acc = "Analizar demoras y capacitar en las actividades con mayor tiempo real."
        else:
            diag, acc = "El Fd es 1.00: tiempos reales y estándar coinciden.", "Mantener el control de tiempos."
        props.append(Propuesta("Baja", "Precisión de los tiempos estándar de MOD", diag, acc,
                               "Mejora la exactitud del costo por actividad; no cambia el total de MOD del periodo.", "Ingeniería / Producción"))

    # 7. Calidad de datos
    datos = []
    if not lector.tabla("servicios_directos"):
        datos.append("CSD = 0 por falta de servicios directos en el Excel")
    if datos:
        props.append(Propuesta("Baja", "Completar la información de costeo",
                               "; ".join(datos) + ".",
                               "Si existen servicios de terceros identificables con una OT, registrarlos en la hoja Servicios_Directos.",
                               "Costeo más completo y trazable; no altera las cifras actuales.", "Contabilidad de costos"))
    return props


def propuestas_como_texto(props: list[Propuesta]) -> str:
    if not props:
        return "Sin resultados calculados para proponer mejoras."
    lineas = []
    for n, p in enumerate(props, start=1):
        lineas += [f"{n}. [{p.prioridad}] {p.titulo}",
                   f"   Diagnóstico: {p.diagnostico}",
                   f"   Propuesta: {p.accion}",
                   f"   Impacto estimado: {p.impacto}",
                   f"   Responsable sugerido: {p.area}", ""]
    total = sum(p.impacto_soles for p in props)
    if total:
        lineas.append(f"Impacto potencial sumado de las propuestas cuantificadas: {soles(total)} de margen operativo "
                      "(estimación; las propuestas pueden solaparse).")
    lineas.append("El agente no aplica cambios: cada propuesta debe aprobarse y registrarse en el Excel maestro por el área responsable.")
    return "\n".join(lineas)
