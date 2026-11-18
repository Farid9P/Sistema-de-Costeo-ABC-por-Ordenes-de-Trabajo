"""Kardex de inventario valorizado por PROMEDIO PONDERADO (hoja Compras).

Movimientos admitidos (columna tipo_movimiento de la hoja Compras):

- Saldo inicial ........ cantidad + importe_soles (valor total del saldo).
- Compra ............... ingreso; costo unitario = (cantidad x precio sin IGV + flete) / cantidad,
                         con precio sin IGV = precio con IGV / (1 + IGV).
- Devolucion a proveedor ingreso NEGATIVO al costo unitario de la compra devuelta.
- Consumo .............. salida al costo promedio vigente.
- Devolucion de consumo  salida NEGATIVA (reingreso) al costo promedio vigente.

Despues de cada movimiento: saldo = saldo anterior + ingresos - salidas y
costo promedio = valor del saldo / cantidad del saldo.
Comprobacion: saldo inicial + compras netas - consumos netos = saldo final.

El kardex registra el almacen de materias primas (materiales directos SW-50,
FRE-9000 y FRE-9002) y es la FUENTE DEL CMD: las salidas (consumos) valorizadas
al costo promedio vigente se asignan a la OT de su columna destino (codigo_ot u
ot_id); ver salidas_por_ot() y costing_direct.calculate_cmd().
"""

from __future__ import annotations

import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime

IGV_POR_DEFECTO = 0.18
TOLERANCIA = 1e-9

SALDO_INICIAL = "Saldo inicial"
COMPRA = "Compra"
DEVOLUCION_PROVEEDOR = "Devolucion a proveedor"
CONSUMO = "Consumo"
DEVOLUCION_CONSUMO = "Devolucion de consumo"
TIPOS = [SALDO_INICIAL, COMPRA, DEVOLUCION_PROVEEDOR, CONSUMO, DEVOLUCION_CONSUMO]

ETIQUETA_TIPO = {
    SALDO_INICIAL: "Saldo inicial",
    COMPRA: "Compras",
    DEVOLUCION_PROVEEDOR: "Devolución",
    CONSUMO: "Consumo",
    DEVOLUCION_CONSUMO: "Devolución",
}


def _norm(texto) -> str:
    texto = unicodedata.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode().lower().strip()
    return " ".join(texto.split())


_ALIAS_TIPO = {
    "saldo inicial": SALDO_INICIAL, "inventario inicial": SALDO_INICIAL,
    "compra": COMPRA, "compras": COMPRA, "ingreso": COMPRA,
    "devolucion a proveedor": DEVOLUCION_PROVEEDOR, "devolucion al proveedor": DEVOLUCION_PROVEEDOR,
    "consumo": CONSUMO, "consumos": CONSUMO, "salida": CONSUMO,
    "devolucion de consumo": DEVOLUCION_CONSUMO, "devolucion de mantenimiento": DEVOLUCION_CONSUMO,
    "reingreso": DEVOLUCION_CONSUMO,
}


def tipo_movimiento(valor) -> str | None:
    return _ALIAS_TIPO.get(_norm(valor))


def _num(valor) -> float:
    try:
        return float(valor) if valor not in (None, "") else 0.0
    except (TypeError, ValueError):
        return 0.0


def _fecha(valor) -> str:
    if isinstance(valor, datetime):
        return valor.date().isoformat()
    if isinstance(valor, date):
        return valor.isoformat()
    return str(valor or "")[:10]


@dataclass
class FilaKardex:
    compra_id: str
    fecha: str
    tipo: str
    descripcion: str
    detalle: str
    documento: str
    destino: str
    ing_cantidad: float | None = None
    ing_costo_unitario: float | None = None
    ing_total: float | None = None
    sal_cantidad: float | None = None
    sal_costo_unitario: float | None = None
    sal_total: float | None = None
    saldo_cantidad: float = 0.0
    saldo_costo_unitario: float = 0.0
    saldo_total: float = 0.0
    precio_con_igv: float | None = None
    igv_pct: float | None = None
    flete: float | None = None


@dataclass
class KardexItem:
    item_codigo: str
    item_descripcion: str = ""
    almacen: str = ""
    unidad: str = ""
    periodo_id: str = ""
    filas: list[FilaKardex] = field(default_factory=list)
    avisos: list[str] = field(default_factory=list)
    saldo_inicial_cantidad: float = 0.0
    saldo_inicial_total: float = 0.0
    _cantidad: float = 0.0
    _total: float = 0.0
    _compras_cu: dict = field(default_factory=dict)

    # -- saldos --------------------------------------------------------------
    @property
    def costo_promedio(self) -> float:
        return self._total / self._cantidad if abs(self._cantidad) > TOLERANCIA else 0.0

    def _cerrar(self, fila: FilaKardex) -> FilaKardex:
        if abs(self._cantidad) < TOLERANCIA:
            self._cantidad = 0.0
            self._total = 0.0
        fila.saldo_cantidad = self._cantidad
        fila.saldo_total = self._total
        fila.saldo_costo_unitario = self.costo_promedio if self._cantidad else (fila.sal_costo_unitario or 0.0)
        self.filas.append(fila)
        return fila

    # -- movimientos ---------------------------------------------------------
    def registrar(self, mov: dict) -> FilaKardex | None:
        tipo = tipo_movimiento(mov.get("tipo_movimiento"))
        base = dict(
            compra_id=str(mov.get("compra_id") or ""), fecha=_fecha(mov.get("fecha")), tipo=tipo or "",
            descripcion=ETIQUETA_TIPO.get(tipo, str(mov.get("tipo_movimiento") or "")),
            detalle=str(mov.get("detalle") or ""), documento=str(mov.get("documento") or ""),
            destino=str(mov.get("destino") or ""),
        )
        cantidad = abs(_num(mov.get("cantidad")))
        if tipo is None:
            self.avisos.append(f"{base['compra_id']}: tipo de movimiento no reconocido ({mov.get('tipo_movimiento')}).")
            return None
        if cantidad <= 0:
            self.avisos.append(f"{base['compra_id']}: la cantidad debe ser mayor que cero.")
            return None
        fila = FilaKardex(**base)

        if tipo == SALDO_INICIAL:
            if self.filas:
                self.avisos.append(f"{fila.compra_id}: el saldo inicial debe ser el primer movimiento del ítem.")
            total = _num(mov.get("importe_soles"))
            if not total and _num(mov.get("costo_unitario_soles")):
                total = cantidad * _num(mov.get("costo_unitario_soles"))
            self._cantidad, self._total = cantidad, total
            self.saldo_inicial_cantidad, self.saldo_inicial_total = cantidad, total
            return self._cerrar(fila)

        if tipo in (COMPRA, DEVOLUCION_PROVEEDOR):
            igv = _num(mov.get("igv_pct")) if mov.get("igv_pct") not in (None, "") else IGV_POR_DEFECTO
            precio = _num(mov.get("precio_unitario_con_igv"))
            cu = precio / (1 + igv) if precio else _num(mov.get("costo_unitario_soles"))
            fila.precio_con_igv, fila.igv_pct = precio or None, igv
            signo = 1 if tipo == COMPRA else -1
            if tipo == COMPRA:
                flete = _num(mov.get("flete_soles"))
                if flete:
                    fila.flete = flete
                    cu = (cantidad * cu + flete) / cantidad
                if fila.documento:
                    self._compras_cu[fila.documento] = cu
            if tipo == DEVOLUCION_PROVEEDOR:
                original = self._compras_cu.get(fila.documento)
                if original is None:
                    self.avisos.append(f"{fila.compra_id}: la devolución no encuentra una compra previa con el documento {fila.documento}.")
                else:
                    cu = original  # al costo unitario de la compra devuelta (incluye su flete)
                if cantidad - self._cantidad > TOLERANCIA:
                    self.avisos.append(f"{fila.compra_id}: se devuelve más de lo que hay en almacén.")
                    return None
            fila.ing_cantidad = signo * cantidad
            fila.ing_costo_unitario = cu
            fila.ing_total = signo * cantidad * cu
            self._cantidad += fila.ing_cantidad
            self._total += fila.ing_total
            return self._cerrar(fila)

        # CONSUMO o DEVOLUCION_CONSUMO: al costo promedio vigente
        promedio = self.costo_promedio
        if tipo == CONSUMO and cantidad - self._cantidad > TOLERANCIA:
            self.avisos.append(
                f"{fila.compra_id}: stock insuficiente (saldo {self._cantidad:,.2f}, consumo {cantidad:,.2f}); movimiento omitido."
            )
            return None
        signo = 1 if tipo == CONSUMO else -1
        fila.sal_cantidad = signo * cantidad
        fila.sal_costo_unitario = promedio
        fila.sal_total = signo * cantidad * promedio
        self._cantidad -= fila.sal_cantidad
        self._total -= fila.sal_total
        return self._cerrar(fila)

    # -- totales -------------------------------------------------------------
    def _suma(self, campo: str) -> float:
        return sum(getattr(f, campo) or 0.0 for f in self.filas)

    @property
    def ingresos_cantidad(self) -> float:
        return self._suma("ing_cantidad")

    @property
    def ingresos_total(self) -> float:
        return self._suma("ing_total")

    @property
    def salidas_cantidad(self) -> float:
        return self._suma("sal_cantidad")

    @property
    def salidas_total(self) -> float:
        return self._suma("sal_total")

    @property
    def saldo_final_cantidad(self) -> float:
        return self._cantidad

    @property
    def saldo_final_total(self) -> float:
        return self._total

    @property
    def saldo_final_costo_unitario(self) -> float:
        return self.costo_promedio

    def comprobacion(self) -> list[tuple[str, float, float]]:
        disponible_q = self.saldo_inicial_cantidad + self.ingresos_cantidad
        disponible_t = self.saldo_inicial_total + self.ingresos_total
        return [
            ("Saldo inicial", self.saldo_inicial_cantidad, self.saldo_inicial_total),
            ("Compras", self.ingresos_cantidad, self.ingresos_total),
            ("Disponible", disponible_q, disponible_t),
            ("Consumos", self.salidas_cantidad, self.salidas_total),
            ("Saldo final", disponible_q - self.salidas_cantidad, disponible_t - self.salidas_total),
        ]

    @property
    def cuadra(self) -> bool:
        _, q, t = self.comprobacion()[-1]
        return abs(q - self.saldo_final_cantidad) < 1e-6 and abs(t - self.saldo_final_total) < 0.005


def valorizar(movimientos: list[dict]) -> list[KardexItem]:
    """Agrupa por ítem (en el orden de aparición) y valoriza por fecha y compra_id."""
    items: dict[str, KardexItem] = {}
    por_item: dict[str, list[dict]] = {}
    for mov in movimientos:
        codigo = str(mov.get("item_codigo") or "SIN-CODIGO")
        if codigo not in items:
            items[codigo] = KardexItem(codigo, str(mov.get("item_descripcion") or ""), str(mov.get("almacen") or ""),
                                       str(mov.get("unidad") or ""), str(mov.get("periodo_id") or ""))
            por_item[codigo] = []
        por_item[codigo].append(mov)
    for codigo, movs in por_item.items():
        orden_tipo = {SALDO_INICIAL: 0}
        movs = sorted(enumerate(movs), key=lambda par: (_fecha(par[1].get("fecha")),
                                                         orden_tipo.get(tipo_movimiento(par[1].get("tipo_movimiento")), 1),
                                                         par[0]))
        for _, mov in movs:
            items[codigo].registrar(mov)
    return list(items.values())


def kardex_desde_base(conn) -> list[KardexItem]:
    """Lee la tabla compras (si existe) y devuelve el kardex por ítem. Solo lectura."""
    try:
        existe = conn.execute("SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'compras'").fetchone()
    except Exception:
        return []
    if not existe:
        return []
    cursor = conn.execute("SELECT * FROM compras ORDER BY rowid")
    columnas = [c[0] for c in cursor.description]
    return valorizar([dict(zip(columnas, fila)) for fila in cursor.fetchall()])


def filas_para_tabla(item: KardexItem) -> list[dict]:
    """Filas listas para mostrar (pestaña Datos importados y agente)."""

    def f(valor, dec=2):
        return "" if valor is None else f"{valor:,.{dec}f}"

    salida = []
    for fila in item.filas:
        salida.append({
            "descripcion": fila.descripcion,
            "fecha": fila.fecha,
            "detalle": fila.detalle,
            "ing_cantidad": f(fila.ing_cantidad),
            "ing_cto_unit": f(fila.ing_costo_unitario, 4),
            "ing_cto_total": f(fila.ing_total),
            "sal_cantidad": f(fila.sal_cantidad),
            "sal_cto_unit": f(fila.sal_costo_unitario, 4),
            "sal_cto_total": f(fila.sal_total),
            "saldo_cantidad": f(fila.saldo_cantidad),
            "saldo_cto_unit": f(fila.saldo_costo_unitario, 4),
            "saldo_cto_total": f(fila.saldo_total),
            "documento": fila.documento,
            "destino": fila.destino,
        })
    salida.append({
        "descripcion": "TOTALES", "fecha": "", "detalle": "Ingresos y salidas del periodo",
        "ing_cantidad": f(item.ingresos_cantidad), "ing_cto_unit": "", "ing_cto_total": f(item.ingresos_total),
        "sal_cantidad": f(item.salidas_cantidad), "sal_cto_unit": "", "sal_cto_total": f(item.salidas_total),
        "saldo_cantidad": f(item.saldo_final_cantidad), "saldo_cto_unit": f(item.saldo_final_costo_unitario, 4),
        "saldo_cto_total": f(item.saldo_final_total), "documento": "", "destino": "",
    })
    return salida


def mapa_ot(conn) -> dict[str, str]:
    """codigo_ot y ot_id (normalizados) -> ot_id, para leer el destino de las salidas."""
    mapa: dict[str, str] = {}
    try:
        for fila in conn.execute("SELECT ot_id, codigo_ot FROM ordenes_trabajo"):
            ot_id = fila[0]
            for clave in (fila[0], fila[1]):
                if clave:
                    mapa[_norm(clave)] = ot_id
    except Exception:
        pass
    return mapa


def salidas_por_ot(items: list[KardexItem], ots: dict[str, str]) -> tuple[dict[tuple[str, str], dict], list[str]]:
    """Suma las salidas valorizadas (consumos menos devoluciones de consumo) por (ot_id, material).

    Devuelve el acumulado y la lista de avisos de salidas cuyo destino no es una OT del Excel.
    """
    acumulado: dict[tuple[str, str], dict] = {}
    avisos: list[str] = []
    for item in items:
        for fila in item.filas:
            if fila.tipo not in (CONSUMO, DEVOLUCION_CONSUMO) or not fila.sal_cantidad:
                continue
            ot_id = ots.get(_norm(fila.destino))
            if ot_id is None:
                avisos.append(f"{fila.compra_id}: la salida de {item.item_codigo} tiene destino '{fila.destino}', "
                              "que no es una OT de Ordenes_Trabajo; no se asigna al CMD.")
                continue
            clave = (ot_id, item.item_codigo)
            dato = acumulado.setdefault(clave, {"ot_id": ot_id, "material_id": item.item_codigo,
                                                "material": item.item_descripcion, "cantidad": 0.0, "total": 0.0,
                                                "documentos": []})
            dato["cantidad"] += fila.sal_cantidad
            dato["total"] += fila.sal_total or 0.0
            if fila.documento:
                dato["documentos"].append(fila.documento)
    return acumulado, avisos


def conciliacion_cmd(conn, items: list[KardexItem] | None = None) -> list[dict]:
    """Salidas del kardex de cada material frente al CMD que calculo el sistema (tabla cmd_ot de la
    ultima ejecucion). El CMD sale del kardex, asi que deben coincidir. Solo lectura."""
    items = items if items is not None else kardex_desde_base(conn)
    try:
        run = conn.execute(
            "SELECT run_id FROM calculation_runs WHERE status = 'COMPLETADO' ORDER BY run_id DESC LIMIT 1"
        ).fetchone()
        cursor = conn.execute(
            """SELECT c.material_id, c.material, c.ot_id, SUM(c.cantidad_kg) AS cantidad, SUM(c.cmd_soles) AS cmd,
                      o.codigo_ot
               FROM cmd_ot c LEFT JOIN ordenes_trabajo o ON o.ot_id = c.ot_id
               WHERE c.run_id = ? GROUP BY c.material_id, c.ot_id""",
            (run[0] if run else -1,),
        )
    except Exception:
        return []
    columnas = [c[0] for c in cursor.description]
    cmd = {}
    for fila in (dict(zip(columnas, f)) for f in cursor.fetchall()):
        cmd.setdefault(fila["material_id"], []).append(fila)
    salida = []
    for item in items:
        filas = cmd.get(item.item_codigo)
        if not filas:
            continue
        cmd_q = sum(_num(f["cantidad"]) for f in filas)
        cmd_s = sum(_num(f["cmd"]) for f in filas)
        salida.append({
            "material_id": item.item_codigo, "material": filas[0]["material"], "ot_id": filas[0]["ot_id"],
            "codigo_ot": ", ".join(str(f.get("codigo_ot") or f["ot_id"]) for f in filas),
            "salida_cantidad": item.salidas_cantidad, "salida_soles": item.salidas_total,
            "cmd_cantidad": cmd_q, "cmd_soles": cmd_s,
            "diferencia_cantidad": item.salidas_cantidad - cmd_q, "diferencia_soles": item.salidas_total - cmd_s,
            "cuadra": abs(item.salidas_cantidad - cmd_q) < 1e-6 and abs(item.salidas_total - cmd_s) < 0.005,
        })
    return salida
