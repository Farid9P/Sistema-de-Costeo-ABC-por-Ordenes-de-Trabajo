"""Kardex de valorizacion de inventarios por PROMEDIO PONDERADO (movil).

Mismas reglas del kardex del sistema (inventory_average.py):
- Entrada: costo unitario de entrada = (cantidad x precio + flete) / cantidad;
  el saldo se revaloriza y el promedio = valor del saldo / cantidad del saldo.
- Salida: se valoriza al costo promedio vigente y va a la OT que consume (CMD).
- No se permite stock negativo.
- Ecuacion del almacen: Inventario inicial + Compras - Inventario final = Consumo.

Este modulo trabaja en memoria: el agente lo usa para leer el kardex de la hoja
Compras (fuente del CMD) y verificar lo guardado en ``kardex_promedio``; nunca
escribe en la base.
"""

from __future__ import annotations

from dataclasses import dataclass, field

TOLERANCIA = 0.005  # medio centimo


class StockInsuficiente(ValueError):
    pass


@dataclass
class FilaKardex:
    detalle: str
    tipo: str  # SALDO INICIAL | ENTRADA | SALIDA
    referencia: str
    ent_cantidad: float = 0.0
    ent_costo_unitario: float = 0.0
    ent_total: float = 0.0
    sal_cantidad: float = 0.0
    sal_costo_unitario: float = 0.0
    sal_total: float = 0.0
    saldo_cantidad: float = 0.0
    saldo_costo_promedio: float = 0.0
    saldo_total: float = 0.0


@dataclass
class KardexPromedioPonderado:
    material: str
    material_id: str = ""
    filas: list[FilaKardex] = field(default_factory=list)
    saldo_cantidad: float = 0.0
    saldo_total: float = 0.0
    inventario_inicial: float = 0.0
    compras: float = 0.0
    consumo: float = 0.0

    @property
    def costo_promedio(self) -> float:
        return self.saldo_total / self.saldo_cantidad if self.saldo_cantidad else 0.0

    def saldo_inicial(self, cantidad: float, costo_unitario: float, detalle: str = "Saldo inicial") -> FilaKardex:
        self.saldo_cantidad = cantidad
        self.saldo_total = cantidad * costo_unitario
        self.inventario_inicial = self.saldo_total
        fila = FilaKardex(detalle, "SALDO INICIAL", "", saldo_cantidad=cantidad,
                          saldo_costo_promedio=self.costo_promedio, saldo_total=self.saldo_total)
        self.filas.append(fila)
        return fila

    def entrada(self, cantidad: float, precio_unitario: float, flete: float = 0.0,
                detalle: str = "Ingreso al almacen", referencia: str = "") -> FilaKardex:
        if cantidad <= 0:
            raise ValueError("La cantidad de entrada debe ser mayor que cero.")
        total = cantidad * precio_unitario + flete
        costo_unitario = total / cantidad
        self.saldo_cantidad += cantidad
        self.saldo_total += total
        self.compras += total
        fila = FilaKardex(detalle, "ENTRADA", referencia, ent_cantidad=cantidad, ent_costo_unitario=costo_unitario,
                          ent_total=total, saldo_cantidad=self.saldo_cantidad,
                          saldo_costo_promedio=self.costo_promedio, saldo_total=self.saldo_total)
        self.filas.append(fila)
        return fila

    def salida(self, cantidad: float, detalle: str = "Consumo", referencia: str = "") -> FilaKardex:
        if cantidad <= 0:
            raise ValueError("La cantidad de salida debe ser mayor que cero.")
        if cantidad - self.saldo_cantidad > 1e-9:
            raise StockInsuficiente(
                f"Stock insuficiente de {self.material}: saldo {self.saldo_cantidad:,.2f}, salida {cantidad:,.2f}."
            )
        promedio = self.costo_promedio
        total = cantidad * promedio
        self.saldo_cantidad -= cantidad
        self.saldo_total -= total
        if abs(self.saldo_cantidad) < 1e-9:
            self.saldo_cantidad = 0.0
            self.saldo_total = 0.0
        self.consumo += total
        fila = FilaKardex(detalle, "SALIDA", referencia, sal_cantidad=cantidad, sal_costo_unitario=promedio,
                          sal_total=total, saldo_cantidad=self.saldo_cantidad,
                          saldo_costo_promedio=promedio if self.saldo_cantidad else 0.0, saldo_total=self.saldo_total)
        self.filas.append(fila)
        return fila

    @property
    def inventario_final(self) -> float:
        return self.saldo_total

    def ecuacion_almacen(self) -> dict:
        consumo_ecuacion = self.inventario_inicial + self.compras - self.inventario_final
        return {
            "inventario_inicial": self.inventario_inicial,
            "compras": self.compras,
            "inventario_final": self.inventario_final,
            "consumo": consumo_ecuacion,
            "cuadra": abs(consumo_ecuacion - self.consumo) <= TOLERANCIA,
        }


def kardex_desde_base(lector) -> tuple[list[KardexPromedioPonderado], list[str]]:
    """Kardex de cada material directo a partir de la hoja Compras (fuente del CMD), valorizado con el
    mismo motor del sistema (app.kardex_compras), y conciliado con lo guardado en ``kardex_promedio``.

    Devuelve los kardex y una lista de observaciones de conciliacion. Solo lectura.
    """
    from ..kardex_compras import COMPRA, CONSUMO, DEVOLUCION_CONSUMO, DEVOLUCION_PROVEEDOR, SALDO_INICIAL, valorizar

    movimientos = lector.tabla("compras") if lector.existe_tabla("compras") else []
    guardado: dict[str, float] = {}
    for row in lector.calculado("kardex_promedio"):
        guardado[row.get("material_id")] = guardado.get(row.get("material_id"), 0.0) + float(row.get("costo_salida") or 0)
    kardexs: list[KardexPromedioPonderado] = []
    observaciones: list[str] = []
    for item in valorizar(movimientos):
        k = KardexPromedioPonderado(material=item.item_descripcion or item.item_codigo, material_id=item.item_codigo)
        for f in item.filas:
            if f.tipo == SALDO_INICIAL:
                tipo, detalle = "SALDO INICIAL", "Saldo inicial"
            elif f.tipo in (COMPRA, DEVOLUCION_PROVEEDOR):
                tipo, detalle = "ENTRADA", f.descripcion
            else:
                tipo, detalle = "SALIDA", f.descripcion
            k.filas.append(FilaKardex(
                detalle=detalle, tipo=tipo, referencia=f.destino or f.documento or "",
                ent_cantidad=f.ing_cantidad or 0.0, ent_costo_unitario=f.ing_costo_unitario or 0.0, ent_total=f.ing_total or 0.0,
                sal_cantidad=f.sal_cantidad or 0.0, sal_costo_unitario=f.sal_costo_unitario or 0.0, sal_total=f.sal_total or 0.0,
                saldo_cantidad=f.saldo_cantidad, saldo_costo_promedio=f.saldo_costo_unitario, saldo_total=f.saldo_total))
        k.inventario_inicial = item.saldo_inicial_total
        k.compras = item.ingresos_total
        k.consumo = item.salidas_total
        k.saldo_cantidad = item.saldo_final_cantidad
        k.saldo_total = item.saldo_final_total
        kardexs.append(k)
        observaciones += [f"{k.material}: {aviso}" for aviso in item.avisos]
        if item.item_codigo in guardado and abs(guardado[item.item_codigo] - k.consumo) > TOLERANCIA:
            observaciones.append(f"{k.material}: el CMD guardado difiere del kardex recalculado "
                                 f"(S/ {abs(guardado[item.item_codigo] - k.consumo):,.2f}).")
    if movimientos and not guardado:
        observaciones.append("No hay kardex guardado para la ultima ejecucion; se muestra el kardex recalculado en memoria.")
    return kardexs, observaciones


def salidas_por_material(k: KardexPromedioPonderado) -> tuple[float, float, list[str]]:
    """Cantidad, costo total y destinos de las salidas de un kardex."""
    salidas = [f for f in k.filas if f.tipo == "SALIDA"]
    destinos = sorted({f.referencia for f in salidas if f.referencia})
    return sum(f.sal_cantidad for f in salidas), sum(f.sal_total for f in salidas), destinos
