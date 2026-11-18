"""Periodos cargados (un Excel maestro por mes), calculo por mes y consolidacion de varios meses.

Flujo de la aplicacion:

1. Pantalla inicial -> "Importar": cada Excel se valida, se importa y se calcula
   (Modelo 1) en una base temporal. Se guardan los datos importados del mes
   (``mp_datos``, valores iniciales) y sus resultados (``mp_resultados``). El mes se
   identifica por la hoja Periodo (fecha_inicio / fecha_fin); solo periodos mensuales.
2. El usuario elige el tipo de analisis segun los meses consecutivos cargados
   (mensual, bimestral, trimestral, cuatrimestral, semestral, anual) y "Ejecutar".
   - Un mes: sus datos se cargan en la base de trabajo y se ejecuta el calculo ABC
     normal (mismos resultados de siempre).
   - Varios meses: se consolidan los resultados de cada mes (consolidar_en_base).
3. Simulacion: los cambios del usuario se guardan en ``sim_ajustes`` (por mes) y
   se aplican al recalcular; "Reestablecer valores iniciales" los borra.

Consolidacion de varios meses:

- Importes (soles, cantidades) ............ suma de los meses.
- Costo unitario OT ....................... costo total / cantidad producida del rango.
- Rentabilidad operativa .................. margen operativo / ventas netas del rango.
- Peso AC-OT efectivo ..................... costo asignado a la OT / costo de la actividad.
- Las OT de cada mes se agrupan por producto (codigo_producto): la OT del rango es
  la suma de las OT mensuales de ese producto.
- Kardex: saldo inicial del primer mes, compras y salidas sumadas, saldo final del
  ultimo mes.

Con un solo mes, la consolidacion devuelve exactamente los resultados de ese mes.
"""

from __future__ import annotations

import json
import re
import tempfile
from dataclasses import dataclass
from datetime import date
from pathlib import Path

from .constants import SHEET_TO_TABLE
from .database import db_session, initialize_database, latest_run_id, now_iso, quote_identifier, table_columns

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "setiembre", "octubre",
         "noviembre", "diciembre"]
MESES_CORTOS = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "set", "oct", "nov", "dic"]

TIPOS_ANALISIS = {
    "Mensual": 1,
    "Bimestral": 2,
    "Trimestral": 3,
    "Cuatrimestral": 4,
    "Semestral": 6,
    "Anual": 12,
}
NOMBRE_BLOQUE = {2: "Bimestre", 3: "Trimestre", 4: "Cuatrimestre", 6: "Semestre"}

TABLAS_CALCULADAS = [
    "kardex_promedio", "cmd_ot", "csd_ot", "mod_actividad", "otros_recursos_ac_detalle", "otros_recursos_actividad",
    "soporte_mantenimiento_ac", "soporte_control_calidad_ac", "soporte_ingenieria_ac", "soporte_actividad",
    "actividad_total", "produccion_ot_detalle", "produccion_ot", "soporte_ot_detalle", "soporte_ot",
    "hoja_costos_ot", "estado_resultados_ot", "dashboard_alertas",
]
TABLAS_DATOS = ["periodo", "ordenes_trabajo", "objetos_costo", "ventas", "compras"]
TABLAS_IMPORTADAS = list(dict.fromkeys(SHEET_TO_TABLE.values()))

SQL_TABLAS = """
CREATE TABLE IF NOT EXISTS mp_periodos (
    anio INTEGER NOT NULL, mes INTEGER NOT NULL, periodo_id TEXT, periodo TEXT, tipo_periodo TEXT,
    fecha_inicio TEXT, fecha_fin TEXT, archivo TEXT, origen TEXT, procesado_en TEXT,
    ventas_netas REAL, costo_total REAL, margen_operativo REAL, rentabilidad_operativa REAL,
    PRIMARY KEY (anio, mes)
);
CREATE TABLE IF NOT EXISTS mp_datos (
    anio INTEGER NOT NULL, mes INTEGER NOT NULL, tabla TEXT NOT NULL, filas_json TEXT,
    PRIMARY KEY (anio, mes, tabla)
);
CREATE TABLE IF NOT EXISTS mp_resultados (
    anio INTEGER NOT NULL, mes INTEGER NOT NULL, tabla TEXT NOT NULL, filas_json TEXT,
    PRIMARY KEY (anio, mes, tabla)
);
CREATE TABLE IF NOT EXISTS sim_ajustes (
    ajuste_id INTEGER PRIMARY KEY AUTOINCREMENT,
    anio INTEGER NOT NULL, mes INTEGER NOT NULL, tipo TEXT NOT NULL, tabla TEXT NOT NULL,
    fila INTEGER, clave TEXT, columna TEXT NOT NULL, valor_original REAL, valor_nuevo REAL,
    descripcion TEXT, creado_en TEXT
);
"""

SUFIJO_MES = re.compile(r"-\d{4}-\d{2}$")


# --------------------------------------------------------------------------- utilidades
def _num(valor) -> float:
    try:
        return float(valor or 0)
    except (TypeError, ValueError):
        return 0.0


def _div(a: float, b: float) -> float:
    return a / b if b else 0.0


def nombre_mes(anio: int, mes: int, corto: bool = False) -> str:
    nombre = (MESES_CORTOS if corto else MESES)[mes - 1]
    return f"{nombre}-{anio}" if corto else f"{nombre.capitalize()} {anio}"


def etiqueta_rango(meses: list[tuple[int, int]]) -> str:
    if not meses:
        return "Sin meses"
    meses = sorted(meses)
    (a0, m0), (a1, m1) = meses[0], meses[-1]
    if (a0, m0) == (a1, m1):
        return nombre_mes(a0, m0)
    if a0 == a1:
        return f"{MESES[m0 - 1].capitalize()}–{MESES[m1 - 1]} {a0}"
    return f"{nombre_mes(a0, m0)} – {nombre_mes(a1, m1)}"


def siguiente(mes: tuple[int, int]) -> tuple[int, int]:
    a, m = mes
    return (a + 1, 1) if m == 12 else (a, m + 1)


def meses_entre(desde: tuple[int, int], hasta: tuple[int, int]) -> list[tuple[int, int]]:
    if desde > hasta:
        desde, hasta = hasta, desde
    salida, actual = [], desde
    while actual <= hasta:
        salida.append(actual)
        actual = siguiente(actual)
    return salida


def bloques(tipo: str, anio: int) -> list[tuple[str, list[tuple[int, int]]]]:
    """Subperiodos calendario del año (enero inicia cada bloque); se usan en el comparativo."""
    tamano = TIPOS_ANALISIS.get(tipo)
    if not tamano:
        return []
    salida = []
    for i, inicio in enumerate(range(1, 13, tamano), start=1):
        meses = [(anio, m) for m in range(inicio, inicio + tamano)]
        if tamano == 1:
            etiqueta = nombre_mes(anio, inicio)
        elif tamano == 12:
            etiqueta = f"Año {anio}"
        else:
            etiqueta = f"{NOMBRE_BLOQUE[tamano]} {i}: {etiqueta_rango(meses)}"
        salida.append((etiqueta, meses))
    return salida


def ventanas(meses_cargados: list[tuple[int, int]]) -> dict[str, list[tuple[str, list[tuple[int, int]]]]]:
    """Opciones de analisis segun los meses cargados: cada tipo se ofrece solo si hay esa cantidad de
    meses consecutivos; las opciones son todas las ventanas de meses consecutivos de ese largo."""
    cargados = set(meses_cargados)
    salida: dict[str, list[tuple[str, list[tuple[int, int]]]]] = {}
    for tipo, tamano in TIPOS_ANALISIS.items():
        opciones = []
        for inicio in sorted(cargados):
            meses = [inicio]
            while len(meses) < tamano and siguiente(meses[-1]) in cargados:
                meses.append(siguiente(meses[-1]))
            if len(meses) == tamano:
                opciones.append((etiqueta_rango(meses), meses))
        if opciones:
            salida[tipo] = opciones
    return salida


def _mes_del_periodo(periodo: dict) -> tuple[int, int]:
    try:
        inicio = date.fromisoformat(str(periodo.get("fecha_inicio"))[:10])
    except ValueError as exc:
        raise ValueError(f"El periodo {periodo.get('periodo_id')} no tiene fecha_inicio valida.") from exc
    fin_txt = str(periodo.get("fecha_fin") or "")[:10]
    if fin_txt:
        try:
            fin = date.fromisoformat(fin_txt)
        except ValueError:
            fin = inicio
        if (fin.year, fin.month) != (inicio.year, inicio.month):
            raise ValueError(
                f"El periodo {periodo.get('periodo_id')} ({periodo.get('periodo')}) va de {inicio} a {fin}: "
                "solo se procesan periodos mensuales (un Excel por mes)."
            )
    return inicio.year, inicio.month


def leer_periodo_excel(excel_path: str | Path) -> dict:
    """Lee la hoja Periodo de un Excel maestro (sin importarlo) e identifica su mes."""
    import openpyxl

    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    try:
        if "Periodo" not in wb.sheetnames:
            raise ValueError(f"{Path(excel_path).name}: no tiene hoja Periodo.")
        filas = [r for r in wb["Periodo"].iter_rows(values_only=True)]
    finally:
        wb.close()
    encabezado = [str(c or "").strip() for c in filas[0]]
    datos = [dict(zip(encabezado, f)) for f in filas[1:] if any(v not in (None, "") for v in f)]
    if len(datos) != 1:
        raise ValueError(f"{Path(excel_path).name}: la hoja Periodo debe tener un solo periodo (tiene {len(datos)}).")
    periodo = datos[0]
    for campo in ("fecha_inicio", "fecha_fin"):
        valor = periodo.get(campo)
        if hasattr(valor, "isoformat"):
            periodo[campo] = valor.isoformat()[:10]
    periodo["anio"], periodo["mes"] = _mes_del_periodo(periodo)
    return periodo


# --------------------------------------------------------------------------- almacenamiento
def asegurar_tablas(conn) -> None:
    conn.executescript(SQL_TABLAS)


@dataclass
class PeriodoProcesado:
    anio: int
    mes: int
    periodo_id: str
    periodo: str
    archivo: str
    origen: str
    reemplazado: bool
    ventas: float
    costo_total: float
    margen_operativo: float

    @property
    def rentabilidad(self) -> float:
        return _div(self.margen_operativo, self.ventas)

    @property
    def etiqueta(self) -> str:
        return nombre_mes(self.anio, self.mes)


def _filas(conn, tabla: str, where: str = "", params: tuple = ()) -> list[dict]:
    try:
        return [dict(r) for r in conn.execute(f"SELECT * FROM {quote_identifier(tabla)} {where} ORDER BY rowid", params)]
    except Exception:
        return []


def volcar_resultados(conn) -> dict[str, list[dict]]:
    """Resultados de la ultima ejecucion (tablas calculadas) + datos que usa la consolidacion."""
    run_id = latest_run_id(conn)
    if run_id is None:
        raise ValueError("No hay una ejecucion de calculo completada.")
    datos: dict[str, list[dict]] = {}
    for tabla in TABLAS_CALCULADAS:
        filas = _filas(conn, tabla, "WHERE run_id = ?", (run_id,))
        for fila in filas:
            fila.pop("run_id", None)
        datos[tabla] = filas
    for tabla in TABLAS_DATOS:
        filas = _filas(conn, tabla)
        for fila in filas:
            fila.pop("batch_id", None)
        datos[tabla] = filas
    return datos


def volcar_datos(conn) -> dict[str, list[dict]]:
    """Datos importados (todas las hojas) en el orden de importacion."""
    salida = {}
    for tabla in TABLAS_IMPORTADAS:
        filas = _filas(conn, tabla)
        for fila in filas:
            fila.pop("batch_id", None)
        salida[tabla] = filas
    return salida


def _resumen(resultados: dict) -> tuple[float, float, float]:
    estado = resultados.get("estado_resultados_ot", [])
    hoja = resultados.get("hoja_costos_ot", [])
    return (sum(_num(f.get("ventas_netas_soles")) for f in estado), sum(_num(f.get("costo_total_ot")) for f in hoja),
            sum(_num(f.get("margen_operativo")) for f in estado))


def guardar_mes(destino_db: str | Path, datos: dict, resultados: dict, archivo: str = "", origen: str = "Excel") -> PeriodoProcesado:
    hoja = resultados.get("hoja_costos_ot") or []
    if not hoja:
        raise ValueError("La ejecucion no tiene hoja de costos.")
    periodo_ids = {fila.get("periodo_id") for fila in hoja}
    periodos = [p for p in datos.get("periodo", []) if p.get("periodo_id") in periodo_ids] or datos.get("periodo", [])
    if len(periodos) != 1:
        raise ValueError(f"Se esperaba un periodo mensual y se encontraron {len(periodos)} en {archivo or 'la base'}.")
    periodo = periodos[0]
    anio, mes = _mes_del_periodo(periodo)
    ventas, costo, margen = _resumen(resultados)
    with db_session(destino_db) as dst:
        asegurar_tablas(dst)
        reemplazado = dst.execute("SELECT 1 FROM mp_periodos WHERE anio = ? AND mes = ?", (anio, mes)).fetchone() is not None
        for tabla in ("mp_periodos", "mp_datos", "mp_resultados", "sim_ajustes"):
            dst.execute(f"DELETE FROM {tabla} WHERE anio = ? AND mes = ?", (anio, mes))
        dst.execute(
            """INSERT INTO mp_periodos (anio, mes, periodo_id, periodo, tipo_periodo, fecha_inicio, fecha_fin, archivo,
                   origen, procesado_en, ventas_netas, costo_total, margen_operativo, rentabilidad_operativa)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (anio, mes, periodo.get("periodo_id"), periodo.get("periodo"), periodo.get("tipo_periodo"),
             periodo.get("fecha_inicio"), periodo.get("fecha_fin"), str(archivo), origen, now_iso(),
             ventas, costo, margen, _div(margen, ventas)),
        )
        dst.executemany("INSERT INTO mp_datos (anio, mes, tabla, filas_json) VALUES (?, ?, ?, ?)",
                        [(anio, mes, t, json.dumps(f, ensure_ascii=False, default=str)) for t, f in datos.items()])
        dst.executemany("INSERT INTO mp_resultados (anio, mes, tabla, filas_json) VALUES (?, ?, ?, ?)",
                        [(anio, mes, t, json.dumps(f, ensure_ascii=False, default=str)) for t, f in resultados.items()])
    return PeriodoProcesado(anio, mes, str(periodo.get("periodo_id")), str(periodo.get("periodo")), str(archivo),
                            origen, reemplazado, ventas, costo, margen)


def guardar_resultados(origen_db: str | Path, destino_db: str | Path, archivo: str = "", origen: str = "Excel") -> PeriodoProcesado:
    """Guarda como un mes los datos importados y los resultados de la ultima ejecucion de ``origen_db``."""
    with db_session(origen_db) as src:
        datos = volcar_datos(src)
        resultados = volcar_resultados(src)
    return guardar_mes(destino_db, datos, resultados, archivo, origen)


def procesar_excel(excel_path: str | Path, db_path: str | Path) -> PeriodoProcesado:
    """Valida, importa y calcula un Excel mensual en una base temporal y guarda el mes."""
    from .costing_abc import run_model_1_calculation
    from .importers import import_excel_to_sqlite

    excel_path = Path(excel_path)
    leer_periodo_excel(excel_path)  # identifica el mes y rechaza periodos no mensuales antes de importar
    with tempfile.TemporaryDirectory(prefix="abc_periodo_") as tmp:
        base = Path(tmp) / "periodo.sqlite"
        initialize_database(base)
        resultado = import_excel_to_sqlite(excel_path, base)
        if not resultado.ok:
            errores = "; ".join(issue.message for issue in resultado.errors[:3])
            raise ValueError(f"{excel_path.name}: el Excel no pasa la validacion ({errores}).")
        run_model_1_calculation(base)
        return guardar_resultados(base, db_path, archivo=excel_path.name, origen="Excel")


def procesar_varios(rutas: list[str | Path], db_path: str | Path, progreso=None) -> tuple[list[PeriodoProcesado], list[tuple[str, str]]]:
    hechos, errores = [], []
    for i, ruta in enumerate(rutas, start=1):
        if progreso:
            progreso(i, len(rutas), Path(ruta).name)
        try:
            hechos.append(procesar_excel(ruta, db_path))
        except Exception as exc:  # se informa y se sigue con el siguiente archivo
            errores.append((Path(ruta).name, str(exc)))
    return hechos, errores


def periodos_procesados(db_path: str | Path) -> list[dict]:
    with db_session(db_path) as conn:
        asegurar_tablas(conn)
        filas = [dict(r) for r in conn.execute("SELECT * FROM mp_periodos ORDER BY anio, mes")]
    for fila in filas:
        fila["etiqueta"] = nombre_mes(fila["anio"], fila["mes"])
    return filas


def eliminar_periodos(db_path: str | Path, meses: list[tuple[int, int]] | None = None) -> int:
    with db_session(db_path) as conn:
        asegurar_tablas(conn)
        tablas = ("mp_periodos", "mp_datos", "mp_resultados", "sim_ajustes")
        if meses is None:
            n = conn.execute("SELECT COUNT(*) FROM mp_periodos").fetchone()[0]
            for tabla in tablas:
                conn.execute(f"DELETE FROM {tabla}")
            return n
        for anio, mes in meses:
            for tabla in tablas:
                conn.execute(f"DELETE FROM {tabla} WHERE anio = ? AND mes = ?", (anio, mes))
        return len(meses)


def iniciar_sesion_vacia(db_path: str | Path) -> None:
    """Al abrir la aplicacion no hay archivos cargados: borra meses, ajustes, datos y resultados de trabajo."""
    from .database import clear_calculated_tables, clear_imported_tables

    initialize_database(db_path)
    eliminar_periodos(db_path)
    with db_session(db_path) as conn:
        clear_imported_tables(conn)
        clear_calculated_tables(conn)


def _cargar_json(db_path: str | Path, tabla_mp: str, meses: list[tuple[int, int]]) -> dict[tuple[int, int], dict[str, list[dict]]]:
    datos: dict[tuple[int, int], dict[str, list[dict]]] = {}
    with db_session(db_path) as conn:
        asegurar_tablas(conn)
        for anio, mes in sorted(set(meses)):
            filas = conn.execute(f"SELECT tabla, filas_json FROM {tabla_mp} WHERE anio = ? AND mes = ?", (anio, mes))
            tablas = {r["tabla"]: json.loads(r["filas_json"] or "[]") for r in filas}
            if tablas:
                datos[(anio, mes)] = tablas
    return datos


def _cargar_meses(db_path: str | Path, meses: list[tuple[int, int]]) -> dict[tuple[int, int], dict[str, list[dict]]]:
    return _cargar_json(db_path, "mp_resultados", meses)


def datos_mes(db_path: str | Path, anio: int, mes: int, con_ajustes: bool = True) -> dict[str, list[dict]]:
    """Datos importados del mes (valores iniciales), con los ajustes de simulacion aplicados si se pide."""
    datos = _cargar_json(db_path, "mp_datos", [(anio, mes)]).get((anio, mes), {})
    if con_ajustes:
        for aj in ajustes(db_path, [(anio, mes)], tipo="dato"):
            filas = datos.get(aj["tabla"]) or []
            if aj["fila"] is not None and 0 <= aj["fila"] < len(filas):
                filas[aj["fila"]][aj["columna"]] = aj["valor_nuevo"]
    return datos


# --------------------------------------------------------------------------- simulacion (ajustes)
def registrar_ajuste(db_path: str | Path, anio: int, mes: int, tipo: str, tabla: str, columna: str,
                     valor_original: float | None, valor_nuevo: float, fila: int | None = None,
                     clave: str | None = None, descripcion: str = "") -> None:
    """Guarda (o reemplaza) un cambio de simulacion. tipo 'dato' = celda de una tabla importada (fila =
    indice en el mes); tipo 'resultado' = valor calculado fijado por el usuario (clave = actividad u OT).
    El valor original es el del Excel (o el calculado sin cambios), aunque se edite varias veces."""
    with db_session(db_path) as conn:
        asegurar_tablas(conn)
        previo = conn.execute(
            """SELECT ajuste_id, valor_original FROM sim_ajustes WHERE anio = ? AND mes = ? AND tipo = ? AND tabla = ?
                   AND columna = ? AND COALESCE(fila, -1) = COALESCE(?, -1) AND COALESCE(clave, '') = COALESCE(?, '')""",
            (anio, mes, tipo, tabla, columna, fila, clave),
        ).fetchone()
        if previo:
            valor_original = previo["valor_original"]
            conn.execute("DELETE FROM sim_ajustes WHERE ajuste_id = ?", (previo["ajuste_id"],))
        if valor_original is not None and abs(_num(valor_original) - _num(valor_nuevo)) < 1e-12:
            return  # volvio al valor inicial: no queda ajuste
        conn.execute(
            """INSERT INTO sim_ajustes (anio, mes, tipo, tabla, fila, clave, columna, valor_original, valor_nuevo,
                   descripcion, creado_en) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (anio, mes, tipo, tabla, fila, clave, columna, valor_original, valor_nuevo, descripcion, now_iso()),
        )


def ajustes(db_path: str | Path, meses: list[tuple[int, int]] | None = None, tipo: str | None = None) -> list[dict]:
    with db_session(db_path) as conn:
        asegurar_tablas(conn)
        filas = [dict(r) for r in conn.execute("SELECT * FROM sim_ajustes ORDER BY ajuste_id")]
    if meses is not None:
        meses = set(meses)
        filas = [f for f in filas if (f["anio"], f["mes"]) in meses]
    if tipo:
        filas = [f for f in filas if f["tipo"] == tipo]
    return filas


def borrar_ajustes(db_path: str | Path) -> int:
    with db_session(db_path) as conn:
        asegurar_tablas(conn)
        n = conn.execute("SELECT COUNT(*) FROM sim_ajustes").fetchone()[0]
        conn.execute("DELETE FROM sim_ajustes")
    return n


def ajustes_motor(db_path: str | Path, anio: int, mes: int) -> list[dict]:
    """Ajustes de resultados del mes en el formato de run_model_1_calculation(ajustes=...)."""
    return [{"tabla": a["tabla"], "clave": a["clave"], "columna": a["columna"], "valor": a["valor_nuevo"]}
            for a in ajustes(db_path, [(anio, mes)], tipo="resultado")]


def cargar_datos_en_base(conn, datos: dict[str, list[dict]]) -> None:
    """Reemplaza las tablas importadas de ``conn`` por los datos dados (rowid = posicion + 1)."""
    for tabla in TABLAS_IMPORTADAS:
        conn.execute(f"DELETE FROM {quote_identifier(tabla)}")
        filas = datos.get(tabla) or []
        if not filas:
            continue
        columnas = [c for c in table_columns(conn, tabla) if c != "batch_id"]
        usadas = [c for c in columnas if any(c in f for f in filas)]
        sql = (f"INSERT INTO {quote_identifier(tabla)} (rowid, {', '.join(quote_identifier(c) for c in usadas)}) "
               f"VALUES (?, {', '.join('?' for _ in usadas)})")
        conn.executemany(sql, [(i + 1, *[f.get(c) for c in usadas]) for i, f in enumerate(filas)])


def preparar_mes_en_base(db_path: str | Path, anio: int, mes: int, base_trabajo: str | Path | None = None) -> None:
    """Carga los datos del mes (con ajustes) en la base de trabajo (por defecto, la misma base)."""
    datos = datos_mes(db_path, anio, mes)
    with db_session(base_trabajo or db_path) as conn:
        cargar_datos_en_base(conn, datos)


def calcular_mes_en_base(db_path: str | Path, anio: int, mes: int) -> int:
    """Carga el mes en la base de trabajo y ejecuta el calculo ABC con los ajustes de simulacion."""
    from .costing_abc import run_model_1_calculation

    preparar_mes_en_base(db_path, anio, mes)
    return run_model_1_calculation(db_path, ajustes=ajustes_motor(db_path, anio, mes))


def calcular_mes(db_path: str | Path, anio: int, mes: int) -> dict[str, list[dict]]:
    """Recalcula un mes con sus ajustes en una base temporal y devuelve sus resultados."""
    from .costing_abc import run_model_1_calculation

    with tempfile.TemporaryDirectory(prefix="abc_sim_") as tmp:
        base = Path(tmp) / "mes.sqlite"
        initialize_database(base)
        preparar_mes_en_base(db_path, anio, mes, base_trabajo=base)
        run_model_1_calculation(base, ajustes=ajustes_motor(db_path, anio, mes))
        with db_session(base) as conn:
            return volcar_resultados(conn)


def resultados_meses(db_path: str | Path, meses: list[tuple[int, int]], simulados: bool = True) -> dict:
    """Resultados por mes: los guardados (valores iniciales) o recalculados si el mes tiene ajustes."""
    base = _cargar_meses(db_path, meses)
    if not simulados:
        return base
    con_ajustes = {(a["anio"], a["mes"]) for a in ajustes(db_path, meses)}
    for mes in sorted(con_ajustes):
        if mes in base:
            base[mes] = calcular_mes(db_path, *mes)
    return base


# --------------------------------------------------------------------------- consolidacion
CAMPOS_HOJA = ["cantidad_producida", "cmd_ot", "mod_ot", "otros_recursos_ot", "costo_produccion_ot", "costo_soporte_ot",
               "csd_ot", "costo_total_ot"]
CAMPOS_ESTADO = ["ventas_netas_soles", "materia_prima_directa", "costos_conversion", "margen_bruto", "procesos_soporte",
                 "margen_operativo"]
CAMPOS_ACTIVIDAD = ["costo_mod_ac", "costo_otros_recursos_ac", "costo_conversion_ac", "costo_soporte_ac", "costo_total_ac"]


def _codigo_ot_base(fila: dict, codigos: dict[str, str]) -> str:
    codigo = fila.get("codigo_ot") or codigos.get(fila.get("ot_id"), "") or fila.get("ot_id") or ""
    return SUFIJO_MES.sub("", str(codigo))


def _ot_por_mes(tablas: dict) -> dict[str, dict]:
    """ot_id -> {codigo_ot, codigo_producto} del mes."""
    productos = {o.get("objeto_costo_id"): o.get("codigo_producto") for o in tablas.get("objetos_costo", [])}
    salida = {}
    for ot in tablas.get("ordenes_trabajo", []):
        salida[ot["ot_id"]] = {"codigo_ot": ot.get("codigo_ot"), "codigo_producto": productos.get(ot.get("objeto_costo_id"))}
    for fila in tablas.get("hoja_costos_ot", []):
        info = salida.setdefault(fila["ot_id"], {})
        info["codigo_ot"] = info.get("codigo_ot") or fila.get("codigo_ot")
        info["codigo_producto"] = fila.get("codigo_producto") or info.get("codigo_producto")
    return salida


class _Grupo:
    def __init__(self, clave, base: dict):
        self.clave = clave
        self.base = dict(base)
        self.sumas: dict[str, float] = {}
        self.meses: list[tuple[int, int]] = []
        self.ots: list[str] = []
        self.desglose: list[dict] = []

    def sumar(self, fila: dict, campos: list[str], mes: tuple[int, int], extra: dict | None = None) -> None:
        for campo in campos:
            self.sumas[campo] = self.sumas.get(campo, 0.0) + _num(fila.get(campo))
        if mes not in self.meses:
            self.meses.append(mes)
        detalle = {"mes": nombre_mes(*mes)}
        detalle.update({c: _num(fila.get(c)) for c in campos})
        if extra:
            detalle.update(extra)
        self.desglose.append(detalle)


def _grupos(orden: list, grupos: dict) -> list[_Grupo]:
    return [grupos[c] for c in orden]


def analizar(db_path: str | Path, meses: list[tuple[int, int]], tipo: str = "Mensual",
             titulo: str | None = None, resultados: dict | None = None) -> dict:
    """Consolida los resultados mensuales (guardados, o los dados en ``resultados``) de los meses pedidos."""
    pedidos = sorted(set(meses))
    if resultados is None:
        datos = _cargar_meses(db_path, pedidos)
    else:
        datos = {m: resultados[m] for m in pedidos if m in resultados}
    con_datos = sorted(datos)
    faltantes = [m for m in pedidos if m not in datos]
    titulo = titulo or etiqueta_rango(pedidos)
    texto_meses = ", ".join(nombre_mes(a, m, corto=True) for a, m in con_datos)
    payload: dict = {
        "titulo": titulo, "tipo": tipo, "meses": pedidos, "meses_con_datos": con_datos, "meses_faltantes": faltantes,
        "completo": not faltantes and bool(con_datos), "texto_meses": texto_meses,
    }
    if not con_datos:
        return payload
    fuente = "Resultados mensuales procesados: " + ", ".join(
        f"{nombre_mes(a, m, corto=True)} ({(datos[(a, m)].get('periodo') or [{}])[0].get('periodo_id', '')})" for a, m in con_datos
    )
    rango = f"{titulo} ({len(con_datos)} {'mes' if len(con_datos) == 1 else 'meses'})"
    periodos = [p for (a, m) in con_datos for p in (datos[(a, m)].get("periodo") or [])[:1]]
    payload["periodo"] = [{
        "periodo_id": " + ".join(str(p.get("periodo_id")) for p in periodos),
        "periodo": titulo, "tipo_periodo": tipo.split(" (")[0],
        "fecha_inicio": min(str(p.get("fecha_inicio")) for p in periodos),
        "fecha_fin": max(str(p.get("fecha_fin")) for p in periodos),
    }]
    payload["fuente"] = fuente

    # ---- OT (hoja de costos + estado de resultados), agrupadas por producto
    hoja_g: dict[str, _Grupo] = {}
    est_g: dict[str, _Grupo] = {}
    orden_ot: list[str] = []
    prod_g: dict[tuple, _Grupo] = {}
    sop_g: dict[tuple, _Grupo] = {}
    orden_prod: list[tuple] = []
    orden_sop: list[tuple] = []
    act_g: dict[str, _Grupo] = {}
    orden_act: list[str] = []
    mensual = []
    for mes in con_datos:
        tablas = datos[mes]
        ots = _ot_por_mes(tablas)
        codigos = {k: v.get("codigo_ot") for k, v in ots.items()}
        for fila in tablas.get("hoja_costos_ot", []):
            prod = fila.get("codigo_producto") or fila.get("ot_id")
            if prod not in hoja_g:
                orden_ot.append(prod)
                hoja_g[prod] = _Grupo(prod, {"codigo_ot": _codigo_ot_base(fila, codigos), "codigo_producto": prod,
                                             "empresa_id": fila.get("empresa_id"), "modelo_id": fila.get("modelo_id")})
            g = hoja_g[prod]
            g.ots.append(str(fila.get("codigo_ot") or fila.get("ot_id")))
            g.sumar(fila, CAMPOS_HOJA, mes, {"codigo_ot": fila.get("codigo_ot"),
                                             "costo_unitario_ot": _num(fila.get("costo_unitario_ot"))})
        for fila in tablas.get("estado_resultados_ot", []):
            prod = fila.get("codigo_producto") or fila.get("ot_id")
            if prod not in est_g:
                est_g[prod] = _Grupo(prod, {"codigo_ot": _codigo_ot_base(fila, codigos), "codigo_producto": prod})
            est_g[prod].ots.append(str(codigos.get(fila.get("ot_id")) or fila.get("ot_id")))
            est_g[prod].sumar(fila, CAMPOS_ESTADO, mes, {"codigo_ot": codigos.get(fila.get("ot_id")),
                                                          "rentabilidad_operativa": _num(fila.get("rentabilidad_operativa"))})
        for fila in tablas.get("actividad_total", []):
            clave = fila.get("actividad_id")
            if clave not in act_g:
                orden_act.append(clave)
                act_g[clave] = _Grupo(clave, {k: fila.get(k) for k in ("actividad_id", "codigo_actividad", "nombre_actividad")})
            act_g[clave].sumar(fila, CAMPOS_ACTIVIDAD, mes)
        produccion_ot = {f["ot_id"]: f for f in tablas.get("produccion_ot", [])}
        for fila in tablas.get("produccion_ot_detalle", []):
            prod = fila.get("codigo_producto") or (ots.get(fila.get("ot_id")) or {}).get("codigo_producto")
            clave = (prod, fila.get("actividad_id"))
            if clave not in prod_g:
                orden_prod.append(clave)
                prod_g[clave] = _Grupo(clave, {k: fila.get(k) for k in ("actividad_id", "codigo_actividad", "nombre_actividad")}
                                       | {"codigo_ot": _codigo_ot_base(fila, codigos), "codigo_producto": prod})
            prod_g[clave].sumar(fila, ["costo_conversion_ac", "costo_conversion_ot", "costo_mod_ot", "costo_otros_recursos_ot"],
                                mes, {"peso_ac_ot": _num(fila.get("peso_ac_ot")), "codigo_ot": codigos.get(fila.get("ot_id")),
                                      "costo_produccion_ot": _num((produccion_ot.get(fila.get("ot_id")) or {}).get("costo_produccion_ot"))})
        soporte_ot = {f["ot_id"]: f for f in tablas.get("soporte_ot", [])}
        for fila in tablas.get("soporte_ot_detalle", []):
            prod = fila.get("codigo_producto") or (ots.get(fila.get("ot_id")) or {}).get("codigo_producto")
            clave = (prod, fila.get("actividad_id"))
            if clave not in sop_g:
                orden_sop.append(clave)
                sop_g[clave] = _Grupo(clave, {k: fila.get(k) for k in ("actividad_id", "codigo_actividad", "nombre_actividad")}
                                      | {"codigo_ot": _codigo_ot_base(fila, codigos), "codigo_producto": prod})
            sop_g[clave].sumar(fila, ["costo_soporte_ac", "costo_soporte_ot"], mes,
                               {"peso_ac_ot": _num(fila.get("peso_ac_ot")), "codigo_ot": codigos.get(fila.get("ot_id")),
                                "costo_soporte_ot_total": _num((soporte_ot.get(fila.get("ot_id")) or {}).get("costo_soporte_ot"))})
        hoja_mes = tablas.get("hoja_costos_ot", [])
        est_mes = tablas.get("estado_resultados_ot", [])
        fila_mes = {"mes": nombre_mes(*mes), "anio": mes[0], "num_mes": mes[1],
                    "periodo_id": (tablas.get("periodo") or [{}])[0].get("periodo_id")}
        for campo in CAMPOS_HOJA:
            fila_mes[campo] = sum(_num(f.get(campo)) for f in hoja_mes)
        for campo in CAMPOS_ESTADO:
            fila_mes[campo] = sum(_num(f.get(campo)) for f in est_mes)
        fila_mes["costo_unitario_medio"] = _div(fila_mes["costo_total_ot"], fila_mes["cantidad_producida"])
        fila_mes["rentabilidad_operativa"] = _div(fila_mes["margen_operativo"], fila_mes["ventas_netas_soles"])
        mensual.append(fila_mes)
    payload["mensual"] = mensual

    nota = "" if not faltantes else f" Faltan {len(faltantes)} mes(es): " + ", ".join(nombre_mes(a, m, corto=True) for a, m in faltantes) + "."
    # Hoja de costos
    hoja = []
    for g in _grupos(orden_ot, hoja_g):
        s = g.sumas
        fila = {"periodo_analizado": rango, "empresa_id": g.base.get("empresa_id"),
                "periodo_id": payload["periodo"][0]["periodo_id"], "modelo_id": g.base.get("modelo_id"),
                "ot_id": g.clave, "codigo_ot": g.base["codigo_ot"], "ot_incluidas": ", ".join(g.ots),
                "objeto_costo_id": g.clave, "codigo_producto": g.clave}
        fila.update({c: s.get(c, 0.0) for c in CAMPOS_HOJA})
        fila["costo_unitario_ot"] = _div(s.get("costo_total_ot", 0.0), s.get("cantidad_producida", 0.0))
        fila["formula"] = ("Suma de las OT mensuales del producto; costo_unitario_OT = costo_total_OT del rango / "
                           "cantidad_producida del rango")
        fila["fuente"] = fuente
        fila["observacion"] = f"Consolidado de {len(g.meses)} mes(es).{nota}"
        hoja.append(fila)
    # Estado de resultados
    estado = []
    for prod in orden_ot:
        g = est_g.get(prod)
        if g is None:
            continue
        s = g.sumas
        fila = {"periodo_analizado": rango, "ot_id": prod, "codigo_ot": g.base["codigo_ot"], "ot_incluidas": ", ".join(g.ots),
                "objeto_costo_id": prod, "codigo_producto": prod}
        fila.update({c: s.get(c, 0.0) for c in CAMPOS_ESTADO})
        fila["rentabilidad_operativa"] = _div(s.get("margen_operativo", 0.0), s.get("ventas_netas_soles", 0.0))
        fila["formula"] = ("Suma de los meses; margen_bruto = ventas - MPD - conversion; margen_operativo = margen_bruto - "
                           "soporte; rentabilidad = margen_operativo / ventas del rango")
        fila["fuente"] = fuente
        fila["observacion"] = f"Consolidado de {len(g.meses)} mes(es).{nota}"
        estado.append(fila)
    # Proceso 1 (actividades)
    actividades, proceso1 = [], []
    for g in _grupos(orden_act, act_g):
        s = g.sumas
        actividades.append({**g.base, **{c: s.get(c, 0.0) for c in CAMPOS_ACTIVIDAD}})
        proceso1.append({"periodo_analizado": rango, **g.base, "costo_pd_ac": s.get("costo_mod_ac", 0.0),
                         "costo_otros_recursos_ac": s.get("costo_otros_recursos_ac", 0.0),
                         "costo_conversion_ac": s.get("costo_conversion_ac", 0.0),
                         "costo_soporte_ac": s.get("costo_soporte_ac", 0.0), "costo_total_ac": s.get("costo_total_ac", 0.0),
                         "formula": "Suma mensual: costo_total_AC = PD_AC + otros_recursos_AC + soporte_AC",
                         "fuente": fuente, "observacion": f"Consolidado de {len(g.meses)} mes(es).{nota}"})
    # Proceso 2 produccion y soporte
    produccion_tot = {prod: hoja_g[prod].sumas.get("costo_produccion_ot", 0.0) for prod in orden_ot}
    soporte_tot = {prod: hoja_g[prod].sumas.get("costo_soporte_ot", 0.0) for prod in orden_ot}
    proceso2p = []
    for g in _grupos(orden_prod, prod_g):
        s = g.sumas
        proceso2p.append({"periodo_analizado": rango, **{k: g.base[k] for k in ("actividad_id", "codigo_actividad", "nombre_actividad")},
                          "ot_id": g.base["codigo_producto"], "codigo_ot": g.base["codigo_ot"], "objeto_costo_id": g.base["codigo_producto"],
                          "codigo_producto": g.base["codigo_producto"],
                          "peso_ac_ot": _div(s.get("costo_conversion_ot", 0.0), s.get("costo_conversion_ac", 0.0)),
                          "costo_conversion_ac": s.get("costo_conversion_ac", 0.0), "costo_conversion_ot": s.get("costo_conversion_ot", 0.0),
                          "costo_produccion_ot": produccion_tot.get(g.base["codigo_producto"], 0.0),
                          "formula": "Suma mensual de costo_conversion_OT; peso efectivo = costo_conversion_OT / costo_conversion_AC del rango",
                          "fuente": fuente, "observacion": f"Consolidado de {len(g.meses)} mes(es).{nota}"})
    proceso2s = []
    for g in _grupos(orden_sop, sop_g):
        s = g.sumas
        proceso2s.append({"periodo_analizado": rango, **{k: g.base[k] for k in ("actividad_id", "codigo_actividad", "nombre_actividad")},
                          "ot_id": g.base["codigo_producto"], "codigo_ot": g.base["codigo_ot"], "objeto_costo_id": g.base["codigo_producto"],
                          "codigo_producto": g.base["codigo_producto"],
                          "peso_ac_ot": _div(s.get("costo_soporte_ot", 0.0), s.get("costo_soporte_ac", 0.0)),
                          "costo_soporte_ac": s.get("costo_soporte_ac", 0.0), "costo_soporte_ot_detalle": s.get("costo_soporte_ot", 0.0),
                          "costo_soporte_ot": soporte_tot.get(g.base["codigo_producto"], 0.0),
                          "formula": "Suma mensual de costo_soporte_OT; peso efectivo = costo_soporte_OT / costo_soporte_AC del rango",
                          "fuente": fuente, "observacion": f"Consolidado de {len(g.meses)} mes(es).{nota}"})

    payload.update({
        "hoja": hoja, "estado": estado, "proceso1": proceso1, "proceso2_produccion": proceso2p, "proceso2_soporte": proceso2s,
        "costos_ot": hoja, "estado_resultados": estado, "costos_actividad": actividades,
        "desglose": {
            "hoja": {g.clave: g.desglose for g in hoja_g.values()},
            "estado": {g.clave: g.desglose for g in est_g.values()},
            "proceso1": {g.clave: g.desglose for g in act_g.values()},
            "proceso2_produccion": {g.clave: g.desglose for g in prod_g.values()},
            "proceso2_soporte": {g.clave: g.desglose for g in sop_g.values()},
        },
    })
    payload["kardex"] = _kardex_rango(datos, con_datos)
    payload["kpis"] = _kpis(hoja, estado)
    payload["alertas"] = _alertas(payload)
    return payload


def clave_desglose(modulo: str, fila: dict):
    if modulo in ("hoja", "estado"):
        return fila.get("codigo_producto")
    if modulo == "proceso1":
        return fila.get("actividad_id")
    return (fila.get("codigo_producto"), fila.get("actividad_id"))


def _kpis(hoja: list[dict], estado: list[dict]) -> dict:
    costo = sum(f["costo_total_ot"] for f in hoja)
    unidades = sum(f["cantidad_producida"] for f in hoja)
    ventas = sum(f["ventas_netas_soles"] for f in estado)
    margen = sum(f["margen_operativo"] for f in estado)
    return {
        "costo_total": costo,
        # mismas definiciones del dashboard de un periodo: promedio simple entre OT
        "costo_unitario_prom": _div(sum(f["costo_unitario_ot"] for f in hoja), len(hoja)),
        "ventas": ventas,
        "rentabilidad_prom": _div(sum(f["rentabilidad_operativa"] for f in estado), len(estado)),
        # adicionales del consolidado
        "unidades": unidades,
        "costo_unitario_medio": _div(costo, unidades),
        "margen_bruto": sum(f["margen_bruto"] for f in estado),
        "margen_operativo": margen,
        "rentabilidad_global": _div(margen, ventas),
    }


def _alertas(payload: dict) -> list[dict]:
    alertas = []
    actividades = payload.get("costos_actividad", [])
    if actividades:
        top = max(actividades, key=lambda f: f["costo_total_ac"])
        alertas.append({"tipo": "Actividad relevante", "severidad": "Media",
                        "mensaje": f"La actividad con mayor costo total del periodo analizado es {top['nombre_actividad']}.",
                        "valor_referencia": top["costo_total_ac"], "observacion": "Revisar consumo de MOD, otros recursos y soporte."})
    estado = payload.get("estado", [])
    if estado:
        bajo = min(estado, key=lambda f: f["rentabilidad_operativa"])
        alertas.append({"tipo": "Rentabilidad", "severidad": "Alta",
                        "mensaje": f"El objeto de costo con menor rentabilidad operativa del periodo es la {bajo['codigo_ot']} "
                                   f"({bajo['codigo_producto']}).",
                        "valor_referencia": bajo["rentabilidad_operativa"],
                        "observacion": "Evaluar precio, mezcla de costos y actividades de mayor peso."})
    mensual = payload.get("mensual", [])
    negativos = [f for f in mensual if f["margen_operativo"] < 0]
    if negativos:
        alertas.append({"tipo": "Meses con perdida", "severidad": "Alta",
                        "mensaje": f"{len(negativos)} de {len(mensual)} meses con margen operativo negativo: "
                                   + ", ".join(f["mes"] for f in negativos) + ".",
                        "valor_referencia": sum(f["margen_operativo"] for f in negativos),
                        "observacion": "Suma del margen operativo de esos meses."})
    if len(mensual) > 1:
        mejor = max(mensual, key=lambda f: f["rentabilidad_operativa"])
        peor = min(mensual, key=lambda f: f["rentabilidad_operativa"])
        alertas.append({"tipo": "Evolucion", "severidad": "Media",
                        "mensaje": f"Mejor mes: {mejor['mes']} ({mejor['rentabilidad_operativa'] * 100:,.2f}%); "
                                   f"peor mes: {peor['mes']} ({peor['rentabilidad_operativa'] * 100:,.2f}%).",
                        "valor_referencia": mejor["rentabilidad_operativa"] - peor["rentabilidad_operativa"],
                        "observacion": "Diferencia de rentabilidad operativa entre el mejor y el peor mes."})
    if payload.get("meses_faltantes"):
        alertas.append({"tipo": "Datos incompletos", "severidad": "Alta",
                        "mensaje": "Periodo incompleto: faltan " + ", ".join(nombre_mes(a, m) for a, m in payload["meses_faltantes"])
                                   + ". Procese esos Excel mensuales para completar el analisis.",
                        "valor_referencia": len(payload["meses_faltantes"]), "observacion": ""})
    return alertas


def _kardex_rango(datos: dict, meses: list[tuple[int, int]]) -> list[dict]:
    from .kardex_compras import valorizar

    resumen: dict[str, dict] = {}
    orden = []
    for mes in meses:
        tablas = datos[mes]
        compras = tablas.get("compras") or []
        if not compras:
            continue
        cmd_mes: dict[str, tuple[float, float]] = {}
        for c in tablas.get("cmd_ot", []):
            q, t = cmd_mes.get(c.get("material_id"), (0.0, 0.0))
            cmd_mes[c.get("material_id")] = (q + _num(c.get("cantidad_kg")), t + _num(c.get("cmd_soles")))
        for item in valorizar(compras):
            r = resumen.get(item.item_codigo)
            cmd_q, cmd = cmd_mes.get(item.item_codigo, (0.0, 0.0))
            if r is None:
                orden.append(item.item_codigo)
                r = resumen[item.item_codigo] = {
                    "material_id": item.item_codigo, "material": item.item_descripcion, "unidad": item.unidad,
                    "saldo_inicial_cantidad": item.saldo_inicial_cantidad, "saldo_inicial_soles": round(item.saldo_inicial_total, 2),
                    "compras_cantidad": 0.0, "compras_soles": 0.0, "salidas_cantidad": 0.0, "salidas_soles": 0.0,
                    "cmd_cantidad": 0.0, "cmd_soles": 0.0, "meses": 0, "continuidad": "OK", "_fin": None,
                }
            elif r["_fin"] is not None:
                q, v = r["_fin"]
                if abs(q - item.saldo_inicial_cantidad) > 1e-6 or abs(v - item.saldo_inicial_total) > 0.01:
                    r["continuidad"] = f"Revisar {nombre_mes(*mes, corto=True)}"
            # importes de cada mes al centimo (como se registran en el kardex mensual)
            r["compras_cantidad"] += item.ingresos_cantidad
            r["compras_soles"] += round(item.ingresos_total, 2)
            r["salidas_cantidad"] += item.salidas_cantidad
            r["salidas_soles"] += round(item.salidas_total, 2)
            r["cmd_cantidad"] += cmd_q
            r["cmd_soles"] += round(cmd, 2)
            r["saldo_final_cantidad"] = item.saldo_final_cantidad
            r["saldo_final_soles"] = round(item.saldo_final_total, 2)
            r["saldo_final_cto_unit"] = item.saldo_final_costo_unitario
            r["meses"] += 1
            r["_fin"] = (item.saldo_final_cantidad, item.saldo_final_total)
    salida = []
    for codigo in orden:
        r = resumen[codigo]
        r.pop("_fin", None)
        r["diferencia_vs_cmd"] = round(r["salidas_soles"] - r["cmd_soles"], 2)
        r["salidas_vs_cmd"] = "coincide" if r["diferencia_vs_cmd"] == 0 and abs(r["salidas_cantidad"] - r["cmd_cantidad"]) < 1e-6 else "NO coincide"
        # SI + compras - salidas - SF: centimos por redondear cada mes (saldos e importes al centimo)
        r["ajuste_redondeo"] = round(r["saldo_inicial_soles"] + r["compras_soles"] - r["salidas_soles"] - r["saldo_final_soles"], 2)
        salida.append(r)
    return salida


def comparar_bloques(db_path: str | Path, tipo: str, anio: int, resultados: dict | None = None) -> list[dict]:
    """Resumen de cada subperiodo calendario del año (p. ej. T1..T4) con los meses disponibles."""
    salida = []
    for etiqueta, meses in bloques(tipo, anio):
        p = analizar(db_path, meses, tipo, etiqueta, resultados=resultados)
        k = p.get("kpis") or {}
        salida.append({
            "subperiodo": etiqueta, "meses_con_datos": f"{len(p['meses_con_datos'])}/{len(meses)}",
            "unidades": k.get("unidades", 0.0), "ventas_netas_soles": k.get("ventas", 0.0),
            "costo_total_ot": k.get("costo_total", 0.0), "margen_bruto": k.get("margen_bruto", 0.0),
            "margen_operativo": k.get("margen_operativo", 0.0), "rentabilidad_operativa": k.get("rentabilidad_global", 0.0),
            "costo_unitario_medio": k.get("costo_unitario_medio", 0.0),
            "estado": "Sin datos" if not p["meses_con_datos"] else ("Completo" if p["completo"] else "Incompleto"),
            "_meses": meses,
        })
    return salida


# --------------------------------------------------------------------------- consolidado en la base de trabajo
# (claves de agrupacion, campos que se suman, razones recalculadas = numerador / denominador)
AGREGADOS = {
    "kardex_promedio": (["material_id", "ot_id"], ["entrada_cantidad", "salida_cantidad", "costo_salida"],
                        {"costo_promedio": ("costo_salida", "salida_cantidad")}),
    "cmd_ot": (["ot_id", "material_id"], ["cantidad_kg", "cmd_soles"], {"costo_unitario_promedio": ("cmd_soles", "cantidad_kg")}),
    "csd_ot": (["ot_id", "servicio_id"], ["cantidad", "csd_soles"], {"costo_unitario_soles": ("csd_soles", "cantidad")}),
    "mod_actividad": (["actividad_id"], ["tiempo_real_minutos", "total_mod_soles", "costo_mod_ac"],
                      {"porcentaje_usado": ("costo_mod_ac", "total_mod_soles")}),
    "otros_recursos_ac_detalle": (["elemento_costo_id", "actividad_id"], ["importe_recurso_soles", "costo_recurso_ac"],
                                  {"porcentaje_inductor": ("costo_recurso_ac", "importe_recurso_soles")}),
    "otros_recursos_actividad": (["actividad_id"], ["costo_otros_recursos_ac"], {}),
    "soporte_mantenimiento_ac": (["actividad_id"], ["costo_mantenimiento_ac"], {}),
    "soporte_control_calidad_ac": (["actividad_id"], ["costo_total_control_calidad", "costo_control_calidad_ac"],
                                   {"porcentaje_distribucion": ("costo_control_calidad_ac", "costo_total_control_calidad")}),
    "soporte_ingenieria_ac": (["actividad_id"], ["costo_total_ingenieria", "costo_ingenieria_ac"],
                              {"porcentaje_distribucion": ("costo_ingenieria_ac", "costo_total_ingenieria")}),
    "soporte_actividad": (["actividad_id"], ["costo_mantenimiento_ac", "costo_control_calidad_ac", "costo_ingenieria_ac",
                                             "costo_soporte_ac"], {}),
    "actividad_total": (["actividad_id"], CAMPOS_ACTIVIDAD, {}),
    "produccion_ot_detalle": (["ot_id", "actividad_id"], ["costo_conversion_ac", "costo_mod_ac", "costo_otros_recursos_ac",
                                                          "costo_conversion_ot", "costo_mod_ot", "costo_otros_recursos_ot"],
                              {"peso_ac_ot": ("costo_conversion_ot", "costo_conversion_ac")}),
    "produccion_ot": (["ot_id"], ["costo_mod_ot", "costo_otros_recursos_ot", "costo_produccion_ot"], {}),
    "soporte_ot_detalle": (["ot_id", "actividad_id"], ["costo_soporte_ac", "costo_soporte_ot"],
                           {"peso_ac_ot": ("costo_soporte_ot", "costo_soporte_ac")}),
    "soporte_ot": (["ot_id"], ["costo_soporte_ot"], {}),
    "hoja_costos_ot": (["ot_id"], CAMPOS_HOJA, {"costo_unitario_ot": ("costo_total_ot", "cantidad_producida")}),
    "estado_resultados_ot": (["ot_id"], CAMPOS_ESTADO, {"rentabilidad_operativa": ("margen_operativo", "ventas_netas_soles")}),
}


def consolidar_tabla(tabla: str, resultados: dict, meses: list[tuple[int, int]]) -> list[dict]:
    """Agrega una tabla calculada de varios meses (las OT mensuales se agrupan por producto)."""
    claves, sumas, razones = AGREGADOS[tabla]
    grupos: dict[tuple, dict] = {}
    conteo: dict[tuple, int] = {}
    promedios: dict[tuple, dict[str, list[float]]] = {}
    for mes in meses:
        tablas = resultados.get(mes) or {}
        ots = _ot_por_mes(tablas)
        for fila in tablas.get(tabla, []):
            fila = dict(fila)
            if "ot_id" in fila and fila.get("ot_id") in ots:
                info = ots[fila["ot_id"]]
                producto = info.get("codigo_producto") or fila["ot_id"]
                fila["ot_id"] = producto
                if "objeto_costo_id" in fila:
                    fila["objeto_costo_id"] = producto
                if "codigo_ot" in fila:
                    fila["codigo_ot"] = SUFIJO_MES.sub("", str(fila.get("codigo_ot") or info.get("codigo_ot") or ""))
            clave = tuple(fila.get(c) for c in claves)
            if clave not in grupos:
                grupos[clave] = dict(fila)
                for campo in sumas:
                    grupos[clave][campo] = 0.0
                conteo[clave] = 0
                promedios[clave] = {}
            g = grupos[clave]
            conteo[clave] += 1
            for campo in sumas:
                g[campo] += _num(fila.get(campo))
            for campo, valor in fila.items():
                if campo not in sumas and campo not in razones and isinstance(valor, (int, float)) and not isinstance(valor, bool):
                    promedios[clave].setdefault(campo, []).append(float(valor))
    salida = []
    for clave, g in grupos.items():
        for campo, valores in promedios[clave].items():
            g[campo] = sum(valores) / len(valores)
        for campo, (num, den) in razones.items():
            g[campo] = _div(g.get(num, 0.0), g.get(den, 0.0))
        if "observacion" in g:
            g["observacion"] = f"Consolidado de {conteo[clave]} mes(es): {etiqueta_rango(meses)}."
        salida.append(g)
    return salida


def consolidar_en_base(db_path: str | Path, payload: dict, resultados: dict, base_trabajo: str | Path | None = None) -> int:
    """Escribe el consolidado del periodo como una ejecucion en las tablas calculadas de la base de trabajo
    (asi el dashboard, los reportes, los documentos y el agente leen los resultados del periodo)."""
    from .database import clear_calculated_tables, finish_calculation_run, start_calculation_run

    meses = payload.get("meses_con_datos") or []
    with db_session(base_trabajo or db_path) as conn:
        clear_calculated_tables(conn)
        run_id = start_calculation_run(conn, "MODELO_1", f"Consolidado {payload.get('titulo')} ({len(meses)} meses)")
        for tabla in AGREGADOS:
            filas = consolidar_tabla(tabla, resultados, meses)
            if tabla == "hoja_costos_ot":
                periodo_id = (payload.get("periodo") or [{}])[0].get("periodo_id")
                for fila in filas:
                    fila["periodo_id"] = periodo_id
            columnas = [c for c in table_columns(conn, tabla) if c not in ("kardex_id", "alerta_id")]
            for fila in filas:
                fila["run_id"] = run_id
                usadas = [c for c in columnas if c in fila]
                conn.execute(
                    f"INSERT INTO {tabla} ({', '.join(usadas)}) VALUES ({', '.join('?' for _ in usadas)})",
                    [fila.get(c) for c in usadas],
                )
        for alerta in payload.get("alertas", []):
            conn.execute(
                "INSERT INTO dashboard_alertas (run_id, tipo, severidad, mensaje, valor_referencia, observacion) VALUES (?, ?, ?, ?, ?, ?)",
                (run_id, alerta["tipo"], alerta["severidad"], alerta["mensaje"], alerta.get("valor_referencia"), alerta.get("observacion")),
            )
        finish_calculation_run(conn, run_id, "COMPLETADO", f"Consolidado de {etiqueta_rango(meses)}.")
    return run_id
