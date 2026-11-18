# Sistema Informatico de Costos ABC - Frenosa

Aplicativo desktop offline para la Monografia 01 del curso GE605U Sistema y Gestion Financiera.

El sistema implementa el **Modelo ABC por Ordenes de Trabajo** para Frenosa, importando datos desde un Excel maestro y generando calculos, reportes, dashboard, documentos Word y pruebas funcionales.

## Ejecutar

```bash
python -m app.main
```

## Flujo de uso

1. Abrir la aplicacion: empieza sin archivos cargados.
2. *Seleccionar* uno o varios Excel maestros (Ctrl/Shift + clic), por ejemplo `data/frenosa_datos_entrada_abc.xlsx` o los de `data/simulacion_2010/`. Cada archivo es un mes: el sistema lo identifica por su hoja `Periodo`.
3. *Validar estructura* (valida todos) e *Importar a SQLite* (importa y calcula cada mes).
4. En *Seleccion del periodo de analisis* elegir el tipo (mensual, bimestral, trimestral, cuatrimestral, semestral, anual) y el periodo. Solo aparecen los tipos posibles con los meses consecutivos cargados (enero y febrero → mensual o bimestral).
5. Seleccionar `Modelo 1 - Modelo ABC por Ordenes de Trabajo` y presionar `Ejecutar calculo`.
6. Revisar Datos importados, Proceso 1, Proceso 2 (produccion y soporte), hoja de costos, estado de resultados, dashboard y evolucion.
7. Simular: clic en una celda con lapiz (✎), nuevo valor, *Ejecutar nuevos cálculos*; *Reestablecer valores iniciales* vuelve a los datos de los Excel.
8. Generar reportes Excel y documentos Word.

## Modelos

- Modelo 1: disponible. Flujo: `EC -> AC`, `AC -> OT (Produccion)` y `AC -> OT (Soporte)`.
- Modelo 2: visible, pero bloqueado por falta de centros de actividad e inductores `EC-CA` y `CA-AC`.
- Modelo 3: visible, pero bloqueado por falta de centros estrategicos, operativos, de apoyo y relaciones escalonadas.

## Reportes generados

- `Reporte_Proceso_1_EC_AC.xlsx`
- `Reporte_Proceso_2_AC_OT.xlsx`
- `Reporte_Hoja_Costos_OT.xlsx`
- `Reporte_Estado_Resultados.xlsx`
- `Reporte_Dashboard.xlsx`

## Documentos generados

- `Manual_descriptivo.docx`
- `Manual_usuario.docx`
- `Informe_pruebas.docx`
- `Informe_propuestas_mejora.docx`
- `Analisis_funcional_diseno_sistema.docx`
- `Codigos_programacion_principales.docx`

## Objeto de costo = orden de trabajo

En este sistema el objeto de costo es la **orden de trabajo (OT)**: cada OT es un lote mensual de un producto del caso (OT-FRENOSA-SW50, OT-FRENOSA-FRE9000, OT-FRENOSA-FRE9002). En el Excel maestro `objeto_costo_id` es igual a `ot_id` y el producto queda como atributo (`codigo_producto`). Pantallas, dashboard, reportes, documentos y el agente identifican cada objeto de costo por su OT. Los cálculos no cambian.

## Kardex de materiales directos (hoja Compras) = fuente del CMD

- Ya no existe la hoja/tabla `Materiales_Directos`. El costo de material directo (CMD) de cada OT sale del **kardex por promedio ponderado** de la hoja `Compras` (obligatoria): saldo inicial, compras con IGV (19 % en 2010) y flete, devoluciones a proveedor y salidas (consumos) a la OT de la columna `destino` (codigo_ot u ot_id).
- Costo de compra = (cantidad × precio con IGV / (1 + IGV) + flete) / cantidad; salidas al costo promedio vigente; saldo inicial + compras − salidas = saldo final (`app/kardex_compras.py`). CMD de la OT = suma de sus salidas valorizadas (`app/costing_direct.py`).
- En los Excel incluidos, las salidas reproducen el CMD anterior (en julio, las cantidades y costos unitarios del Cuadro 1 del caso) con diferencias menores a S/ 0.000001: los resultados no cambian. Los objetivos usados estan en `tools/cmd_objetivo_2010.json`.
- **Donde se ve:** *Datos importados* → `compras` o `kardex_compras` (kardex valorizado por material); *Reportes* → *Generar kardex promedio (Excel)*.
- **Hoja `Ventas`:** `cantidad_vendida` y `valor_venta_unitario_soles` (= ventas netas / cantidad).

## Varios meses: analisis mensual, bimestral, trimestral, semestral, anual

- Cada Excel importado se valida, importa y calcula por separado; se guardan sus datos (valores iniciales) y resultados (`mp_periodos`, `mp_datos`, `mp_resultados`). Si se importa otra vez el mismo mes, se reemplaza. *Quitar archivos cargados* deja el sistema vacio.
- Un mes: el calculo ABC normal (con *Ver calculos*). Varios meses: las pestañas muestran el consolidado del periodo (importes sumados; costo unitario = costo / cantidad del periodo; rentabilidad = margen operativo / ventas del periodo; peso AC-OT efectivo; las OT de cada mes se agrupan por producto). *Ver calculos* muestra la fila mes por mes y *Datos importados* tiene un selector de mes.
- Pestaña *Evolucion*: mes a mes del periodo, comparativo por subperiodo del año (con los meses cargados) y kardex del periodo.
- *Generar reportes Excel* agrega `Reporte_Analisis_Periodos_<periodo>.xlsx` (hojas consolidadas con SUMIFS sobre el detalle mensual).

## Simulacion para la toma de decisiones

- Celdas modificables (el cursor muestra un lapiz ✎): valores numericos de *Datos importados* del mes visible; en *Proceso 1* PD, otros recursos y soporte de cada actividad; en *Proceso 2* el peso AC-OT; en *Hoja de costos* cantidad producida, CMD y CSD; en *Estado de resultados* ventas netas. Clic → cuadro con el nuevo valor → *Aceptar* o *Cancelar*.
- *Ejecutar nuevos cálculos* recalcula todo con los valores modificados; la franja amarilla y *Ver comparacion y cambios* muestran valores iniciales vs. nuevos (ventas, costos, margen, rentabilidad por OT) y la lista de cambios.
- *Reestablecer valores iniciales* (junto a *Volver a pantalla inicial*) descarta los cambios y recalcula.
- En un periodo de varios meses, el nuevo total se reparte entre los meses en proporcion a sus valores (los pesos se aplican igual a todos los meses). Los cambios se guardan en `sim_ajustes`; los Excel no se modifican.

## Pantallas

- Pestañas con barras de desplazamiento generales (el contenido no se aprieta), pestañas mas espaciadas y numeros con separador de miles.
- Pantalla inicial desplazable con el grafico del modelo grande, zoom, barras y *Ver en pantalla completa*.

## Simulación 2010 (enero a diciembre)

`data/simulacion_2010/` trae 12 Excel maestros (uno por mes) con la estructura del caso de julio 2010: julio con los datos reales y los demás meses con datos ficticios. Según el sistema, enero, febrero y noviembre tienen rentabilidad operativa negativa y junio, agosto, setiembre y octubre superan a julio. `Resumen_simulacion_2010.xlsx` resume los resultados. Ver `data/simulacion_2010/LEEME.md`; generador en `tools/simular_periodos_2010.py`.

## Agente IA de Finanzas y Contabilidad

Pestaña **Agente IA Finanzas** (y consola `python -m app.agente_ia`). El agente:

1. **Interpreta** preguntas en lenguaje cotidiano y las reformula con el vocabulario del Modelo ABC por Órdenes de Trabajo:
   - `¿cuál es producto con más ventas?` → `¿Cuál es el objeto de costo que genera más ingresos?`
   - `¿Qué producto generó más costos?` → `¿Cuál objeto de costo generó más costo de ventas?`
2. **Responde** leyendo los datos importados y los cálculos ya generados (fórmula, periodo, fuente y observaciones).
3. **Propone mejoras** cuantificadas (impacto estimado en margen operativo) y sugiere el área responsable.
4. Usa el **kardex por promedio ponderado** (`app/agente_ia/kardex.py`, mismas reglas de `inventory_average.py`) para el CMD, el costo promedio de salida y el saldo de inventario, y lo concilia con la tabla `kardex_promedio`.

**Solo lectura:** abre SQLite con `mode=ro` y `PRAGMA query_only`, acepta únicamente `SELECT`, no importa, no calcula, no genera archivos ni corre pruebas. Ante un pedido de cambio ("modifica", "borra", "sube el costo…") responde que no realiza cambios.

Base de conocimiento: `app/agente_ia/preguntas_base.txt` (48 preguntas en 6 categorías). Se puede editar el txt para agregar preguntas: las líneas `N. pregunta` son preguntas y las demás, categorías. Ejemplos de redacción y plantillas de "el mayor / el menor" en `app/agente_ia/catalogo.py`; sinónimos en `app/agente_ia/lexico.py`. Funciona sin internet (no usa modelos externos), por eso va dentro del `.exe`.

```bash
python -m app.agente_ia "¿Qué producto generó más costos?"
python -m app.agente_ia --propuestas
python -m unittest tests.test_agente_ia tests.test_kardex_compras tests.test_simulacion_2010 tests.test_analisis_periodos -v
```

## Pruebas

```bash
python -m app.tests_runner
```

## Compilar EXE

En Windows:

```bat
build_windows.bat
```

La compilacion genera un unico archivo ejecutable:

```text
dist\Sistema_Costos_ABC_Frenosa.exe
```

Para usarlo en otra maquina, copie solo ese `.exe`. No requiere internet, Python ni instalacion de librerias. Al ser `onefile`, puede tardar unos segundos adicionales al abrir porque Windows extrae dependencias internas en una carpeta temporal.

El sistema no contiene datos calculados hardcodeados. Los calculos se reproducen desde el Excel maestro importado y se guardan en SQLite local.
