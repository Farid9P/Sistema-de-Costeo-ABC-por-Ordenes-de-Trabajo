# Simulación 2010 (enero a diciembre)

Doce Excel maestros con la misma estructura que `data/frenosa_datos_entrada_abc.xlsx` (julio 2010), uno por mes:
`frenosa_datos_entrada_abc_2010_01.xlsx` … `frenosa_datos_entrada_abc_2010_12.xlsx`.

- **Julio** conserva los datos reales del caso (mismos resultados que el Excel maestro); solo cambian los códigos del periodo y de las OT.
- **Los demás meses son datos ficticios**: producción por OT, costo del material, operarios, horas, otros recursos, mantenimiento,
  pruebas de calidad, planos, pesos AC-OT (según la mezcla del mes), cantidades y precios de venta. Cada archivo explica su
  escenario en la hoja README y en la hoja Periodo.
- **Rentabilidad operativa calculada por el sistema**: enero −5.95 %, febrero −14.01 %, marzo 13.00 %, abril 21.97 %,
  mayo 24.96 %, junio 30.00 %, julio 27.96 % (caso real), agosto 32.00 %, setiembre 36.00 %, octubre 31.01 %,
  noviembre −5.04 %, diciembre 16.06 %. Año: 22.64 %.
- **Kardex de la hoja Compras** (materiales directos SW-50, FRE-9000 y FRE-9002, almacén de materias primas): continuo durante el año (el saldo final de cada mes es el saldo inicial del siguiente); sus salidas a cada OT son el costo de material directo (CMD).
- `Resumen_simulacion_2010.xlsx`: resultados de los 12 meses, rentabilidad por OT y saldos del kardex.

Uso: en la aplicación, *Seleccionar* → elegir uno o varios meses (Ctrl/Shift + clic) → *Validar estructura* → *Importar a SQLite*
→ elegir el tipo de análisis (mensual, bimestral, trimestral, cuatrimestral, semestral, anual) y el periodo → *Ejecutar cálculo*.

Regenerar: `python tools/simular_periodos_2010.py --indices tools/indices_precio_2010.json` (luego abrir y guardar en Excel,
o recalcular, para que la fórmula del valor de venta unitario tenga su valor) y `python tools/resumen_simulacion_2010.py`.

Ya no hay hoja `Materiales_Directos`: el CMD de cada OT sale de las salidas del kardex (hoja `Compras`). Los objetivos de CMD usados para generar el kardex están en `tools/cmd_objetivo_2010.json` (diferencia < S/ 0.000001).
