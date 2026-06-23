# AGENTS.md — Instrucciones para Codex

## 1. Rol del proyecto

Este repositorio se usará para construir el aplicativo de la **Monografía 01 del curso GE605U Sistema y Gestión Financiera**.

El objetivo principal es desarrollar un:

**Sistema Informático de Costos Basado en Actividades por Órdenes de Trabajo, usando información real de Frenosa y una metodología ABC por órdenes de trabajo.**

El sistema debe importar información desde Excel, calcular costos directos e indirectos, asignar costos a actividades y órdenes de trabajo, generar hojas de costos, reportes, dashboard, informes Word, pruebas y un ejecutable `.exe` offline.

---

## 2. Cambio obligatorio del proyecto

Este proyecto **ya no usa FIMAR, ni Modelo II, ni la empresa anterior**.

La empresa real será:

```text
Frenosa
```

El modelo debe denominarse en pantallas, reportes y documentos como:

```text
Modelo ABC por Órdenes de Trabajo
```

No presentar el aplicativo final como “Modelo I”, “Modelo II” ni “Modelo III”.

Nota técnica interna: el archivo `Modelo 1 Costeo ABC Textil La Tela.xls` se usará como **referencia metodológica del flujo ABC tipo Modelo I**, pero el producto final debe llamarse **Modelo ABC por Órdenes de Trabajo** para ajustarse al enfoque solicitado.

---

## 3. Archivos obligatorios del repositorio

Antes de implementar o modificar código, Codex debe revisar estos archivos:

```text
AGENTS.md
CONTEXT.md
```

En la carpeta `docs/`:

```text
docs/MONOGRADIA 01 GE605U 2026 01.pdf
docs/Semana 07 GE605U 2026 01.pdf
docs/Frenosa_Caso.pdf
```

En la carpeta `data/`:

```text
data/frenosa_datos_entrada_abc.xlsx
data/Modelo 1 Costeo ABC Textil La Tela.xls
```

### Uso de cada archivo

- `MONOGRADIA 01 GE605U 2026 01.pdf`: instrucciones obligatorias, entregables y criterios de evaluación de la monografía.
- `Semana 07 GE605U 2026 01.pdf`: teoría del Sistema de Costeo Basado en Actividades, elementos ABC, inductores, objetos de costo y modelos de costeo.
- `Frenosa_Caso.pdf`: fuente real de Frenosa y datos principales de materiales, personal directo, otros recursos, procesos de producción, procesos de soporte, actividades, inductores, productos, costos y ventas.
- `frenosa_datos_entrada_abc.xlsx`: Excel maestro de entrada que debe importar el sistema. Este archivo contiene datos fuente, no reportes finales.
- `Modelo 1 Costeo ABC Textil La Tela.xls`: archivo de referencia metodológica para entender la estructura de trabajo ABC con actividades y objetos de costo. No debe usarse como fuente de datos reales de Frenosa.

---

## 4. Restricción principal

El aplicativo **no debe tener datos precargados dentro del código**.

Al abrir el `.exe`, el usuario debe poder:

1. Seleccionar un Excel maestro desde cualquier ruta.
2. Importar el archivo.
3. Validar hojas y columnas.
4. Crear o actualizar una base SQLite local.
5. Ejecutar los cálculos del Modelo ABC por Órdenes de Trabajo.
6. Generar reportes Excel, documentos Word y dashboard.

El archivo principal de prueba será:

```text
data/frenosa_datos_entrada_abc.xlsx
```

El sistema debe permitir importar otros Excel en el futuro si respetan la estructura esperada.

---

## 5. Requisitos obligatorios de la monografía

El sistema debe cumplir con lo siguiente:

1. Registrar múltiples periodos de costeo.
2. Registrar múltiples modelos de costeo.
3. Importar información para tablas maestras desde archivos Excel.
4. Usar valorización promedio para inventarios.
5. Calcular costos de producción por objeto de costo, periodo, modelo y empresa.
6. Presentar hojas de costos por objeto de costo, periodo, modelo y empresa.
7. Realizar análisis de relevancia de la estructura de costos.
8. Analizar la data de costeo por periodo y objeto de costo.
9. Presentar dashboard de costos.
10. Elaborar informe de propuestas de mejora con impacto positivo en la rentabilidad.
11. Generar reportes Excel.
12. Generar documentos Word.
13. Generar informe de pruebas.
14. Compilar un archivo ejecutable `.exe` que funcione sin internet y sin instalar software adicional.

---

## 6. Regla del ejecutable

La entrega final debe incluir obligatoriamente:

```text
Sistema_Costos_ABC_Frenosa.exe
```

El `.exe` debe:

- funcionar sin internet,
- no requerir instalación de Python,
- no requerir instalación de librerías,
- no requerir instalación de software adicional,
- permitir importar manualmente el Excel maestro,
- usar SQLite local,
- generar reportes Excel,
- generar documentos Word,
- mostrar dashboard de costos.

No aceptar como solución final “no se pudo compilar”. El proyecto debe dejar el `.exe` generado o un proceso automatizado de compilación para Windows.

---

## 7. Enfoque ABC que debe implementar el sistema

El sistema trabajará con la lógica:

```text
Elementos de costo y recursos → Actividades → Órdenes de trabajo / Objetos de costo
EC → AC → OT
```

La planta de producción se interpreta como el ámbito principal de costeo. Las actividades productivas y de soporte consumen recursos, y luego sus costos se asignan a las órdenes de trabajo mediante inductores.

### Regla importante para Frenosa

En el caso Frenosa:

```text
CMD → OT directamente
Personal directo / MOD → Actividad → OT
Otros recursos indirectos → Actividad → OT
Mantenimiento → Actividad → OT
Control de calidad → Actividad → OT
Ingeniería y desarrollo → Actividad → OT
```

Aunque el personal directo es un costo directo en la clasificación contable, en este caso debe asignarse primero a las actividades usando tiempos estándares, tiempo suplementario y factor de distribución, porque Frenosa calcula la MOD por actividad antes de transferirla a los productos u órdenes de trabajo.

---

## 8. Empresa real: Frenosa

Frenosa es una empresa industrial ubicada en la zona industrial del Callao, dedicada al diseño, desarrollo y producción de materiales de fricción, líquido de freno, refrigerante, remaches, herramental para instalar materiales de fricción y autopartes.

La línea principal usada para el sistema será:

```text
Forros de embrague
```

Los productos reales del caso deben registrarse como objetos de costo y también como órdenes de trabajo equivalentes para la monografía.

Ejemplo de codificación operativa:

```text
OT-FRENOSA-SW50
OT-FRENOSA-FRE9000
OT-FRENOSA-FRE9002
```

Esta codificación no cambia los datos reales; solo adapta el caso a un sistema por órdenes de trabajo.

---

## 9. Actividades productivas de Frenosa

Las actividades productivas principales son:

```text
FO-1 Impregnación
FO-2 Enrollado
FO-3 Prensado
FO-4 Horneado
FO-5 Lijado
FO-6 Taladrado
FO-7 Acabado
```

Estas actividades reciben costos desde:

- personal directo / MOD,
- otros recursos indirectos,
- mantenimiento,
- control de calidad,
- ingeniería y desarrollo.

Luego se asignan a las órdenes de trabajo mediante inductores AC-OT.

---

## 10. Clasificación de costos en el sistema

### 10.1 CMD — Costo de Material Directo

Material identificable directamente con la orden de trabajo.

En Frenosa:

```text
Material SW-50
Material FRE-9000
Material FRE-9002
```

Regla:

```text
CMD por OT = cantidad consumida × costo unitario promedio
```

El CMD debe ir directamente a la OT.

---

### 10.2 CPD / MOD — Personal Directo

Personal que ejecuta las actividades productivas.

En Frenosa, no asignar la MOD directamente a la OT. Debe calcularse primero por actividad:

```text
MOD → AC → OT
```

Cálculo base:

```text
Ts_total = SUM(tolerancias)
Te = ROUND(tiempo_normal_min_kg × (1 + Ts_total), redondear_te_decimales)
horas_estandar_total = horas_estandar_totales_periodo
o, si no viene dado: SUM(Te × kg_producidos_periodo) / 60
Fd = horas_reales_trabajadas_periodo / horas_estandar_total
tiempo_real_AC = Te × kg_producidos_periodo × Fd
porcentaje_AC = tiempo_real_AC / SUM(tiempo_real_AC)
total_MOD = cantidad_operarios × salario_mensual_operario_soles
costo_PD_AC = porcentaje_AC × total_PD
```

Para replicar los valores del caso Frenosa, permitir configuración de redondeo:

```text
Fd redondeado a 2 decimales
porcentaje de actividad redondeado a porcentaje entero
costo MOD redondeado a soles enteros
```

---

### 10.3 CSD — Costo de Servicio Directo

Servicio de tercero directamente aplicado a una OT.

El sistema debe tener estructura para CSD, pero si el Excel no contiene servicios directos, debe registrar CSD = 0 con observación.

No inventar servicios directos.

---

### 10.4 CI / CIF — Costos indirectos

En Frenosa son costos indirectos:

```text
Personal indirecto
Servicios
Operacionales
No operacionales
Instalaciones
Mantenimiento
Control de calidad
Ingeniería y desarrollo
```

También deben tratarse como indirectos los procesos de soporte:

```text
Mantenimiento
Control de calidad
Ingeniería y desarrollo
```

Estos costos se calculan aparte y se asignan a actividades antes de pasar a las órdenes de trabajo.

---

## 11. Cálculo de otros recursos indirectos hacia actividades

Los otros recursos indirectos del proceso productivo se asignan a actividades usando inductores EC-AC.

Fórmula general:

```text
costo_recurso_actividad = importe_recurso × porcentaje_inductor_recurso_actividad
```

Luego:

```text
costo_otros_recursos_AC = SUM(costo_recurso_actividad)
```

El sistema debe mostrar este cálculo en un módulo de asignación EC-AC.

---

## 12. Cálculo de soporte hacia actividades

Los procesos de soporte deben calcularse separadamente de los otros recursos.

### 12.1 Mantenimiento

Mantenimiento se asigna a la actividad productiva según el equipo relacionado.

Fórmula:

```text
costo_mantenimiento_AC = SUM(costo_equipo_mantenimiento asociado a la actividad)
```

Ejemplo conceptual:

```text
Tina de impregnación → Impregnación
Tanque mezclador → Impregnación
Prensa hidráulica → Prensado
Lijadora de banda → Lijado
```

### 12.2 Control de calidad

Primero calcular el costo total del proceso de control de calidad:

```text
costo_control_calidad_total = SUM(cantidad_prueba × costo_unitario_prueba)
```

Luego asignarlo a actividades usando los porcentajes del cuadro 12:

```text
costo_control_calidad_AC = costo_control_calidad_total × porcentaje_cuadro_12_AC
```

### 12.3 Ingeniería y desarrollo

Primero calcular el costo total del proceso de ingeniería:

```text
costo_ingenieria_total = SUM(cantidad_plano × costo_unitario_plano)
```

Luego asignarlo a actividades usando los porcentajes del cuadro 12:

```text
costo_ingenieria_AC = costo_ingenieria_total × porcentaje_cuadro_12_AC
```

### 12.4 Costo total de soporte por actividad

```text
costo_soporte_AC = costo_mantenimiento_AC + costo_control_calidad_AC + costo_ingenieria_AC
```

Este resultado no debe estar precargado; debe ser generado por el sistema.

---

## 13. Cálculo de actividad hacia orden de trabajo

El sistema debe calcular dos asignaciones AC-OT:

1. Actividades productivas / costos de conversión.
2. Actividades de soporte transferidas.

### 13.1 Producción AC-OT

```text
costo_produccion_OT = SUM(costo_actividad_productiva_AC × peso_AC_OT_produccion)
```

### 13.2 Soporte AC-OT

```text
costo_soporte_OT = SUM(costo_soporte_AC × peso_AC_OT_soporte)
```

### 13.3 Costo total por OT

```text
costo_total_OT = CMD_OT + costo_produccion_OT + costo_soporte_OT + CSD_OT
```

Si no existe CSD:

```text
CSD_OT = 0
```

### 13.4 Costo unitario

```text
costo_unitario_OT = costo_total_OT / cantidad_producida_OT
```

---

## 14. Módulos mínimos del aplicativo

1. Pantalla inicial
   - Importar Excel maestro.
   - Validar estructura.
   - Crear base SQLite.
   - Ingresar al sistema.

2. Menú principal
   - Empresa.
   - Periodos.
   - Modelos de costeo.
   - Órdenes de trabajo.
   - Objetos de costo.
   - Materiales directos.
   - Personal directo / MOD.
   - Servicios directos.
   - Costos indirectos.
   - Actividades.
   - Inductores EC-AC.
   - Inductores AC-OT.
   - Proceso EC → AC.
   - Proceso Soporte → AC.
   - Proceso AC → OT.
   - Hoja de costos OT.
   - Resumen de costos.
   - Estado de resultados.
   - Análisis de rentabilidad.
   - Dashboard.
   - Reportes Excel.
   - Documentos Word.
   - Informe de pruebas.

3. Módulo de importación Excel
   - Leer `frenosa_datos_entrada_abc.xlsx`.
   - Validar hojas obligatorias.
   - Validar columnas.
   - Mostrar errores de importación.
   - Cargar datos a SQLite.

4. Módulo de cálculo MOD → AC
   - Calcular tiempo estándar.
   - Calcular factor de distribución.
   - Calcular tiempo real por actividad.
   - Calcular porcentaje por actividad.
   - Calcular costo MOD por actividad.

5. Módulo EC → AC
   - Asignar otros recursos indirectos a actividades.
   - Mostrar fórmula, datos de entrada y resultados.

6. Módulo Soporte → AC
   - Calcular mantenimiento por actividad.
   - Calcular control de calidad por actividad.
   - Calcular ingeniería por actividad.
   - Calcular soporte total por actividad.

7. Módulo AC → OT
   - Asignar actividades productivas a OT.
   - Asignar soporte a OT.

8. Módulo de hojas de costos
   - Por OT.
   - Por producto/objeto de costo.
   - Por periodo.
   - Por modelo.
   - Por empresa.

9. Módulo dashboard
   - Costos por OT.
   - Costo unitario por OT.
   - Participación CMD, MOD, CSD y CIF.
   - Costos por actividad.
   - Costos por proceso de soporte.
   - Ranking de actividades con mayor costo.
   - Rentabilidad por OT.

10. Módulo de documentos
   - Manual descriptivo.
   - Manual de usuario.
   - Informe de pruebas.
   - Informe de propuestas de mejora.

---

## 15. Hojas esperadas del Excel maestro

El sistema debe validar, como mínimo, las hojas del Excel maestro `frenosa_datos_entrada_abc.xlsx` que existan para este flujo:

```text
Empresa
Periodos
Modelos_Costeo
Objetos_Costo
Ordenes_Trabajo
Materiales_Directos
Personal_directos
Servicios_Directos
costo_otros_recursos(CI)
Actividades
EC_AC_Inductores
Soporte_Mantenimiento(CI)
Soporte_Control_Calidad(CI)
Soporte_Ingenieria(CI)
Soporte_CC_ING_Inductores
AC_OT_Pesos_Produccion
AC_OT_Pesos_Soporte
Ventas
Reglas_Calculo
Fuentes
Diccionario
```

No exigir hojas de salida como requisito de importación.

No exigir:

```text
Proceso_1
Proceso_2
Proceso_3
Hoja_Costos_OT
Resumen_Costos
Dashboard
```

Estas deben ser generadas por el aplicativo.

---

## 16. Estructura esperada del repositorio

```text
/proyecto
  ├── AGENTS.md
  ├── CONTEXT.md
  ├── app/
  │   ├── main.py
  │   ├── database.py
  │   ├── importers.py
  │   ├── inventory_average.py
  │   ├── costing_direct.py
  │   ├── costing_mod.py
  │   ├── costing_abc.py
  │   ├── reports.py
  │   ├── dashboard.py
  │   ├── documents.py
  │   ├── tests_runner.py
  │   └── ui/
  ├── data/
  │   ├── frenosa_datos_entrada_abc.xlsx
  │   └── Modelo 1 Costeo ABC Textil La Tela.xls
  ├── docs/
  │   ├── MONOGRADIA 01 GE605U 2026 01.pdf
  │   ├── Semana 07 GE605U 2026 01.pdf
  │   └── Frenosa_Caso.pdf
  ├── exports/
  ├── reports/
  ├── tests/
  ├── build_exe.py
  ├── build_windows.bat
  ├── requirements.txt
  └── README.md
```

---

## 17. Estilo esperado del sistema

Las pantallas, reportes y documentos deben usar lenguaje claro y académico. Cada cálculo debe mostrar:

1. Fórmula usada.
2. Datos de entrada.
3. Resultado.
4. Periodo.
5. Fuente.
6. Observaciones.

El sistema debe diferenciar claramente:

```text
Datos importados
Cálculos generados
Reportes exportados
```

No mezclar datos fuente con resultados calculados.

---

## 18. Reglas finales para Codex

- No volver a FIMAR.
- No usar Modelo II como enfoque del proyecto.
- No usar el archivo de Textil La Tela como datos reales.
- No hardcodear los datos de Frenosa en el código.
- No poner resultados calculados como si fueran datos de entrada.
- Todo cálculo debe ser reproducible desde el Excel maestro.
- Todo reporte debe poder exportarse.
- Todo documento Word debe generarse desde el aplicativo.
- El sistema debe quedar listo para ejecutar offline en Windows.
