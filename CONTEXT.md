# CONTEXT.md — Contexto académico y operativo del proyecto

## 1. Objetivo del proyecto

Este proyecto busca desarrollar un **Sistema Informático de Costos Basado en Actividades por Órdenes de Trabajo** para la Monografía 01 del curso **GE605U Sistema y Gestión Financiera**.

El aplicativo debe permitir importar un Excel maestro, procesar costos directos e indirectos, asignar costos a actividades y órdenes de trabajo, calcular costos unitarios, generar hojas de costos, dashboard, reportes Excel, documentos Word e informe de propuestas de mejora.

El nombre funcional del modelo será:

```text
Modelo ABC por Órdenes de Trabajo
```

No se debe presentar el sistema final como Modelo I, Modelo II o Modelo III. El archivo `Modelo 1 Costeo ABC Textil La Tela.xls` se usará solo como referencia metodológica interna del flujo ABC basado en actividades y objetos de costo.

---

## 2. Archivos base del proyecto

El repositorio debe contener:

```text
/proyecto
  ├── AGENTS.md
  ├── CONTEXT.md
  ├── data/
  │   ├── frenosa_datos_entrada_abc.xlsx
  │   └── Modelo 1 Costeo ABC Textil La Tela.xls
  ├── docs/
  │   ├── MONOGRADIA 01 GE605U 2026 01.pdf
  │   ├── Semana 07 GE605U 2026 01.pdf
  │   └── Frenosa_Caso.pdf
  ├── app/
  ├── reports/
  ├── exports/
  ├── tests/
  └── README.md
```

### Función de cada archivo

| Archivo | Función |
|---|---|
| `MONOGRADIA 01 GE605U 2026 01.pdf` | Instrucciones, funcionalidades, entregables y criterios de evaluación. |
| `Semana 07 GE605U 2026 01.pdf` | Teoría del Sistema ABC: elementos de costo, actividades, objetos de costo, inductores y modelos de costeo. |
| `Frenosa_Caso.pdf` | Fuente real del caso Frenosa y datos de materiales, personal directo, otros recursos, procesos, actividades, soporte, inductores, productos y ventas. |
| `frenosa_datos_entrada_abc.xlsx` | Excel maestro de entrada que importará el sistema. |
| `Modelo 1 Costeo ABC Textil La Tela.xls` | Referencia metodológica del flujo ABC basado en actividades. No es fuente de datos de Frenosa. |

---

## 3. Cambio importante del proyecto

El proyecto ya no usará:

```text
FIMAR
Modelo II
Modelo III
```

La fuente real será:

```text
Frenosa
```

El enfoque operativo será un sistema ABC por órdenes de trabajo:

```text
Recursos / elementos de costo → Actividades → Órdenes de trabajo / objetos de costo
EC → AC → OT
```

La adaptación principal es que el **material directo se asigna directamente a la OT**, pero el **personal directo se calcula primero por actividad** para respetar la lógica del caso Frenosa.

---

## 4. Requisitos de la monografía

La monografía solicita construir con IA un sistema informático que:

1. Registre multiperiodos de costeo.
2. Registre multimodelos de costeo.
3. Importe tablas maestras desde Excel.
4. Use valorización promedio de inventarios.
5. Calcule costos de producción por objeto de costo, periodo, modelo y empresa.
6. Presente hojas de costos.
7. Realice análisis de relevancia de estructura de costos.
8. Analice la data de costeo por periodo y objeto de costo.
9. Presente dashboard de costos.
10. Genere informe de propuestas de mejora con impacto positivo en rentabilidad.
11. Genere reportes Excel.
12. Genere manuales y documentos Word.
13. Genere informe de pruebas.
14. Entregue un archivo `.exe` que funcione sin internet y sin instalar software adicional.

---

## 5. Fuente real: Frenosa

Frenosa es una empresa industrial ubicada en la zona industrial del Callao, dedicada al diseño, desarrollo y producción de materiales de fricción, líquido de freno, refrigerante, remaches, herramental para instalar materiales de fricción y autopartes.

El caso de estudio costea la línea:

```text
Forros de embrague
```

Los productos principales del caso son:

```text
Forro FEF 0014 P — SW-50
Forro FEF 0014 P — FRE-9000
Forro FEF 0011 P — FRE-9002
```

Para la monografía, el objeto de costo es la orden de trabajo (OT): cada producto se fabrica en una OT y esa OT es lo que se costea (`objeto_costo_id = ot_id`):

```text
OT-FRENOSA-SW50
OT-FRENOSA-FRE9000
OT-FRENOSA-FRE9002
```

---

## 6. Elementos del sistema ABC

El sistema debe manejar los siguientes conceptos:

| Elemento | Interpretación |
|---|---|
| Empresa | Frenosa. |
| Periodo | Mes o periodo de costeo del caso. |
| Modelo de costeo | Modelo ABC por Órdenes de Trabajo. |
| Elemento de costo | Recurso económico consumido. |
| Actividad | Tarea productiva o de soporte que consume recursos. |
| Inductor | Criterio proporcional usado para repartir costos. |
| Objeto de costo | La orden de trabajo (OT) que se desea costear; cada OT fabrica un producto del caso. |
| Orden de trabajo | Lote mensual de un producto; es el objeto de costo del sistema. |

---

## 7. Actividades productivas

Actividades productivas del caso Frenosa:

| Código | Actividad |
|---|---|
| FO-1 | Impregnación |
| FO-2 | Enrollado |
| FO-3 | Prensado |
| FO-4 | Horneado |
| FO-5 | Lijado |
| FO-6 | Taladrado |
| FO-7 | Acabado |

Estas actividades reciben costos desde personal directo, otros recursos indirectos y procesos de soporte.

---

## 8. Flujo de costos del sistema

El flujo principal será:

```text
CMD → OT
MOD → AC → OT
Otros recursos indirectos → AC → OT
Mantenimiento → AC → OT
Control de calidad → AC → OT
Ingeniería y desarrollo → AC → OT
```

El sistema debe mostrar módulos separados para:

```text
MOD → AC
Otros recursos → AC
Soporte → AC
AC → OT
Hoja de costos OT
Rentabilidad
Dashboard
```

---

## 9. CMD — Costo de Material Directo

El material directo se identifica con el producto y se asigna directamente a la orden de trabajo.

Materiales del caso:

```text
Material SW-50
Material FRE-9000
Material FRE-9002
```

Fórmula general:

```text
CMD_OT = cantidad_consumida × costo_unitario_promedio
```

La monografía exige valorización promedio de inventarios; por ello, el sistema debe tener capacidad de Kardex promedio, aunque el caso Frenosa proporcione costos unitarios publicados.

---

## 10. MOD / Personal directo

El personal directo corresponde a los operarios que ejecutan actividades productivas.

En Frenosa se realiza un estudio de métodos y tiempos para calcular tiempos estándares por actividad, luego se distribuye la MOD por actividad y posteriormente hacia los productos u órdenes de trabajo.

### Datos necesarios

La hoja `Personal_directos` debe permitir calcular:

```text
tiempo_normal_min_kg
tolerancias
tiempo_suplementario_total
kg_producidos_periodo
horas_reales_trabajadas_periodo
horas_estandar_totales_periodo
tiempo_real_total_minutos_periodo
cantidad_operarios
salario_mensual_operario_soles
redondeos aplicables
```

### Fórmula recomendada

```text
Ts_total = SUM(tolerancias)
Te = ROUND(tiempo_normal_min_kg × (1 + Ts_total), redondear_te_decimales)
horas_estandar_total = horas_estandar_totales_periodo
Fd = horas_reales_trabajadas_periodo / horas_estandar_total
tiempo_real_AC = Te × kg_producidos_periodo × Fd
porcentaje_AC = tiempo_real_AC / tiempo_real_total_minutos_periodo
total_MOD = cantidad_operarios × salario_mensual_operario_soles
costo_PD_AC = porcentaje_AC × total_PD
```

Para replicar el cuadro de Frenosa, permitir usar:

```text
Fd redondeado a 2 decimales
porcentaje_AC redondeado a porcentaje entero
costo_PD_AC redondeado a soles enteros
```

---

## 11. CSD — Servicio directo

CSD representa servicios de terceros identificables directamente con una orden de trabajo.

El sistema debe tener estructura para CSD, pero en el caso Frenosa puede quedar vacío si el Excel no registra servicios directos.

Regla:

```text
Si no hay CSD publicado, CSD_OT = 0 y observación = Sin datos de servicio directo publicado.
```

No inventar servicios directos.

---

## 12. Otros recursos indirectos

El caso Frenosa agrupa otros recursos como:

```text
Personal indirecto
Servicios
Operacionales
No operacionales
Instalaciones
```

Estos recursos se asignan a las actividades productivas usando inductores como:

```text
H-H
H-M
N° kg
```

Fórmula:

```text
costo_recurso_AC = importe_recurso × porcentaje_inductor_AC
```

Resultado:

```text
costo_otros_recursos_AC = SUM(costo_recurso_AC)
```

---

## 13. Procesos de soporte como costos indirectos

Mantenimiento, control de calidad e ingeniería son **costos indirectos / procesos de soporte**. Deben calcularse por separado y luego asignarse a actividades.

### 13.1 Mantenimiento

Mantenimiento brinda servicio a equipos de planta e instalaciones.

Fórmula:

```text
costo_mantenimiento_AC = SUM(costo_mantenimiento_equipo relacionado con la actividad)
```

### 13.2 Control de calidad

Control de calidad realiza pruebas de laboratorio para controlar, verificar e inspeccionar materiales y forros de embrague.

Primero:

```text
costo_control_calidad_total = SUM(cantidad_prueba × costo_unitario_prueba)
```

Luego:

```text
costo_control_calidad_AC = costo_control_calidad_total × porcentaje_cuadro_12_AC
```

### 13.3 Ingeniería y desarrollo

Ingeniería elabora y actualiza planos para la fabricación.

Primero:

```text
costo_ingenieria_total = SUM(cantidad_plano × costo_unitario_plano)
```

Luego:

```text
costo_ingenieria_AC = costo_ingenieria_total × porcentaje_cuadro_12_AC
```

### 13.4 Soporte total por actividad

```text
costo_soporte_AC = costo_mantenimiento_AC + costo_control_calidad_AC + costo_ingenieria_AC
```

Este resultado debe calcularlo el sistema y mostrarse en un módulo de soporte hacia actividades.

---

## 14. Actividad hacia orden de trabajo

El sistema debe calcular dos rutas AC-OT:

### 14.1 Producción / costos de conversión

```text
costo_produccion_OT = SUM(costo_actividad_productiva_AC × peso_AC_OT_produccion)
```

### 14.2 Soporte

```text
costo_soporte_OT = SUM(costo_soporte_AC × peso_AC_OT_soporte)
```

### 14.3 Costo total por OT

```text
costo_total_OT = CMD_OT + costo_produccion_OT + costo_soporte_OT + CSD_OT
```

### 14.4 Costo unitario por OT

```text
costo_unitario_OT = costo_total_OT / cantidad_producida_OT
```

---

## 15. Hoja de costos por OT

La hoja de costos debe mostrar:

```text
CMD
+ MOD asignada vía actividades
+ CSD
+ Otros recursos indirectos asignados vía actividades
+ Soporte asignado vía actividades
= Costo total OT
```

También debe mostrar:

```text
Costo unitario = Costo total OT / cantidad producida
```

Si existe valor de venta:

```text
Margen = Valor de venta - Costo total
Rentabilidad % = Margen / Valor de venta
```

---

## 16. Estado de resultados por producto u OT

El sistema debe poder construir un estado de resultados por objeto de costo, tomando como referencia el caso Frenosa:

```text
Ventas netas
(-) Materia prima directa
(-) Costos de conversión
= Margen bruto
(-) Procesos de soporte
= Margen operativo
```

En el sistema, los productos se mostrarán como órdenes de trabajo equivalentes.

---

## 17. Dashboard esperado

El dashboard debe incluir:

- Costo total por OT.
- Costo unitario por OT.
- Participación de CMD, MOD, CSD y costos indirectos.
- Costo por actividad.
- Costo por recurso indirecto.
- Costo por proceso de soporte.
- Ranking de actividades más costosas.
- Comparación de productos/OT.
- Margen bruto y margen operativo por OT.
- Indicadores de mejora.

---

## 18. Informe de propuestas de mejora

El informe debe usar resultados calculados por el sistema para proponer mejoras como:

- Revisar actividades con mayor consumo de recursos.
- Optimizar actividades con mayor costo de MOD.
- Revisar mantenimiento de equipos con mayor peso.
- Mejorar precisión de inductores.
- Separar actividades que agregan valor y que no agregan valor.
- Evaluar productos con menor margen operativo.
- Mejorar el uso de información ABC para decisiones comerciales.

---

## 19. Reglas de trazabilidad

Cada dato importado debe conservar:

- Empresa.
- Periodo.
- Modelo de costeo.
- Fuente.
- Hoja de origen.
- Tipo de dato.
- Observación.

Fuentes esperadas:

```text
Frenosa_Caso.pdf
Semana 07 GE605U 2026 01.pdf
MONOGRADIA 01 GE605U 2026 01.pdf
Modelo 1 Costeo ABC Textil La Tela.xls
frenosa_datos_entrada_abc.xlsx
```

No inventar datos ni reemplazar datos faltantes con supuestos sin marcarlos explícitamente.

---

## 20. Valorización promedio

La monografía exige valorización promedio para inventarios.

Fórmula general:

```text
Costo promedio = costo total disponible / cantidad total disponible
```

Para salidas:

```text
Costo de salida = cantidad salida × costo promedio vigente
```

El sistema debe permitir Kardex promedio para materiales, aunque el caso Frenosa proporcione cantidades y costos unitarios directamente.

---

## 21. Reglas generales de costos

1. Trabajar costos sin IGV cuando corresponda.
2. Separar costos de producción de gastos administrativos, ventas o financieros.
3. No mezclar datos fuente con resultados calculados.
4. No usar datos de Textil La Tela como si fueran de Frenosa.
5. No usar FIMAR.
6. No usar Modelo II como enfoque del proyecto.
7. No exigir hojas de resultado como entrada.
8. Todo cálculo debe poder reproducirse desde Excel.

---

## 22. Recomendaciones de implementación

Crear módulos separados:

```text
importers.py
database.py
inventory_average.py
costing_direct.py
costing_mod.py
costing_abc.py
reports.py
dashboard.py
documents.py
tests_runner.py
```

Cada función debe recibir datos desde SQLite o estructuras importadas del Excel, ejecutar el cálculo y devolver resultado con trazabilidad.

---

## 23. Estilo esperado de salidas

Cada pantalla, reporte o documento debe mostrar:

1. Empresa.
2. Periodo.
3. Modelo de costeo.
4. Orden de trabajo.
5. Fórmula usada.
6. Datos de entrada.
7. Resultado.
8. Fuente.
9. Observación.

El sistema debe diferenciar claramente:

```text
Dato importado
Cálculo generado
Reporte exportado
```

---

## 24. Entregables finales esperados

La entrega final debe incluir:

1. Word de análisis funcional, diseño de base de datos, diseño del sistema y casos de uso.
2. Word con códigos de programación.
3. Excel maestro Frenosa.
4. Manual descriptivo.
5. Manual de usuario.
6. Ejecutable `.exe` offline.
7. Reportes Excel generados.
8. Informe de pruebas.
9. Informe de propuestas de mejora.
10. ZIP o RAR final.
