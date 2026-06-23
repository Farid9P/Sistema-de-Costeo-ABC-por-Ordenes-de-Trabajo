# Sistema Informatico de Costos ABC - Frenosa

Aplicativo desktop offline para la Monografia 01 del curso GE605U Sistema y Gestion Financiera.

El sistema implementa el **Modelo ABC por Ordenes de Trabajo** para Frenosa, importando datos desde un Excel maestro y generando calculos, reportes, dashboard, documentos Word y pruebas funcionales.

## Ejecutar

```bash
python -m app.main
```

## Flujo de uso

1. Abrir la pantalla inicial.
2. Seleccionar `data/frenosa_datos_entrada_abc.xlsx` o un Excel compatible.
3. Presionar `Validar estructura`.
4. Presionar `Importar a SQLite`.
5. Seleccionar `Modelo 1 - Modelo ABC por Ordenes de Trabajo`.
6. Presionar `Ejecutar calculo`.
7. Revisar Proceso 1, Proceso 2 AC-OT (Produccion), Proceso 2 AC-OT (Soporte), hoja de costos, estado de resultados y dashboard.
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

## Pruebas

```bash
python -m app.tests_runner
```

## Compilar EXE

En Windows:

```bat
build_windows.bat
```

El ejecutable final esperado queda en:

```text
dist\Sistema_Costos_ABC_Frenosa.exe
```

El sistema no contiene datos calculados hardcodeados. Los calculos se reproducen desde el Excel maestro importado y se guardan en SQLite local.
