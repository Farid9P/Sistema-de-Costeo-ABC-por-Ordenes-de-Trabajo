PRAGMA foreign_keys = OFF;

CREATE TABLE IF NOT EXISTS import_batches (
    batch_id INTEGER PRIMARY KEY AUTOINCREMENT,
    excel_path TEXT NOT NULL,
    imported_at TEXT NOT NULL,
    status TEXT NOT NULL,
    notes TEXT
);

CREATE TABLE IF NOT EXISTS import_errors (
    error_id INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_id INTEGER,
    severity TEXT NOT NULL,
    sheet_name TEXT,
    row_number INTEGER,
    column_name TEXT,
    message TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS source_trace (
    trace_id INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_id INTEGER,
    table_name TEXT NOT NULL,
    sheet_name TEXT NOT NULL,
    source_row INTEGER,
    business_key TEXT,
    source_text TEXT,
    observation TEXT
);

CREATE TABLE IF NOT EXISTS calculation_runs (
    run_id INTEGER PRIMARY KEY AUTOINCREMENT,
    model_code TEXT NOT NULL,
    started_at TEXT NOT NULL,
    finished_at TEXT,
    status TEXT NOT NULL,
    notes TEXT
);

CREATE TABLE IF NOT EXISTS empresas (
    empresa_id TEXT, nombre_empresa TEXT, pais TEXT, departamento TEXT,
    sector TEXT, ciiu TEXT, actividad_principal TEXT, fuente TEXT,
    observacion TEXT, batch_id INTEGER
);

CREATE TABLE IF NOT EXISTS periodos (
    periodo_id TEXT, empresa_id TEXT, periodo TEXT, tipo_periodo TEXT,
    descripcion TEXT, fuente TEXT, batch_id INTEGER
);

CREATE TABLE IF NOT EXISTS modelos_costeo (
    modelo_id TEXT, nombre_modelo TEXT, ruta_modelo TEXT, descripcion TEXT,
    fuente TEXT, batch_id INTEGER
);

CREATE TABLE IF NOT EXISTS objetos_costo (
    objeto_costo_id TEXT, empresa_id TEXT, periodo_id TEXT, modelo_id TEXT,
    codigo_producto TEXT, descripcion_producto TEXT, peso_kg REAL,
    cantidad_piezas REAL, material_base TEXT, fuente_cuadro TEXT,
    fuente TEXT, batch_id INTEGER
);

CREATE TABLE IF NOT EXISTS ordenes_trabajo (
    ot_id TEXT, empresa_id TEXT, periodo_id TEXT, modelo_id TEXT,
    objeto_costo_id TEXT, codigo_ot TEXT, descripcion_ot TEXT,
    cantidad_producida REAL, unidad TEXT, fuente TEXT, observacion TEXT,
    batch_id INTEGER
);

CREATE TABLE IF NOT EXISTS materiales_directos (
    material_id TEXT, ot_id TEXT, objeto_costo_id TEXT, material TEXT,
    cantidad_kg REAL, costo_unitario_soles REAL, clasificacion TEXT,
    tratamiento TEXT, fuente_cuadro TEXT, fuente TEXT, observacion TEXT,
    batch_id INTEGER
);

CREATE TABLE IF NOT EXISTS "Personal_directos" (
    mod_input_id TEXT, empresa_id TEXT, periodo_id TEXT, modelo_id TEXT,
    actividad_id TEXT, codigo_actividad TEXT, nombre_actividad TEXT,
    kg_producidos_periodo REAL, horas_reales_trabajadas_periodo REAL,
    cantidad_operarios REAL, salario_mensual_operario_soles REAL,
    tiempo_normal_min_kg REAL, tolerancia_personal_pct REAL,
    tolerancia_fatiga_pct REAL, tolerancia_pie_pct REAL,
    tolerancia_encorvado_pct REAL, redondear_te_decimales INTEGER,
    redondear_porcentaje_decimales INTEGER, usar_porcentaje_redondeado_pdf TEXT,
    fuente_cuadro TEXT, fuente TEXT, observacion TEXT,
    horas_estandar_totales_periodo REAL, tiempo_real_total_minutos_periodo REAL,
    batch_id INTEGER
);

CREATE TABLE IF NOT EXISTS servicios_directos (
    servicio_id TEXT, ot_id TEXT, servicio_directo TEXT, cantidad REAL,
    costo_unitario_soles REAL, fuente TEXT, observacion TEXT, batch_id INTEGER
);

CREATE TABLE IF NOT EXISTS "costo_otros_recursos(CI)" (
    elemento_costo_id TEXT, codigo_recurso TEXT, nombre_recurso TEXT,
    importe_soles REAL, clasificacion TEXT, tipo TEXT, inductor_recurso TEXT,
    fuente_cuadro TEXT, fuente TEXT, observacion TEXT, batch_id INTEGER
);

CREATE TABLE IF NOT EXISTS actividades (
    actividad_id TEXT, codigo_actividad TEXT, nombre_actividad TEXT,
    proceso TEXT, tiempo_normal_min_kg REAL, clasificacion TEXT,
    fuente_cuadro TEXT, fuente TEXT, observacion TEXT, batch_id INTEGER
);

CREATE TABLE IF NOT EXISTS inductores (
    inductor_id TEXT, nombre_inductor TEXT, descripcion TEXT, unidad TEXT,
    fuente_cuadro TEXT, fuente TEXT, batch_id INTEGER
);

CREATE TABLE IF NOT EXISTS ec_ac_inductores (
    elemento_costo_id TEXT, codigo_recurso TEXT, actividad_id TEXT,
    codigo_actividad TEXT, nombre_actividad TEXT, inductor_recurso TEXT,
    porcentaje_inductor REAL, fuente_cuadro TEXT, fuente TEXT,
    observacion TEXT, batch_id INTEGER
);

CREATE TABLE IF NOT EXISTS "Soporte_Mantenimiento(CI)" (
    soporte_item_id TEXT, proceso_soporte TEXT, equipo_codigo TEXT,
    equipo_descripcion TEXT, costo_equipo_soles REAL, actividad_id_destino TEXT,
    codigo_actividad_destino TEXT, fuente_cuadro TEXT, fuente TEXT,
    observacion TEXT, batch_id INTEGER
);

CREATE TABLE IF NOT EXISTS "Soporte_Control_Calidad(CI)" (
    soporte_item_id TEXT, proceso_soporte TEXT, prueba_laboratorio TEXT,
    cantidad REAL, costo_unitario_soles REAL, fuente_cuadro TEXT,
    fuente TEXT, observacion TEXT, batch_id INTEGER
);

CREATE TABLE IF NOT EXISTS "Soporte_Ingenieria(CI)" (
    soporte_item_id TEXT, proceso_soporte TEXT, tipo_plano TEXT,
    cantidad REAL, costo_unitario_soles REAL, fuente_cuadro TEXT,
    fuente TEXT, observacion TEXT, batch_id INTEGER
);

CREATE TABLE IF NOT EXISTS soporte_cc_ing_inductores (
    proceso_soporte TEXT, actividad_id TEXT, codigo_actividad TEXT,
    nombre_actividad TEXT, porcentaje_distribucion REAL, base_distribucion TEXT,
    fuente_cuadro TEXT, fuente TEXT, observacion TEXT, batch_id INTEGER
);

CREATE TABLE IF NOT EXISTS ac_ot_pesos_produccion (
    actividad_id TEXT, codigo_actividad TEXT, nombre_actividad TEXT,
    ot_id TEXT, objeto_costo_id TEXT, codigo_producto TEXT,
    porcentaje_inductor REAL, tipo_costo TEXT, fuente_cuadro TEXT,
    fuente TEXT, observacion TEXT, batch_id INTEGER
);

CREATE TABLE IF NOT EXISTS ac_ot_pesos_soporte (
    actividad_id TEXT, codigo_actividad TEXT, nombre_actividad TEXT,
    ot_id TEXT, objeto_costo_id TEXT, codigo_producto TEXT,
    porcentaje_inductor REAL, tipo_costo TEXT, fuente_cuadro TEXT,
    fuente TEXT, observacion TEXT, batch_id INTEGER
);

CREATE TABLE IF NOT EXISTS ventas (
    venta_id TEXT, ot_id TEXT, objeto_costo_id TEXT, codigo_producto TEXT,
    ventas_netas_soles REAL, fuente_cuadro TEXT, fuente TEXT,
    observacion TEXT, batch_id INTEGER
);

CREATE TABLE IF NOT EXISTS reglas_calculo (
    regla_id TEXT, modulo TEXT, regla_calculo TEXT, salida_esperada TEXT,
    observacion TEXT, batch_id INTEGER
);

CREATE TABLE IF NOT EXISTS fuentes (
    fuente_id TEXT, documento TEXT, cuadro_o_seccion TEXT, dato_usado TEXT,
    url TEXT, observacion TEXT, batch_id INTEGER
);

CREATE TABLE IF NOT EXISTS diccionario (
    hoja TEXT, tipo TEXT, descripcion TEXT, calculos_no_incluidos TEXT,
    batch_id INTEGER
);

CREATE TABLE IF NOT EXISTS kardex_promedio (
    kardex_id INTEGER PRIMARY KEY AUTOINCREMENT,
    run_id INTEGER, material_id TEXT, ot_id TEXT, material TEXT,
    entrada_cantidad REAL, entrada_costo_unitario REAL, costo_promedio REAL,
    salida_cantidad REAL, costo_salida REAL, formula TEXT, observacion TEXT
);

CREATE TABLE IF NOT EXISTS cmd_ot (
    run_id INTEGER, ot_id TEXT, objeto_costo_id TEXT, material_id TEXT,
    material TEXT, cantidad_kg REAL, costo_unitario_promedio REAL,
    cmd_soles REAL, formula TEXT, fuente TEXT, observacion TEXT
);

CREATE TABLE IF NOT EXISTS csd_ot (
    run_id INTEGER, ot_id TEXT, servicio_id TEXT, servicio_directo TEXT,
    cantidad REAL, costo_unitario_soles REAL, csd_soles REAL,
    formula TEXT, fuente TEXT, observacion TEXT
);

CREATE TABLE IF NOT EXISTS mod_actividad (
    run_id INTEGER, actividad_id TEXT, codigo_actividad TEXT,
    nombre_actividad TEXT, ts_total REAL, te REAL, fd REAL,
    tiempo_real_minutos REAL, porcentaje_calculado REAL, porcentaje_usado REAL,
    total_mod_soles REAL, costo_mod_ac REAL, formula TEXT, fuente TEXT,
    observacion TEXT
);

CREATE TABLE IF NOT EXISTS otros_recursos_ac_detalle (
    run_id INTEGER, elemento_costo_id TEXT, codigo_recurso TEXT,
    nombre_recurso TEXT, actividad_id TEXT, codigo_actividad TEXT,
    nombre_actividad TEXT, importe_recurso_soles REAL, porcentaje_inductor REAL,
    costo_recurso_ac REAL, formula TEXT, fuente TEXT, observacion TEXT
);

CREATE TABLE IF NOT EXISTS otros_recursos_actividad (
    run_id INTEGER, actividad_id TEXT, codigo_actividad TEXT,
    nombre_actividad TEXT, costo_otros_recursos_ac REAL, formula TEXT,
    fuente TEXT, observacion TEXT
);

CREATE TABLE IF NOT EXISTS soporte_mantenimiento_ac (
    run_id INTEGER, actividad_id TEXT, codigo_actividad TEXT,
    costo_mantenimiento_ac REAL, formula TEXT, fuente TEXT, observacion TEXT
);

CREATE TABLE IF NOT EXISTS soporte_control_calidad_ac (
    run_id INTEGER, actividad_id TEXT, codigo_actividad TEXT,
    nombre_actividad TEXT, costo_total_control_calidad REAL,
    porcentaje_distribucion REAL, costo_control_calidad_ac REAL,
    formula TEXT, fuente TEXT, observacion TEXT
);

CREATE TABLE IF NOT EXISTS soporte_ingenieria_ac (
    run_id INTEGER, actividad_id TEXT, codigo_actividad TEXT,
    nombre_actividad TEXT, costo_total_ingenieria REAL,
    porcentaje_distribucion REAL, costo_ingenieria_ac REAL,
    formula TEXT, fuente TEXT, observacion TEXT
);

CREATE TABLE IF NOT EXISTS soporte_actividad (
    run_id INTEGER, actividad_id TEXT, codigo_actividad TEXT,
    nombre_actividad TEXT, costo_mantenimiento_ac REAL,
    costo_control_calidad_ac REAL, costo_ingenieria_ac REAL,
    costo_soporte_ac REAL, formula TEXT, fuente TEXT, observacion TEXT
);

CREATE TABLE IF NOT EXISTS actividad_total (
    run_id INTEGER, actividad_id TEXT, codigo_actividad TEXT,
    nombre_actividad TEXT, costo_mod_ac REAL, costo_otros_recursos_ac REAL,
    costo_conversion_ac REAL, costo_soporte_ac REAL, costo_total_ac REAL,
    formula TEXT, fuente TEXT, observacion TEXT
);

CREATE TABLE IF NOT EXISTS produccion_ot_detalle (
    run_id INTEGER, actividad_id TEXT, codigo_actividad TEXT,
    nombre_actividad TEXT, ot_id TEXT, objeto_costo_id TEXT, codigo_producto TEXT,
    peso_ac_ot REAL, costo_conversion_ac REAL, costo_mod_ac REAL,
    costo_otros_recursos_ac REAL, costo_conversion_ot REAL, costo_mod_ot REAL,
    costo_otros_recursos_ot REAL, formula TEXT, fuente TEXT, observacion TEXT
);

CREATE TABLE IF NOT EXISTS produccion_ot (
    run_id INTEGER, ot_id TEXT, objeto_costo_id TEXT, costo_mod_ot REAL,
    costo_otros_recursos_ot REAL, costo_produccion_ot REAL, formula TEXT,
    fuente TEXT, observacion TEXT
);

CREATE TABLE IF NOT EXISTS soporte_ot_detalle (
    run_id INTEGER, actividad_id TEXT, codigo_actividad TEXT,
    nombre_actividad TEXT, ot_id TEXT, objeto_costo_id TEXT, codigo_producto TEXT,
    peso_ac_ot REAL, costo_soporte_ac REAL, costo_soporte_ot REAL,
    formula TEXT, fuente TEXT, observacion TEXT
);

CREATE TABLE IF NOT EXISTS soporte_ot (
    run_id INTEGER, ot_id TEXT, objeto_costo_id TEXT, costo_soporte_ot REAL,
    formula TEXT, fuente TEXT, observacion TEXT
);

CREATE TABLE IF NOT EXISTS hoja_costos_ot (
    run_id INTEGER, empresa_id TEXT, periodo_id TEXT, modelo_id TEXT,
    ot_id TEXT, codigo_ot TEXT, objeto_costo_id TEXT, codigo_producto TEXT,
    cantidad_producida REAL, cmd_ot REAL, mod_ot REAL, otros_recursos_ot REAL,
    costo_produccion_ot REAL, costo_soporte_ot REAL, csd_ot REAL,
    costo_total_ot REAL, costo_unitario_ot REAL, formula TEXT, fuente TEXT,
    observacion TEXT
);

CREATE TABLE IF NOT EXISTS estado_resultados_ot (
    run_id INTEGER, ot_id TEXT, objeto_costo_id TEXT, codigo_producto TEXT,
    ventas_netas_soles REAL, materia_prima_directa REAL,
    costos_conversion REAL, margen_bruto REAL, procesos_soporte REAL,
    margen_operativo REAL, rentabilidad_operativa REAL, formula TEXT,
    fuente TEXT, observacion TEXT
);

CREATE TABLE IF NOT EXISTS dashboard_alertas (
    run_id INTEGER, alerta_id INTEGER PRIMARY KEY AUTOINCREMENT,
    tipo TEXT, severidad TEXT, mensaje TEXT, valor_referencia REAL,
    observacion TEXT
);
