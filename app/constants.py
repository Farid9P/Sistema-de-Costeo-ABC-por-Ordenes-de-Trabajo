from __future__ import annotations

from pathlib import Path


ROOT_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT_DIR / "data"
DOCS_DIR = ROOT_DIR / "docs"
EXPORTS_DIR = ROOT_DIR / "exports"
REPORTS_DIR = ROOT_DIR / "reports"
TMP_DIR = ROOT_DIR / "tmp"
IMAGES_MODELS_DIR = ROOT_DIR / "images_models"
DEFAULT_DB_PATH = ROOT_DIR / "costos_abc_frenosa.sqlite"
DEFAULT_EXCEL_PATH = DATA_DIR / "frenosa_datos_entrada_abc.xlsx"

APP_TITLE = "Sistema de Costos ABC Frenosa"
MODEL_1_VISUAL_NAME = "Modelo ABC por Ordenes de Trabajo"
MODEL_2_BLOCK_MESSAGE = (
    "El Modelo 2 requiere la ruta EC -> CA -> AC -> OC; no se puede ejecutar "
    "porque el Excel importado no contiene centros de actividad ni los "
    "inductores necesarios para EC-CA y CA-AC."
)
MODEL_3_BLOCK_MESSAGE = (
    "El Modelo 3 requiere una estructura escalonada de centros y actividades "
    "de apoyo, operacion y estrategia; no se puede ejecutar con la informacion "
    "disponible de Frenosa."
)

REQUIRED_FILES = [
    "AGENTS.md",
    "CONTEXT.md",
    "docs/MONOGRADIA 01 GE605U 2026 01.pdf",
    "docs/Semana 07 GE605U 2026 01.pdf",
    "docs/Frenosa_Caso.pdf",
    "data/frenosa_datos_entrada_abc.xlsx",
    "data/Modelo 1 Costeo ABC Textil La Tela.xls",
]

SHEET_SPECS: dict[str, list[str]] = {
    "Empresa": [
        "empresa_id",
        "nombre_empresa",
        "pais",
        "departamento",
        "sector",
        "ciiu",
        "actividad_principal",
        "fuente",
        "observacion",
    ],
    "Periodos": [
        "periodo_id",
        "empresa_id",
        "periodo",
        "tipo_periodo",
        "descripcion",
        "fuente",
    ],
    "Modelos_Costeo": [
        "modelo_id",
        "nombre_modelo",
        "ruta_modelo",
        "descripcion",
        "fuente",
    ],
    "Objetos_Costo": [
        "objeto_costo_id",
        "empresa_id",
        "periodo_id",
        "modelo_id",
        "codigo_producto",
        "descripcion_producto",
        "peso_kg",
        "cantidad_piezas",
        "material_base",
        "fuente_cuadro",
        "fuente",
    ],
    "Ordenes_Trabajo": [
        "ot_id",
        "empresa_id",
        "periodo_id",
        "modelo_id",
        "objeto_costo_id",
        "codigo_ot",
        "descripcion_ot",
        "cantidad_producida",
        "unidad",
        "fuente",
        "observacion",
    ],
    "Materiales_Directos": [
        "material_id",
        "ot_id",
        "objeto_costo_id",
        "material",
        "cantidad_kg",
        "costo_unitario_soles",
        "clasificacion",
        "tratamiento",
        "fuente_cuadro",
        "fuente",
        "observacion",
    ],
    "Personal_directos": [
        "mod_input_id",
        "empresa_id",
        "periodo_id",
        "modelo_id",
        "actividad_id",
        "codigo_actividad",
        "nombre_actividad",
        "kg_producidos_periodo",
        "horas_reales_trabajadas_periodo",
        "cantidad_operarios",
        "salario_mensual_operario_soles",
        "tiempo_normal_min_kg",
        "tolerancia_personal_pct",
        "tolerancia_fatiga_pct",
        "tolerancia_pie_pct",
        "tolerancia_encorvado_pct",
        "redondear_te_decimales",
        "redondear_porcentaje_decimales",
        "usar_porcentaje_redondeado_pdf",
        "fuente_cuadro",
        "fuente",
        "observacion",
        "horas_estandar_totales_periodo",
        "tiempo_real_total_minutos_periodo",
    ],
    "Servicios_Directos": [
        "servicio_id",
        "ot_id",
        "servicio_directo",
        "cantidad",
        "costo_unitario_soles",
        "fuente",
        "observacion",
    ],
    "costo_otros_recursos(CI)": [
        "elemento_costo_id",
        "codigo_recurso",
        "nombre_recurso",
        "importe_soles",
        "clasificacion",
        "tipo",
        "inductor_recurso",
        "fuente_cuadro",
        "fuente",
        "observacion",
    ],
    "Actividades": [
        "actividad_id",
        "codigo_actividad",
        "nombre_actividad",
        "proceso",
        "tiempo_normal_min_kg",
        "clasificacion",
        "fuente_cuadro",
        "fuente",
        "observacion",
    ],
    "EC_AC_Inductores": [
        "elemento_costo_id",
        "codigo_recurso",
        "actividad_id",
        "codigo_actividad",
        "nombre_actividad",
        "inductor_recurso",
        "porcentaje_inductor",
        "fuente_cuadro",
        "fuente",
        "observacion",
    ],
    "Soporte_Mantenimiento(CI)": [
        "soporte_item_id",
        "proceso_soporte",
        "equipo_codigo",
        "equipo_descripcion",
        "costo_equipo_soles",
        "actividad_id_destino",
        "codigo_actividad_destino",
        "fuente_cuadro",
        "fuente",
        "observacion",
    ],
    "Soporte_Control_Calidad(CI)": [
        "soporte_item_id",
        "proceso_soporte",
        "prueba_laboratorio",
        "cantidad",
        "costo_unitario_soles",
        "fuente_cuadro",
        "fuente",
        "observacion",
    ],
    "Soporte_Ingenieria(CI)": [
        "soporte_item_id",
        "proceso_soporte",
        "tipo_plano",
        "cantidad",
        "costo_unitario_soles",
        "fuente_cuadro",
        "fuente",
        "observacion",
    ],
    "Soporte_CC_ING_Inductores": [
        "proceso_soporte",
        "actividad_id",
        "codigo_actividad",
        "nombre_actividad",
        "porcentaje_distribucion",
        "base_distribucion",
        "fuente_cuadro",
        "fuente",
        "observacion",
    ],
    "AC_OT_Pesos_Produccion": [
        "actividad_id",
        "codigo_actividad",
        "nombre_actividad",
        "ot_id",
        "objeto_costo_id",
        "codigo_producto",
        "porcentaje_inductor",
        "tipo_costo",
        "fuente_cuadro",
        "fuente",
        "observacion",
    ],
    "AC_OT_Pesos_Soporte": [
        "actividad_id",
        "codigo_actividad",
        "nombre_actividad",
        "ot_id",
        "objeto_costo_id",
        "codigo_producto",
        "porcentaje_inductor",
        "tipo_costo",
        "fuente_cuadro",
        "fuente",
        "observacion",
    ],
    "Ventas": [
        "venta_id",
        "ot_id",
        "objeto_costo_id",
        "codigo_producto",
        "ventas_netas_soles",
        "fuente_cuadro",
        "fuente",
        "observacion",
    ],
    "Reglas_Calculo": [
        "regla_id",
        "modulo",
        "regla_calculo",
        "salida_esperada",
        "observacion",
    ],
    "Fuentes": [
        "fuente_id",
        "documento",
        "cuadro_o_seccion",
        "dato_usado",
        "url",
        "observacion",
    ],
    "Diccionario": [
        "hoja",
        "tipo",
        "descripcion",
        "calculos_no_incluidos",
    ],
}

OPTIONAL_SHEET_SPECS: dict[str, list[str]] = {
    "README": ["Campo", "Descripcion"],
    "Inductores": [
        "inductor_id",
        "nombre_inductor",
        "descripcion",
        "unidad",
        "fuente_cuadro",
        "fuente",
    ],
}

SHEET_TO_TABLE = {
    "Empresa": "empresas",
    "Periodos": "periodos",
    "Modelos_Costeo": "modelos_costeo",
    "Objetos_Costo": "objetos_costo",
    "Ordenes_Trabajo": "ordenes_trabajo",
    "Materiales_Directos": "materiales_directos",
    "Personal_directos": "Personal_directos",
    "Servicios_Directos": "servicios_directos",
    "costo_otros_recursos(CI)": "costo_otros_recursos(CI)",
    "Actividades": "actividades",
    "Inductores": "inductores",
    "EC_AC_Inductores": "ec_ac_inductores",
    "Soporte_Mantenimiento(CI)": "Soporte_Mantenimiento(CI)",
    "Soporte_Control_Calidad(CI)": "Soporte_Control_Calidad(CI)",
    "Soporte_Ingenieria(CI)": "Soporte_Ingenieria(CI)",
    "Soporte_CC_ING_Inductores": "soporte_cc_ing_inductores",
    "AC_OT_Pesos_Produccion": "ac_ot_pesos_produccion",
    "AC_OT_Pesos_Soporte": "ac_ot_pesos_soporte",
    "Ventas": "ventas",
    "Reglas_Calculo": "reglas_calculo",
    "Fuentes": "fuentes",
    "Diccionario": "diccionario",
}

OUTPUT_REPORTS = [
    "Reporte_Proceso_1_EC_AC.xlsx",
    "Reporte_Proceso_2_AC_OT.xlsx",
    "Reporte_Hoja_Costos_OT.xlsx",
    "Reporte_Estado_Resultados.xlsx",
    "Reporte_Dashboard.xlsx",
]

OUTPUT_DOCUMENTS = [
    "Manual_descriptivo.docx",
    "Manual_usuario.docx",
    "Informe_pruebas.docx",
    "Informe_propuestas_mejora.docx",
    "Analisis_funcional_diseno_sistema.docx",
    "Codigos_programacion_principales.docx",
]
