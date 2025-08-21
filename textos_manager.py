# --- BIBLIOTECA CENTRAL DE PLANTILLAS DE FUENTES ---
# Usamos f-strings para poder insertar valores dinámicos como fechas o nombres.
FUENTES_TEMPLATES = {
    
    # Fuentes para Beneficio Ilícito
    'cok': (
        "{fuente_cos}"
    ),
    'periodo_bi': (
        "El periodo de capitalización es contabilizado a partir del primer día calendario siguiente a la "
        "fecha límite en que se debió presentar la documentación ({fecha_incumplimiento_texto}) hasta la fecha del cálculo de la multa ({fecha_hoy_texto})."
    ),
    'periodo_bi_ext': (
        "El periodo de capitalización es contabilizado a partir del primer día calendario siguiente a la "
        "fecha límite en que se debió presentar la documentación ({fecha_incumplimiento_texto}) hasta la fecha de cumplimiento extemporáneo ({fecha_extemporanea_texto})."
    ),

    'ce_ext':("Costo evitado capitalizado hasta la fecha de cumplimiento extemporáneo."),

    'ajuste_inflacionario_detalle': (
        "Este ajuste se aplica por el efecto inflacionario, considerando los Índice de Precios al Consumidor (IPC) de cada periodo bajo análisis: "
        "F = IPC disponible a la fecha de emisión del presente informe / IPC a la fecha de cumplimiento extemporáneo = "
        "IPC de {mes_ipc_hoy_texto} / IPC de {mes_ipc_ext_texto}, "
        "equivalente a F = {valor_ipc_hoy} / {valor_ipc_ext}. Se considera el redondeo a 3 decimales."
    ),

    'bcrp': (
        "Banco Central de Reserva del Perú (BCRP). Series Estadísticas. Tipo de Cambio Nominal Bancario-Promedio "
        "de los últimos 12 meses. Fecha de consulta: {fecha_hoy_texto}. https://estadisticas.bcrp.gob.pe/estadisticas/series/mensuales/resultados/PN01210PM/html"
    ),
    'ipc_fecha': (
        "Cabe precisar que, si bien la fecha de emisión del informe corresponde al mes de {mes_actual_texto}, la fecha "
        "considerada para el IPC y el TC fue hasta {ultima_fecha_ipc_texto}, toda vez que, dicha información se "
        "encontraba disponible a la fecha de emisión del presente informe."
    ),

    'sunat': ("SUNAT - Índices y tasas. (http://www.sunat.gob.pe/indicestasas/uit.html)"),
    
    # Fuentes para Costo Evitado (puedes añadir más)
    'ce_anexo': (
        "El costo evitado se estimó en un escenario de incumplimiento según el periodo correspondiente, con sus factores de ajuste respectivos (IPC y tipo de cambio). Ver Anexo n.° 1."
    ),
    
    # ...AQUÍ PUEDES AÑADIR CUALQUIER OTRA FUENTE QUE NECESITES EN EL FUTURO...
}

def obtener_fuente_formateada(ref_key, datos):
    """
    Busca una plantilla de fuente por su clave y la formatea con los datos proporcionados.
    Si una clave en los datos no existe en la plantilla, no da error.
    """
    template = FUENTES_TEMPLATES.get(ref_key, f"Error: Fuente '{ref_key}' no encontrada.")
    # Usamos un bucle para reemplazar solo las claves que existen en la plantilla.
    for key, value in datos.items():
        template = template.replace(f"{{{key}}}", str(value))
    return template
