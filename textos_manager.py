# --- BIBLIOTECA CENTRAL DE PLANTILLAS DE FUENTES ---
FUENTES_TEMPLATES = {
    # Textos para tablas de Beneficio Ilícito (BI)
    'BI': {
        # Textos que aparecen en TODAS las tablas de BI (normal y extemporáneo)
        'GENERAL': {
            'cok': (
                "{fuente_cos}"
            ),
            'sunat': ("SUNAT - Índices y tasas. (http://www.sunat.gob.pe/indicestasas/uit.html)"),
            'bcrp': (
                "Banco Central de Reserva del Perú (BCRP). Series Estadísticas. Tipo de Cambio Nominal Bancario-Promedio "
                "de los últimos 12 meses. Fecha de consulta: {fecha_hoy_texto}. https://estadisticas.bcrp.gob.pe/estadisticas/series/mensuales/resultados/PN01210PM/html"
            ),
            'ipc_fecha': (
                "Cabe precisar que, si bien la fecha de emisión del informe corresponde al mes de {mes_actual_texto}, la fecha "
                "considerada para el IPC y el TC fue hasta {ultima_fecha_ipc_texto}, toda vez que, dicha información se "
                "encontraba disponible a la fecha de emisión del presente informe."
            ),
            'ce_anexo': (
                "El costo evitado se estimó en un escenario de incumplimiento según el periodo correspondiente, con sus factores de ajuste respectivos (IPC y tipo de cambio). Ver Anexo n.° 1."
            )
        },
        # Textos que aparecen SÓLO en tablas de BI Extemporáneo
        'EXTEMPORANEO': {
            'ce_ext': ("Costo evitado capitalizado hasta la fecha de cumplimiento extemporáneo."),
            'ajuste_inflacionario_detalle': (
                "Este ajuste se aplica por el efecto inflacionario, considerando los Índice de Precios al Consumidor (IPC) de cada periodo bajo análisis: "
                "F = IPC disponible a la fecha de emisión del presente informe / IPC a la fecha de cumplimiento extemporáneo = "
                "IPC de {mes_ipc_hoy_texto} / IPC de {mes_ipc_ext_texto}, "
                "equivalente a F = {valor_ipc_hoy} / {valor_ipc_ext}. Se considera el redondeo a 3 decimales."
            )
        }
    },
    
    # Textos que varían según la infracción específica
    'INFRACCIONES': {
        'INF003': { # Obstaculización a la supervisión
            'periodo_bi': (
                "El periodo de capitalización se contabiliza a partir del último día de supervisión in situ en la cual no se facilitó el ingreso al equipo supervisor ({fecha_incumplimiento_texto}) "
                "hasta la fecha del cálculo de la multa ({fecha_hoy_texto})."
            ),
            'periodo_bi_ext': (
                "El periodo de capitalización se contabiliza a partir del último día de supervisión in situ en la cual no se facilitó el ingreso al equipo supervisor ({fecha_incumplimiento_texto}) "
                "hasta la fecha en que se subsanó la conducta ({fecha_extemporanea_texto})."
            )
        },
        'INF004': { # No remitir información
            'periodo_bi': (
                "El periodo de capitalización se contabiliza a partir del primer día calendario siguiente a la "
                "fecha límite en que se debió presentar la documentación ({fecha_incumplimiento_texto}) hasta la fecha del cálculo de la multa ({fecha_hoy_texto})."
            ),
            'periodo_bi_ext': (
                "El periodo de capitalización se contabiliza a partir del primer día calendario siguiente a la "
                "fecha límite en que se debió presentar la documentación ({fecha_incumplimiento_texto}) hasta la fecha de cumplimiento extemporáneo ({fecha_extemporanea_texto})."
            )
        },
        'INF005': { # No remitir monitoreo
            'periodo_bi': (
                "El periodo de capitalización se contabiliza a partir del primer día calendario siguiente a la "
                "fecha límite en que se debió presentar la documentación ({fecha_incumplimiento_texto}) hasta la fecha del cálculo de la multa ({fecha_hoy_texto})."
            ),
            'periodo_bi_ext': (
                "El periodo de capitalización se contabiliza a partir del primer día calendario siguiente a la "
                "fecha límite en que se debió presentar la documentación ({fecha_incumplimiento_texto}) hasta la fecha de cumplimiento extemporáneo ({fecha_extemporanea_texto})."
            )
        },
                'INF007': { # No presentar manifiesto
            'periodo_bi': (
                "El período de capitalización se contabiliza a partir del primer día calendario siguiente a la fecha límite establecida para presentar el manifiesto a través de la plataforma SIGERSOL ({fecha_incumplimiento_texto}) hasta la fecha del cálculo de la multa ({fecha_hoy_texto})."
            ),
            'periodo_bi_ext': (
                "El período de capitalización se contabiliza a partir del primer día calendario siguiente a la fecha límite establecida para presentar el manifiesto a través de la plataforma SIGERSOL ({fecha_incumplimiento_texto}) hasta la fecha de cumplimiento extemporáneo ({fecha_extemporanea_texto})."
            )
        },
                'INF008': { # No presentar declaración
            'periodo_bi': (
                "El período de capitalización se contabiliza a partir del primer día calendario siguiente a la fecha límite establecida para presentar la declaración a través de la plataforma SIGERSOL ({fecha_incumplimiento_texto}) hasta la fecha del cálculo de la multa ({fecha_hoy_texto})."
            ),
            'periodo_bi_ext': (
                "El período de capitalización se contabiliza a partir del primer día calendario siguiente a la fecha límite establecida para presentar la declaración a través de la plataforma SIGERSOL ({fecha_incumplimiento_texto}) hasta la fecha de cumplimiento extemporáneo ({fecha_extemporanea_texto})."
            )
        },
                'INF009': { # No presentar declaración
            'periodo_bi': (
                "El período de capitalización se contabiliza a partir del primer día calendario siguiente a la fecha límite establecida para conducir un adecuado registro interno ({fecha_incumplimiento_texto}) hasta la fecha del cálculo de la multa ({fecha_hoy_texto})."
            ),
            'periodo_bi_ext': (
                "El período de capitalización se contabiliza a partir del primer día calendario siguiente a la fecha límite establecida para conducir un adecuado registro interno ({fecha_incumplimiento_texto}) hasta la fecha de cumplimiento extemporáneo ({fecha_extemporanea_texto})."
            )
        },
                'INF002': { # No realizar monitoreo
            'periodo_bi': (
                "El período de capitalización se contabiliza a partir del primer día calendario siguiente al plazo máximo para realizar el monitoreo ({fecha_incumplimiento_texto}) hasta la fecha del cálculo de la multa ({fecha_hoy_texto})."
            ),
            'periodo_bi_ext': (
                "El período de capitalización se contabiliza a partir del primer día calendario siguiente al plazo máximo para realizar el monitoreo ({fecha_incumplimiento_texto}) hasta la fecha de cumplimiento extemporáneo ({fecha_extemporanea_texto})."
            )
        },
        # Un texto DEFAULT como respaldo si una infracción no tiene su propia versión
        'DEFAULT': {
            'periodo_bi': "Periodo de capitalización desde {fecha_incumplimiento_texto} hasta {fecha_hoy_texto}.",
            'periodo_bi_ext': "Periodo de capitalización desde {fecha_incumplimiento_texto} hasta {fecha_extemporanea_texto}."
        }
    }
}

def obtener_fuente_formateada(ref_key, datos, id_infraccion=None, es_extemporaneo=False):
    """
    Busca una plantilla de fuente con una jerarquía lógica y la formatea con los datos.
    """
    template = None
    
    # 1. Buscar en la sección específica de la infracción o en el DEFAULT
    if id_infraccion and id_infraccion in FUENTES_TEMPLATES['INFRACCIONES']:
        template = FUENTES_TEMPLATES['INFRACCIONES'][id_infraccion].get(ref_key)
    if template is None:
        template = FUENTES_TEMPLATES['INFRACCIONES']['DEFAULT'].get(ref_key)

    # 2. Si no es un texto específico de infracción, buscar en las categorías de BI
    if template is None:
        # Si es extemporáneo, buscar primero en la sección EXTEMPORANEO
        if es_extemporaneo:
            template = FUENTES_TEMPLATES['BI']['EXTEMPORANEO'].get(ref_key)
        # Si no se encontró o no es extemporáneo, buscar en la sección GENERAL de BI
        if template is None:
            template = FUENTES_TEMPLATES['BI']['GENERAL'].get(ref_key)
    
    # 3. Si después de todo no se encontró, devolver un error
    if template is None:
        return f"Error: Fuente '{ref_key}' no encontrada."

    # 4. Formatear la plantilla con los datos proporcionados
    for key, value in datos.items():
        template = template.replace(f"{{{key}}}", str(value))
    return template