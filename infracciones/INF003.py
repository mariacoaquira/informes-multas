import streamlit as st
import pandas as pd
from datetime import date
from docxtpl import DocxTemplate, RichText
import io
from num2words import num2words
from funciones import create_table_subdoc, create_main_table_subdoc
from sheets import calcular_beneficio_ilicito, calcular_multa, descargar_archivo_drive
from modulos.calculo_capacitacion import calcular_costo_capacitacion

# --- FUNCIÓN 1: DIBUJAR LOS INPUTS EN LA INTERFAZ ---
def renderizar_inputs_especificos(i):
    """
    Dibuja en la interfaz de Streamlit los campos de entrada específicos para INF003.
    Devuelve un diccionario con los datos ingresados por el usuario.
    """
    st.markdown("##### Detalles de la Supervisión para Capacitación")
    datos_especificos = {}
    
    col1, col2 = st.columns(2)
    with col1:
        fecha_supervision = st.date_input("Fecha de supervisión", value=None, format="DD/MM/YYYY", key=f"fecha_supervision_{i}")
        if fecha_supervision:
            st.info(f"Fecha de Incumplimiento: **{fecha_supervision.strftime('%d/%m/%Y')}**")
            datos_especificos['fecha_incumplimiento'] = fecha_supervision
            
    with col2:
        num_personal = st.number_input("Número de personal para capacitación", min_value=1, step=1, key=f"num_personal_{i}")
        datos_especificos['num_personal_capacitacion'] = num_personal
    
    st.subheader("Anexar análisis económico")
    datos_especificos['doc_adjunto_hecho'] = st.file_uploader(
        "Sube el Word con el análisis económico para este hecho:", 
        type=["docx"], 
        key=f"doc_analisis_{i}"
    )
    
    # Devuelve un diccionario con todos los datos que el usuario ingresó
    return datos_especificos

def validar_inputs(datos_especificos):
    """
    Verifica que los datos específicos para esta infracción estén completos.
    Devuelve True si todo está OK, de lo contrario False.
    """
    if not datos_especificos.get('fecha_incumplimiento'):
        return False
    if not datos_especificos.get('num_personal_capacitacion'):
        return False
    # Puedes añadir más validaciones aquí si es necesario
    
    return True

# --- FUNCIÓN 2: PROCESAR TODOS LOS CÁLCULOS Y DATOS ---
def procesar_infraccion(datos_comunes, datos_especificos):
    # 1. Enriquecer datos comunes con datos específicos
    fecha_inc = datos_especificos.get('fecha_incumplimiento')
    if not fecha_inc: return {'error': 'Falta la fecha de incumplimiento.'}
    
    datos_comunes['fecha_incumplimiento'] = fecha_inc
    df_indices = datos_comunes['df_indices']
    fecha_inc_dt = pd.to_datetime(fecha_inc)
    ipc_row = df_indices[df_indices['Indice_Mes'].dt.to_period('M') == fecha_inc_dt.to_period('M')]
    
    if ipc_row.empty: return {'error': f'No se encontraron datos de IPC/TC para la fecha {fecha_inc_dt.strftime("%m-%Y")}.'}
    datos_comunes['ipc_incumplimiento'] = ipc_row.iloc[0]['IPC_Mensual']
    datos_comunes['tipo_cambio_incumplimiento'] = ipc_row.iloc[0]['TC_Mensual']

    # 2. Cálculos
    resultados_ce = calcular_costo_capacitacion(datos_especificos.get('num_personal_capacitacion', 1), datos_comunes)
    ce_data_raw = resultados_ce.get('items_calculados', [])
    total_soles = sum(item['monto_soles'] for item in ce_data_raw)
    total_dolares = sum(item['monto_dolares'] for item in ce_data_raw)
    
    datos_para_bi = {**datos_comunes, 'ce_soles': total_soles, 'ce_dolares': total_dolares}
    resultados_bi = calcular_beneficio_ilicito(datos_para_bi)
    if resultados_bi.get('error'): return {'error': resultados_bi['error']}
    
    b_ilicito = resultados_bi.get('beneficio_ilicito_uit', 0)
    datos_para_multa = {**datos_comunes, 'beneficio_ilicito': b_ilicito}
    resultados_multa = calcular_multa(datos_para_multa)
    # --- AÑADE ESTA LÍNEA PARA OBTENER EL VALOR DE LA MULTA ---
    multa_hecho_uit = resultados_multa.get('multa_final_uit', 0)
    # -----------------------------------------------------------

    # 3. Preparación del Contexto para Word
    doc_tpl = datos_comunes['doc_tpl']

    # --- AÑADE ESTE BLOQUE PARA PREPARAR LOS PLACEHOLDERS FALTANTES ---
    fecha_inc_dt = pd.to_datetime(datos_comunes.get('fecha_incumplimiento'))
    ipc_incumplimiento = datos_comunes.get('ipc_incumplimiento', 0)
    tc_incumplimiento = datos_comunes.get('tipo_cambio_incumplimiento', 0)
    
    fi_mes_texto = fecha_inc_dt.strftime('%B de %Y').lower() if pd.notna(fecha_inc_dt) else "N/A"
    # ---------------------------------------------------------------------
    
    # --- INICIO DEL CÓDIGO COMPLETO PARA TABLAS ---
    # Tabla Costo Evitado (CE)
    tabla_ce_subdoc = None
    if ce_data_raw:
        ce_table_formatted = []
        for item in ce_data_raw:
            ce_table_formatted.append({
                'descripcion': item.get('descripcion', ''),
                'precio_usd': f"US$ {item.get('precio_dolares', 0):,.3f}",
                'precio_soles': f"S/ {item.get('precio_soles', 0):,.3f}",
                'factor_ajuste': f"{item.get('factor_ajuste', 0):,.3f}",
                'monto_soles': f"S/ {item.get('monto_soles', 0):,.3f}",
                'monto_usd': f"US$ {item.get('monto_dolares', 0):,.3f}"
            })
        ce_table_formatted.append({
            'descripcion': 'Total', 'precio_usd': '', 'precio_soles': '', 'factor_ajuste': '',
            'monto_soles': f"S/ {total_soles:,.3f}", 'monto_usd': f"US$ {total_dolares:,.3f}"
        })
        tabla_ce_subdoc = create_table_subdoc(
            doc_tpl,
            ["Descripción", "Precio asociado (US$)", "Precio asociado (S/)", "Factor de ajuste", "Monto (S/)", "Monto (US$)"],
            ce_table_formatted,
            ['descripcion', 'precio_usd', 'precio_soles', 'factor_ajuste', 'monto_soles', 'monto_usd']
        )

    # Tabla Beneficio Ilícito (BI)
    tabla_bi_subdoc = None
    bi_data_cruda = resultados_bi.get('bi_data_raw', [])
    if bi_data_cruda:
        bi_table_formatted = [{'descripcion': item.get('Descripción', ''), 'monto': item.get('Monto', '')} for item in bi_data_cruda]
        tabla_bi_subdoc = create_main_table_subdoc(
            doc_tpl, ["Descripción", "Monto"], bi_table_formatted, ['descripcion', 'monto']
        )

    # Tabla Multa
    tabla_multa_subdoc = None
    multa_data_cruda = resultados_multa.get('multa_data_raw', [])
    if multa_data_cruda:
        multa_table_formatted = [{'componente': item.get('Componentes', ''), 'monto': item.get('Monto', '')} for item in multa_data_cruda]
        tabla_multa_subdoc = create_main_table_subdoc(
            doc_tpl, ["Componentes", "Monto"], multa_table_formatted, ['componente', 'monto']
        )
    # --- FIN DEL CÓDIGO COMPLETO PARA TABLAS ---

    # Formatear otros placeholders
    num_personas = datos_especificos.get('num_personal_capacitacion', 0)
    texto_personas = "una" if num_personas == 1 else num2words(num_personas, lang='es')
    persona_cap = f"{texto_personas} ({num_personas}) persona{'s' if num_personas != 1 else ''}"

    precio_dol_texto = "No aplica"
    if ce_data_raw and ce_data_raw[0].get('moneda_original') == 'US$':
        costo_orig = ce_data_raw[0].get('costo_original', 0)
        precio_dol_texto = f"US$ {costo_orig:,.3f}"

    # ... (código para formatear cualquier otro placeholder que necesites) ...

    # Agrupar datos para el contenedor 'hecho'
    datos_para_hecho = {
        'numero_imputado': datos_comunes['numero_hecho_actual'],
        'descripcion': RichText(datos_especificos.get('texto_hecho', '')),
        'tabla_ce': tabla_ce_subdoc,
        'tabla_bi': tabla_bi_subdoc,
        'tabla_multa': tabla_multa_subdoc,
        'persona_cap': persona_cap,
    }

    contexto_final = { 
    **datos_comunes['context_data'], 
    **datos_especificos, 
    'hecho': datos_para_hecho,
    
    # --- AÑADE ESTOS PLACEHOLDERS FALTANTES ---
    'precio_dol': precio_dol_texto,
    'fecha_incumplimiento_texto': fecha_inc_dt.strftime('%d de %B de %Y').lower(),
    'fuente_cos': resultados_bi.get('fuente_cos', ''),
    # --- AÑADE ESTAS LÍNEAS AL DICCIONARIO ---
    'fi_mes': fi_mes_texto,
    'fi_ipc': f"{ipc_incumplimiento:,.3f}",
    'fi_tc': f"{tc_incumplimiento:,.3f}",
    # --- AÑADE EL NUEVO PLACEHOLDER AQUÍ ---
    'mh_uit': f"{multa_hecho_uit:,.3f} UIT",
    # --- AÑADE ESTA LÍNEA PARA EL BENEFICIO ILÍCITO DEL HECHO ---
    'bi_uit': f"{b_ilicito:,.3f} UIT"
    # -----------------------------------------------------------
}

    # 4. Generar el Anexo de Costo Evitado
    anexos_ce_finales = []
    fila_infraccion = datos_comunes['df_tipificacion'][datos_comunes['df_tipificacion']['ID_Infraccion'] == datos_comunes['id_infraccion']]
    id_plantilla_anexo_ce = fila_infraccion.iloc[0].get('ID_Plantilla_CE')
    if id_plantilla_anexo_ce:
        buffer_anexo = descargar_archivo_drive(id_plantilla_anexo_ce)
        if buffer_anexo:
            anexo_tpl = DocxTemplate(buffer_anexo)
            anexo_tpl.render(contexto_final)
            buffer_final_anexo = io.BytesIO()
            anexo_tpl.save(buffer_final_anexo)
            anexos_ce_finales.append(buffer_final_anexo)

    # 4. DEVOLVER TODOS LOS RESULTADOS
    return {
        'contexto_final_word': contexto_final,
        'resultados_para_app': {
            'ce_data_raw': ce_data_raw,
            'ce_total_soles': total_soles,
            'ce_total_dolares': total_dolares,
            'bi_data_raw': resultados_bi.get('bi_data_raw', []),
            # --- AÑADE ESTA LÍNEA ---
            'beneficio_ilicito_uit': b_ilicito,
            # ------------------------
            'multa_data_raw': resultados_multa.get('multa_data_raw', []),
            'multa_final_uit': multa_hecho_uit
        },
        # --- AÑADE ESTA LÍNEA ---
        'anexos_ce_generados': anexos_ce_finales,
        # ------------------------
        'ids_anexos': resultados_ce.get('ids_anexos', [])
    }
