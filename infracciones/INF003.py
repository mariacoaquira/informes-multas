import streamlit as st
import pandas as pd
from datetime import date
from docxtpl import DocxTemplate, RichText
from docxcompose.composer import Composer
import io
from num2words import num2words
from babel.dates import format_date
from textos_manager import obtener_fuente_formateada
from funciones import (create_table_subdoc, create_main_table_subdoc, texto_con_numero,
                     create_footnotes_subdoc, create_consolidated_bi_table_subdoc,
                     create_personal_table_subdoc, redondeo_excel, format_decimal_dinamico) # <-- AÑADIR ESTAS DOS
from sheets import calcular_beneficio_ilicito, calcular_multa, descargar_archivo_drive
# ASUNCIÓN: Se mantiene la dependencia del módulo de cálculo específico
from modulos.calculo_capacitacion import calcular_costo_capacitacion


# --- 1. INTERFAZ DE USUARIO (CON ENTEROS EN TABLA) ---
def renderizar_inputs_especificos(i, df_dias_no_laborables=None):
    """
    Dibuja la interfaz para la INF003, con la tabla de personal
    anidada dentro del primer extremo y métricas para las fechas.
    """
    datos_hecho = st.session_state.imputaciones_data[i]

    # 1. Asegurar que la lista de datos existe en el estado
    if 'tabla_personal' not in datos_hecho or not isinstance(datos_hecho['tabla_personal'], list):
        datos_hecho['tabla_personal'] = [{'Perfil': '', 'Descripción': '', 'Cantidad': 0}]

    st.markdown("###### **Extremos del incumplimiento**")
    
    if 'extremos' not in datos_hecho: datos_hecho['extremos'] = [{}]
    
    for j, extremo in enumerate(datos_hecho['extremos']):
        with st.container(border=True):
            
            col1, col_display, col_button = st.columns([2, 2, 1])

            with col1:
                fecha_supervision = st.date_input(
                    f"Fecha del Extremo n.° {j + 1}",
                    key=f"fecha_supervision_{i}_{j}",
                    value=extremo.get('fecha_incumplimiento'),
                    format="DD/MM/YYYY",
                    max_value=date.today()
                )
                extremo['fecha_incumplimiento'] = fecha_supervision
                extremo['fecha_base'] = fecha_supervision

            with col_display:
                if extremo.get('fecha_incumplimiento'):
                    fecha_str = extremo['fecha_incumplimiento'].strftime('%d/%m/%Y')
                    st.metric(label="Fecha Incumplimiento", value=fecha_str)
                else:
                    st.metric(label="Fecha Incumplimiento", value="---")

            with col_button:
                if len(datos_hecho['extremos']) > 1:
                    st.write("") 
                    st.write("") 
                    if st.button(f"🗑️", key=f"del_extremo_{i}_{j}", help="Eliminar Extremo"):
                        datos_hecho['extremos'].pop(j)
                        st.rerun()
            
            st.divider() 

            st.markdown("###### **Personal a capacitar (ingresado en el primer extremo)**")

            df_personal = pd.DataFrame(datos_hecho['tabla_personal'])

            if j == 0:
                edited_df = st.data_editor(
                    df_personal,
                    num_rows="dynamic",
                    key=f"data_editor_personal_{i}",
                    hide_index=True,
                    use_container_width=True,
                    disabled=False,
                    column_config={ 
                        "Perfil": st.column_config.TextColumn("Perfil", help="Ej: Personal operativo, Supervisor", required=True),
                        "Descripción": st.column_config.TextColumn("Descripción", help="Detalle de las funciones...", width="large"),
                        # --- INICIO CORRECCIÓN (Formato Entero) ---
                        "Cantidad": st.column_config.NumberColumn("Cantidad", help="Número de personas con este perfil", min_value=0, step=1, required=True, format="%d"),
                        # --- FIN CORRECCIÓN ---
                    }
                )
                datos_hecho['tabla_personal'] = edited_df.to_dict('records')
            
            else:
                 st.data_editor(
                    df_personal,
                    num_rows="dynamic",
                    key=f"data_editor_personal_{i}_disabled_{j}",
                    hide_index=True,
                    use_container_width=True,
                    disabled=True, 
                    # --- INICIO CORRECCIÓN (Formato Entero) ---
                    column_config={ "Perfil": {}, "Descripción": {}, "Cantidad": st.column_config.NumberColumn(format="%d") } 
                    # --- FIN CORRECCIÓN ---
                )

            cantidades_numericas = [pd.to_numeric(p.get('Cantidad'), errors='coerce') for p in datos_hecho['tabla_personal']]
            total_personal = pd.Series(cantidades_numericas).fillna(0).sum()

            datos_hecho['num_personal_capacitacion'] = int(total_personal)
            st.metric("Total de Personal a Capacitar", f"{datos_hecho['num_personal_capacitacion']} persona(s)")

    if st.button("+ Añadir Extremo", key=f"add_extremo_{i}"):
        datos_hecho['extremos'].append({})
        st.rerun()

    return datos_hecho

# --- 2. VALIDACIÓN (SIN CAMBIOS) ---
def validar_inputs(datos_especificos):
    if not datos_especificos.get('num_personal_capacitacion', 0) > 0: return False
    if not datos_especificos.get('extremos'): return False
    for extremo in datos_especificos['extremos']:
        if not extremo.get('fecha_incumplimiento'): return False
    return True

# --- 3. PROCESADOR SIMPLE (CON CORRECCIONES) ---
def _procesar_hecho_simple(datos_comunes, datos_especificos):
    """
    Procesa un hecho con un único extremo, usando la plantilla de BI simple.
    """
    extremo = datos_especificos['extremos'][0]
    datos_especificos['fecha_incumplimiento'] = extremo['fecha_incumplimiento']
    datos_para_ce = {**datos_comunes, 'fecha_incumplimiento': datos_especificos['fecha_incumplimiento']}
    
    resultados_ce = calcular_costo_capacitacion(num_personal=datos_especificos.get('num_personal_capacitacion', 1), datos_comunes=datos_para_ce)
    if resultados_ce.get('error'): return {'error': resultados_ce['error']}
    
    ce_data_raw = resultados_ce.get('items_calculados', [])

    # --- INICIO PRORRATEO EXTERNO (Entre Hechos) ---
    # Verificar si existe un factor de prorrateo para este año en el hecho general
    fecha_inc = datos_especificos['fecha_incumplimiento']
    if fecha_inc:
        anio_inc = fecha_inc.year
        factor_prorrateo = datos_especificos.get('mapa_factores_prorrateo', {}).get(anio_inc, 1.0)
        
        # Si hay prorrateo (factor < 1.0), ajustamos los montos unitarios
        if factor_prorrateo < 1.0:
            for item in ce_data_raw:
                item['monto_soles'] = redondeo_excel(item['monto_soles'] * factor_prorrateo, 3)
                item['monto_dolares'] = redondeo_excel(item['monto_dolares'] * factor_prorrateo, 3)
                # Opcional: Marcar visualmente en descripción
                # item['descripcion'] += f" (Prorrateado {factor_prorrateo:.0%})" 
    # --- FIN PRORRATEO EXTERNO ---
    
    total_soles = sum(item.get('monto_soles', 0) for item in ce_data_raw)
    total_dolares = sum(item.get('monto_dolares', 0) for item in ce_data_raw)

    # --- INICIO: Descripción del Hecho en CE ---
    texto_hecho_bi = datos_especificos.get('texto_hecho', 'Hecho no especificado')
    # --- FIN ---

    datos_para_bi = { **datos_comunes, 'ce_soles': total_soles, 'ce_dolares': total_dolares, 'fecha_incumplimiento': datos_especificos['fecha_incumplimiento'], 'texto_del_hecho': texto_hecho_bi }
    resultados_bi = calcular_beneficio_ilicito(datos_para_bi)
    if resultados_bi.get('error'): return {'error': resultados_bi['error']}
    b_ilicito_uit = resultados_bi.get('beneficio_ilicito_uit', 0)

    # --- INICIO: Recuperar Factor F y Calcular Multa ---
    factor_f = datos_especificos.get('factor_f_calculado', 1.0)

    resultados_multa = calcular_multa({
        **datos_comunes, 
        'beneficio_ilicito': b_ilicito_uit,
        'factor_f': factor_f # <--- NUEVO
    })
    multa_hecho_uit = resultados_multa.get('multa_final_uit', 0)
    # --- FIN ---

    # --- INICIO: LÓGICA DE REDUCCIÓN Y TOPE (Simple) ---
    datos_hecho_completos = datos_comunes.get('datos_hecho_completos', {})
    aplica_reduccion_str = datos_hecho_completos.get('aplica_reduccion', 'No')
    porcentaje_str = datos_hecho_completos.get('porcentaje_reduccion', '0%')
    multa_con_reduccion_uit = multa_hecho_uit
    
    if aplica_reduccion_str == 'Sí':
        reduccion_factor = 0.5 if porcentaje_str == '50%' else 0.7
        multa_con_reduccion_uit = redondeo_excel(multa_hecho_uit * reduccion_factor, 3)

    infraccion_info = datos_comunes['df_tipificacion'][datos_comunes['df_tipificacion']['ID_Infraccion'] == datos_comunes['id_infraccion']]
    tope_multa_uit = float('inf')
    if not infraccion_info.empty and pd.notna(infraccion_info.iloc[0].get('Tope_Multa_Infraccion')):
        tope_multa_uit = float(infraccion_info.iloc[0]['Tope_Multa_Infraccion'])

    multa_final_del_hecho_uit = min(multa_con_reduccion_uit, tope_multa_uit)
    se_aplica_tope = multa_con_reduccion_uit > tope_multa_uit
    multa_reducida_uit = multa_con_reduccion_uit if aplica_reduccion_str == 'Sí' else multa_hecho_uit
    # --- FIN: LÓGICA DE REDUCCIÓN Y TOPE ---
    
    doc_tpl = datos_comunes['doc_tpl']

    ce_table_formatted = []
    # --- INICIO: Buscar Precios Base de CE2 (Simple) ---
    ce2_base_precio_soles_simple = 0.0
    ce2_base_precio_dolares_simple = 0.0
    if ce_data_raw:
        primer_item = ce_data_raw[0]
        ce2_base_precio_soles_simple = primer_item.get('precio_soles', 0.0)
        ce2_base_precio_dolares_simple = primer_item.get('precio_dolares', 0.0)
    # --- FIN ---

    for item in ce_data_raw:
        ce_table_formatted.append({
            'descripcion': item.get('descripcion', ''),
            'precio_dolares': f"US$ {item.get('precio_dolares', 0):,.3f}",
            'precio_soles': f"S/ {item.get('precio_soles', 0):,.3f}",
            'factor_ajuste': f"{item.get('factor_ajuste', 0):,.3f}",
            'monto_soles': f"S/ {item.get('monto_soles', 0):,.3f}",
            'monto_dolares': f"US$ {item.get('monto_dolares', 0):,.3f}"
        })
    ce_table_formatted.append({
        'descripcion': 'Total', 'precio_dolares': '', 'precio_soles': '', 'factor_ajuste': '',
        'monto_soles': f"S/ {total_soles:,.3f}",
        'monto_dolares': f"US$ {total_dolares:,.3f}"
    })
    tabla_ce_subdoc = create_table_subdoc(
        doc_tpl,
        headers=["Descripción", "Precio (US$)", "Precio (S/)", "Factor de ajuste", "Monto (S/)", "Monto (US$)"],
        data=ce_table_formatted,
        keys=['descripcion', 'precio_dolares', 'precio_soles', 'factor_ajuste', 'monto_soles', 'monto_dolares']
    )
    
    # --- INICIO CORRECCIÓN: Superíndices en BI ---
    filas_bi_crudas = resultados_bi.get('table_rows', [])
    filas_bi_para_tabla = []
    for fila in filas_bi_crudas:
        nueva_fila = fila.copy()
        ref_letra = nueva_fila.get('ref')
        texto_base = str(nueva_fila.get('descripcion_texto', ''))
        super_existente = str(nueva_fila.get('descripcion_superindice', ''))
        if ref_letra: super_existente += f"({ref_letra})"
        nueva_fila['descripcion_texto'] = texto_base
        nueva_fila['descripcion_superindice'] = super_existente
        filas_bi_para_tabla.append(nueva_fila)
    # --- FIN ---

    footnote_mapping = resultados_bi.get('footnote_mapping', {})
    datos_para_fuentes = resultados_bi.get('footnote_data', {})
    footnotes_list = [f"({letra}) {obtener_fuente_formateada(ref_key, datos_para_fuentes, id_infraccion=datos_comunes['id_infraccion'])}" for letra, ref_key in sorted(footnote_mapping.items())]
    footnotes_data = {'list': footnotes_list, 'elaboration': 'Elaboración: SSAG - DFAI.', 'style': 'FuenteTabla'}
    
    tabla_bi_subdoc = create_main_table_subdoc(doc_tpl, headers=["Descripción", "Monto"], data=filas_bi_para_tabla, keys=['descripcion_texto', 'monto'], footnotes_data=footnotes_data, column_widths=(5, 1))
    tabla_multa_subdoc = create_main_table_subdoc(doc_tpl, headers=["Componentes", "Monto"], data=resultados_multa.get('multa_data_raw', []), keys=['Componentes', 'Monto'], column_widths=(5, 1))


    # --- INICIO CORRECCIÓN (Tabla Personal - Enteros y Fuente) ---
    tabla_personal_render = datos_especificos.get('tabla_personal', [])
    tabla_personal_render_sin_total = []
    for fila in tabla_personal_render:
        perfil = fila.get('Perfil')
        cantidad = pd.to_numeric(fila.get('Cantidad'), errors='coerce')
        if perfil and cantidad > 0:
            tabla_personal_render_sin_total.append({
                'Perfil': perfil,
                'Descripción': fila.get('Descripción', ''),
                'Cantidad': int(cantidad)
            })

    num_personal_total_int = int(datos_especificos.get('num_personal_capacitacion', 0))
    datos_tabla_personal_word = tabla_personal_render_sin_total + [{'Perfil': 'Total', 'Descripción': '', 'Cantidad': num_personal_total_int}]
    
    tabla_detalle_personal_subdoc = create_personal_table_subdoc(
        doc_tpl,
        headers=["Perfil", "Descripción", "Cantidad"],
        data=datos_tabla_personal_word,
        keys=['Perfil', 'Descripción', 'Cantidad'],
        column_widths=(2, 3, 1),
        texto_posterior="Elaboración: Subdirección de Sanción y Gestión de Incentivos (SSAG) - DFAI."
    )
    # --- FIN CORRECCIÓN ---

    num_personas_total = datos_especificos.get('num_personal_capacitacion', 0)
    nro_personal_texto_anexo = f"{texto_con_numero(num_personas_total)} persona{'s' if num_personas_total != 1 else ''}"
    
    # --- INICIO: Precios Base para App ---
    # Pasamos la lista de datos_ce2 (que ya contiene los precios unitarios)
    ce2_data_para_app = [{'precio_dolares': item['precio_dolares'], 'precio_soles': item['precio_soles']} for item in ce_data_raw]
    # --- FIN ---
    
    contexto_final = { 
        **datos_comunes['context_data'], **datos_especificos, 
        'hecho': {
            'numero_imputado': datos_comunes['numero_hecho_actual'], 
            'descripcion': RichText(datos_especificos.get('texto_hecho', '')), 
            'tabla_ce': tabla_ce_subdoc, 
            'tabla_bi': tabla_bi_subdoc, 
            'tabla_multa': tabla_multa_subdoc
        }, 
        'fuente_cos': resultados_bi.get('fuente_cos', ''), 
        'multa_original_uit': f"{multa_hecho_uit:,.3f} UIT",
        'mh_uit': f"{multa_final_del_hecho_uit:,.3f} UIT", 
        'bi_uit': f"{b_ilicito_uit:,.3f} UIT", 
        'nro_personal': nro_personal_texto_anexo,
        # --- INICIO ADICIÓN: Precios Base para Plantilla ---
        'precio_base_soles': f"S/ {ce2_base_precio_soles_simple:,.3f}",
        'precio_base_dolares': f"US$ {ce2_base_precio_dolares_simple:,.3f}",
        # --- FIN ADICIÓN ---
        'precio_dolares': f"US$ {resultados_ce.get('precio_dolares', 0):,.3f}",
        'fi_mes': resultados_ce.get('fi_mes', ''),
        'fi_ipc': f"{resultados_ce.get('fi_ipc', 0):,.3f}",
        'fi_tc': f"{resultados_ce.get('fi_tc', 0):,.3f}",
        'numeral_hecho': f"IV.{datos_comunes['numero_hecho_actual']}",
        'texto_explicacion_prorrateo': '', 
        'tabla_detalle_personal': tabla_detalle_personal_subdoc,
        
        # --- INICIO: PLACEHOLDERS DE REDUCCIÓN Y TOPE ---
        'aplica_reduccion': aplica_reduccion_str == 'Sí',
        'porcentaje_reduccion': porcentaje_str,
        'texto_reduccion': datos_hecho_completos.get('texto_reduccion', ''),
        'memo_num': datos_hecho_completos.get('memo_num', ''),
        'memo_fecha': format_date(datos_hecho_completos.get('memo_fecha'), "d 'de' MMMM 'de' yyyy", locale='es') if datos_hecho_completos.get('memo_fecha') else '',
        'escrito_num': datos_hecho_completos.get('escrito_num', ''),
        'escrito_fecha': format_date(datos_hecho_completos.get('escrito_fecha'), "d 'de' MMMM 'de' yyyy", locale='es') if datos_hecho_completos.get('escrito_fecha') else '',
        'multa_con_reduccion_uit': f"{multa_con_reduccion_uit:,.3f} UIT",
        'se_aplica_tope': se_aplica_tope,
        'tope_multa_uit': f"{tope_multa_uit:,.3f} UIT",
        # --- FIN ---
    }
    
    anexos_ce_generados = []
    fila_infraccion = datos_comunes['df_tipificacion'][datos_comunes['df_tipificacion']['ID_Infraccion'] == datos_comunes['id_infraccion']].iloc[0]
    id_plantilla_anexo_ce = fila_infraccion.get('ID_Plantilla_CE')
    if id_plantilla_anexo_ce:
        buffer_anexo = descargar_archivo_drive(id_plantilla_anexo_ce)
        if buffer_anexo:
            anexo_tpl = DocxTemplate(buffer_anexo)
            anexo_tpl.render(contexto_final)
            buffer_final_anexo = io.BytesIO()
            anexo_tpl.save(buffer_final_anexo)
            anexos_ce_generados.append(buffer_final_anexo)

    return {
        'contexto_final_word': contexto_final,
        'resultados_para_app': {
            'ce_data_raw': ce_data_raw, 
            'ce_total_soles': total_soles, 
            'ce_total_dolares': total_dolares, 
            'bi_data_raw': resultados_bi.get('table_rows', []), 
            'beneficio_ilicito_uit': b_ilicito_uit, 
            'multa_data_raw': resultados_multa.get('multa_data_raw', []), 
            'multa_final_uit': multa_hecho_uit, 
            'tabla_personal_data': datos_tabla_personal_word,
            
            # --- Datos para app.py ---
            'totales': { # Estructura para app.py
                'ce2_data_raw': ce2_data_para_app,
                'aplica_reduccion': aplica_reduccion_str,
                'porcentaje_reduccion': porcentaje_str,
                'multa_con_reduccion_uit': multa_con_reduccion_uit,
                'multa_reducida_uit': multa_reducida_uit,
                'multa_final_aplicada': multa_final_del_hecho_uit
            }
        },
        'texto_explicacion_prorrateo': '',
        'tabla_detalle_personal': tabla_detalle_personal_subdoc,
        'usa_capacitacion': True,
        'tabla_personal_data': datos_tabla_personal_word, 
        'anexos_ce_generados': anexos_ce_generados,
        'ids_anexos': resultados_ce.get('ids_anexos', [])
    }

# --- 4. PROCESADOR MÚLTIPLE (CON TODAS LAS CORRECCIONES) ---
def _procesar_hecho_multiple(datos_comunes, datos_especificos):
    """
    Procesa múltiples extremos con la lógica de prorrateo, 
    footnotes consolidados y anexos de sustento.
    """
    # a. Cargar plantillas
    df_tipificacion, id_infraccion = datos_comunes['df_tipificacion'], datos_comunes['id_infraccion']
    fila_infraccion = df_tipificacion[df_tipificacion['ID_Infraccion'] == id_infraccion].iloc[0]
    id_plantilla_principal = fila_infraccion.get('ID_Plantilla_BI_Extremo')
    if not id_plantilla_principal: return {'error': 'Falta el ID de la "ID_Plantilla_BI_Extremo" en Tipificación.'}
    
    id_plantilla_anexo_ce = fila_infraccion.get('ID_Plantilla_CE_Extremo')
    if not id_plantilla_anexo_ce: return {'error': 'Falta el ID de la "ID_Plantilla_CE_Extremo" en Tipificación.'}
    buffer_plantilla_anexo_ce = descargar_archivo_drive(id_plantilla_anexo_ce)
    if not buffer_plantilla_anexo_ce: return {'error': 'Fallo en la descarga de la plantilla para el anexo CE.'}

    # b. Inicializar acumuladores y listas
    total_bi_uit = 0
    lista_resultados_bi = []
    anexos_ce_generados, todos_los_ids_anexos, lista_extremos_para_plantilla = [], set(), []
    lista_ce2_data_consolidada = []

    # --- Tabla Personal - Enteros y Fuente ---
    tabla_personal_render = datos_especificos.get('tabla_personal', [])
    tabla_personal_render_sin_total = []
    for fila in tabla_personal_render:
        perfil = fila.get('Perfil')
        cantidad = pd.to_numeric(fila.get('Cantidad'), errors='coerce')
        if perfil and cantidad > 0:
            tabla_personal_render_sin_total.append({'Perfil': perfil, 'Descripción': fila.get('Descripción', ''), 'Cantidad': int(cantidad)})

    num_personal_total_int = int(datos_especificos.get('num_personal_capacitacion', 0))
    datos_tabla_personal_word = tabla_personal_render_sin_total + [{'Perfil': 'Total', 'Descripción': '', 'Cantidad': num_personal_total_int}]
    
    tabla_detalle_personal_subdoc = create_personal_table_subdoc(
        datos_comunes['doc_tpl'], 
        headers=["Perfil", "Descripción", "Cantidad"],
        data=datos_tabla_personal_word,
        keys=['Perfil', 'Descripción', 'Cantidad'],
        column_widths=(2, 3, 1),
        texto_posterior="Elaboración: Subdirección de Sanción y Gestión de Incentivos (SSAG) - DFAI."
    )
    
    resultados_app = {'extremos': [], 'totales': {}}
    
    # --- Lógica de prorrateo (Cálculo de costos base) ---
    grupos_por_año = {}
    for extremo in datos_especificos['extremos']:
        año = extremo['fecha_incumplimiento'].year
        if año not in grupos_por_año: grupos_por_año[año] = []
        grupos_por_año[año].append(extremo)

    costos_base_prorrateados = {}
    for año, grupo_extremos in grupos_por_año.items():
        num_personal_fijo = datos_especificos.get('num_personal_capacitacion', 1)
        fecha_referencia_grupo = grupo_extremos[0]['fecha_incumplimiento']
        datos_para_ce_grupo = {**datos_comunes, 'fecha_incumplimiento': fecha_referencia_grupo}
        costo_total_grupo = calcular_costo_capacitacion(num_personal_fijo, datos_para_ce_grupo)
        if costo_total_grupo.get('error'): return costo_total_grupo
        num_extremos_en_grupo = len(grupo_extremos)
        costos_base_prorrateados[año] = {
            "precio_soles": costo_total_grupo.get('precio_base_soles_con_igv', 0) / num_extremos_en_grupo,
            "precio_dolares": costo_total_grupo.get('precio_base_dolares_con_igv', 0) / num_extremos_en_grupo,
            "ipc_costeo": costo_total_grupo.get('ipc_costeo', 0),
            "descripcion": costo_total_grupo.get('descripcion', ''),
            "ids_anexos": costo_total_grupo.get('ids_anexos', [])
        }
        if costo_total_grupo.get('ids_anexos'):
            todos_los_ids_anexos.update(costo_total_grupo.get('ids_anexos'))


    texto_hecho_principal = datos_especificos.get('texto_hecho', 'Hecho no especificado')

    # --- Bucle principal (Cálculo por extremo) ---
    for j, extremo in enumerate(datos_especificos['extremos']):
        año_extremo = extremo['fecha_incumplimiento'].year
        costo_base_prorrateado = costos_base_prorrateados.get(año_extremo)
        if not costo_base_prorrateado: continue

        ipc_incumplimiento_extremo = datos_comunes['df_indices'][datos_comunes['df_indices']['Indice_Mes'].dt.to_period('M') == pd.to_datetime(extremo['fecha_incumplimiento']).to_period('M')].iloc[0]['IPC_Mensual']
        tipo_cambio_incumplimiento_extremo = datos_comunes['df_indices'][datos_comunes['df_indices']['Indice_Mes'].dt.to_period('M') == pd.to_datetime(extremo['fecha_incumplimiento']).to_period('M')].iloc[0]['TC_Mensual']
        factor_ajuste_extremo = redondeo_excel(ipc_incumplimiento_extremo / costo_base_prorrateado['ipc_costeo'], 3) if costo_base_prorrateado['ipc_costeo'] > 0 else 0
        
        monto_soles_extremo = redondeo_excel(costo_base_prorrateado['precio_soles'] * factor_ajuste_extremo, 3)
        monto_dolares_extremo = redondeo_excel(monto_soles_extremo / tipo_cambio_incumplimiento_extremo if tipo_cambio_incumplimiento_extremo > 0 else 0, 3)
        
        ce_data_raw_extremo = [{"descripcion": costo_base_prorrateado['descripcion'], "precio_soles": costo_base_prorrateado['precio_soles'], "precio_dolares": costo_base_prorrateado['precio_dolares'], "factor_ajuste": factor_ajuste_extremo, "monto_soles": monto_soles_extremo, "monto_dolares": monto_dolares_extremo}]
        
        # Guardar para totales
        lista_ce2_data_consolidada.extend(ce_data_raw_extremo)

        ce_extremo_formatted = [{'descripcion': item.get('descripcion'), 'precio_dolares': f"US$ {item.get('precio_dolares', 0):,.3f}", 'precio_soles': f"S/ {item.get('precio_soles', 0):,.3f}", 'factor_ajuste': f"{item.get('factor_ajuste', 0):,.3f}", 'monto_soles': f"S/ {item.get('monto_soles', 0):,.3f}", 'monto_dolares': f"US$ {item.get('monto_dolares', 0):,.3f}"} for item in ce_data_raw_extremo]
        ce_extremo_formatted.append({'descripcion': 'Total', 'monto_soles': f"S/ {monto_soles_extremo:,.3f}", 'monto_dolares': f"US$ {monto_dolares_extremo:,.3f}"})
        
        tabla_ce_cuerpo_subdoc = create_table_subdoc(datos_comunes['doc_tpl'], headers=["Descripción", "Precio (US$)", "Precio (S/)", "Factor de ajuste", "Monto (S/)", "Monto (US$)"], data=ce_extremo_formatted, keys=['descripcion', 'precio_dolares', 'precio_soles', 'factor_ajuste', 'monto_soles', 'monto_dolares'])
        
        descripcion_extremo_anexo = f"Incumplimiento del {extremo['fecha_incumplimiento'].strftime('%d/%m/%Y')} (Costo Prorrateado)"
        texto_para_bi = f"{texto_hecho_principal} [Extremo {j + 1}]" 
        
        lista_extremos_para_plantilla.append({"numeral": j + 1, "descripcion": descripcion_extremo_anexo, "tabla_ce": tabla_ce_cuerpo_subdoc})
        
        nro_personal_total_general = datos_especificos.get('num_personal_capacitacion', 0)
        nro_personal_texto_anexo = f"{texto_con_numero(nro_personal_total_general)} persona{'s' if nro_personal_total_general != 1 else ''}"

        if buffer_plantilla_anexo_ce:
            anexo_tpl_loop = DocxTemplate(io.BytesIO(buffer_plantilla_anexo_ce.getvalue()))
            tabla_ce_anexo_subdoc = create_table_subdoc(anexo_tpl_loop, ["Descripción", "Precio (US$)", "Precio (S/)", "Factor de ajuste", "Monto (S/)", "Monto (US$)"], data=ce_extremo_formatted, keys=['descripcion', 'precio_dolares', 'precio_soles', 'factor_ajuste', 'monto_soles', 'monto_dolares'])
            contexto_anexo = {
                'hecho': {'numero_imputado': datos_comunes['numero_hecho_actual']},
                'extremo': {'numeral': j + 1, 
                            'tipo': descripcion_extremo_anexo, 
                            'tabla_ce': tabla_ce_anexo_subdoc},
                'fi_mes': format_date(extremo['fecha_incumplimiento'], "MMMM 'de' yyyy", locale='es'), 
                'fi_ipc': f"{ipc_incumplimiento_extremo:,.3f}", 
                'fi_tc': f"{tipo_cambio_incumplimiento_extremo:,.3f}",
                'nro_personal': nro_personal_texto_anexo, 
                'precio_dolares': f"US$ {costo_base_prorrateado.get('precio_dolares', 0):,.3f}"
            }
            anexo_tpl_loop.render({**datos_comunes['context_data'], **contexto_anexo})
            buffer_final_anexo = io.BytesIO()
            anexo_tpl_loop.save(buffer_final_anexo)
            anexos_ce_generados.append(buffer_final_anexo)

        datos_para_bi_extremo = { **datos_comunes, 'fecha_incumplimiento': extremo['fecha_incumplimiento'], 'ce_soles': monto_soles_extremo, 'ce_dolares': monto_dolares_extremo, 
                                'texto_del_hecho': texto_para_bi } 
        
        resultados_bi_extremo = calcular_beneficio_ilicito(datos_para_bi_extremo)
        if resultados_bi_extremo.get('error'): continue
        
        total_bi_uit += resultados_bi_extremo.get('beneficio_ilicito_uit', 0)
        lista_resultados_bi.append(resultados_bi_extremo)
        resultados_app['extremos'].append({'tipo': f"Fecha {extremo['fecha_incumplimiento'].strftime('%d/%m/%Y')}", 'ce_data': ce_data_raw_extremo, 'bi_data': resultados_bi_extremo.get('table_rows', []), 'bi_uit': resultados_bi_extremo.get('beneficio_ilicito_uit', 0)})

    # 5. Post-Cálculo y preparación de tablas finales
    if not lista_resultados_bi: return {'error': 'No se pudo calcular el BI para ningún extremo.'}
    
    # --- Recuperar Factor F de la interfaz ---
    factor_f = datos_especificos.get('factor_f_calculado', 1.0)

    resultados_multa_final = calcular_multa({
        **datos_comunes, 
        'beneficio_ilicito': total_bi_uit,
        'factor_f': factor_f # <--- NUEVO
    })
    multa_final_uit = resultados_multa_final.get('multa_final_uit', 0)
    
    # --- INICIO: LÓGICA DE REDUCCIÓN Y TOPE (Múltiple) ---
    datos_hecho_completos = datos_comunes.get('datos_hecho_completos', {})
    aplica_reduccion_str = datos_hecho_completos.get('aplica_reduccion', 'No')
    porcentaje_str = datos_hecho_completos.get('porcentaje_reduccion', '0%')
    multa_con_reduccion_uit = multa_final_uit
    
    if aplica_reduccion_str == 'Sí':
        reduccion_factor = 0.5 if porcentaje_str == '50%' else 0.7
        multa_con_reduccion_uit = redondeo_excel(multa_final_uit * reduccion_factor, 3)

    infraccion_info = datos_comunes['df_tipificacion'][datos_comunes['df_tipificacion']['ID_Infraccion'] == id_infraccion]
    tope_multa_uit = float('inf')
    if not infraccion_info.empty and pd.notna(infraccion_info.iloc[0].get('Tope_Multa_Infraccion')):
        tope_multa_uit = float(infraccion_info.iloc[0]['Tope_Multa_Infraccion'])

    multa_final_del_hecho_uit = min(multa_con_reduccion_uit, tope_multa_uit)
    se_aplica_tope = multa_con_reduccion_uit > tope_multa_uit
    multa_reducida_uit = multa_con_reduccion_uit if aplica_reduccion_str == 'Sí' else multa_final_uit
    # --- FIN: LÓGICA DE REDUCCIÓN Y TOPE ---

    # --- Lógica de Footnotes ---
    notas_a_mapear = {} 
    map_clave_a_texto = {} 
    datos_generales_notas = lista_resultados_bi[0].get('footnote_data', {}) if lista_resultados_bi else {}
    id_infraccion_hecho = datos_comunes.get('id_infraccion')

    for i, res_bi in enumerate(lista_resultados_bi):
        datos_notas_este_extremo = res_bi.get('footnote_data', {})
        for letra_original, clave_original in res_bi.get('footnote_mapping', {}).items():
            datos_para_formatear = {**datos_generales_notas, **datos_notas_este_extremo}
            texto_nota = obtener_fuente_formateada(clave_original, datos_para_formatear, id_infraccion_hecho)

            if texto_nota not in notas_a_mapear:
                notas_a_mapear[texto_nota] = set()
            if not texto_nota.startswith("Error: Fuente"):
                notas_a_mapear[texto_nota].add(clave_original)

            map_key = (clave_original, i)
            map_clave_a_texto[map_key] = texto_nota

    desired_key_order = ['ce_anexo', 'cok', 'periodo_bi', 'bcrp', 'ipc_fecha', 'sunat']
    mapeo_texto_a_letra_final = {}
    letra_actual_code = ord('a')
    textos_ya_mapeados = set()

    for clave in desired_key_order:
        textos_de_esta_clave = set()
        for (k, idx), txt in map_clave_a_texto.items():
            if k == clave:
                textos_de_esta_clave.add(txt)
        
        textos_ordenados = sorted(list(textos_de_esta_clave))

        for texto in textos_ordenados:
            if texto not in textos_ya_mapeados:
                letra_final = chr(letra_actual_code)
                mapeo_texto_a_letra_final[texto] = letra_final
                textos_ya_mapeados.add(texto)
                letra_actual_code += 1

    footnotes_list_bi = []
    textos_ya_agregados_a_lista = set()
    for clave in desired_key_order:
        textos_de_esta_clave = set()
        for (k, idx), txt in map_clave_a_texto.items():
            if k == clave:
                textos_de_esta_clave.add(txt)
        textos_ordenados = sorted(list(textos_de_esta_clave))
        
        for texto in textos_ordenados:
            if texto in mapeo_texto_a_letra_final and texto not in textos_ya_agregados_a_lista:
                letra = mapeo_texto_a_letra_final[texto]
                footnotes_list_bi.append(f"({letra}) {texto}")
                textos_ya_agregados_a_lista.add(texto)

    footnotes_data_bi = {'list': footnotes_list_bi, 'elaboration': 'Elaboración: SSAG - DFAI.', 'style': 'FuenteTabla'}
    
    tabla_bi_consolidada_subdoc = create_consolidated_bi_table_subdoc(
        datos_comunes['doc_tpl'], 
        lista_resultados_bi, 
        total_bi_uit, 
        footnotes_data=footnotes_data_bi,
        map_texto_a_letra=mapeo_texto_a_letra_final,
        map_clave_a_texto=map_clave_a_texto
    )
    tabla_multa_final_subdoc = create_main_table_subdoc(datos_comunes['doc_tpl'], headers=["Componentes", "Monto"], data=resultados_multa_final.get('multa_data_raw', []), keys=['Componentes', 'Monto'], column_widths=(5, 1))

    tabla_personal_data_app = datos_tabla_personal_word 

    # 6. Construcción del contexto final
    contexto_final = { 
        **datos_comunes['context_data'], **datos_especificos,
        'hecho': {
            'numero_imputado': datos_comunes['numero_hecho_actual'], 
            'descripcion': RichText(datos_especificos.get('texto_hecho', '')), 
            'lista_extremos': lista_extremos_para_plantilla,
            'tabla_bi': tabla_bi_consolidada_subdoc, 
            'tabla_multa': tabla_multa_final_subdoc
        }, 
        'numeral_hecho': f"IV.{datos_comunes['numero_hecho_actual']}",
        'fuente_cos': lista_resultados_bi[0].get('fuente_cos', '') if lista_resultados_bi else '',
        'multa_original_uit': f"{multa_final_uit:,.3f} UIT",
        'mh_uit': f"{multa_final_del_hecho_uit:,.3f} UIT",
        'bi_uit': f"{total_bi_uit:,.3f} UIT",
        'texto_explicacion_prorrateo': '',
        'tabla_detalle_personal': tabla_detalle_personal_subdoc,
        
        # --- INICIO: PLACEHOLDERS DE REDUCCIÓN Y TOPE ---
        'aplica_reduccion': aplica_reduccion_str == 'Sí',
        'porcentaje_reduccion': porcentaje_str,
        'texto_reduccion': datos_hecho_completos.get('texto_reduccion', ''),
        'memo_num': datos_hecho_completos.get('memo_num', ''),
        'memo_fecha': format_date(datos_hecho_completos.get('memo_fecha'), "d 'de' MMMM 'de' yyyy", locale='es') if datos_hecho_completos.get('memo_fecha') else '',
        'escrito_num': datos_hecho_completos.get('escrito_num', ''),
        'escrito_fecha': format_date(datos_hecho_completos.get('escrito_fecha'), "d 'de' MMMM 'de' yyyy", locale='es') if datos_hecho_completos.get('escrito_fecha') else '',
        'multa_con_reduccion_uit': f"{multa_con_reduccion_uit:,.3f} UIT",
        'se_aplica_tope': se_aplica_tope,
        'tope_multa_uit': f"{tope_multa_uit:,.3f} UIT",
        # --- FIN: PLACEHOLDERS DE REDUCCIÓN Y TOPE ---
    }
    
    buffer_plantilla = descargar_archivo_drive(id_plantilla_principal)
    if not buffer_plantilla: return {'error': "Fallo en la descarga de la plantilla para extremos."}
    
    doc_tpl_final = DocxTemplate(buffer_plantilla)
    doc_tpl_final.render(contexto_final)
    buffer_final_hecho = io.BytesIO()
    doc_tpl_final.save(buffer_final_hecho)
    
    return {
        'doc_pre_compuesto': buffer_final_hecho,
        'resultados_para_app': {
            'totales': {
                'ce2_data_raw': lista_ce2_data_consolidada,
                'beneficio_ilicito_uit': total_bi_uit, 
                'multa_final_uit': multa_final_uit, 
                'bi_data_raw': lista_resultados_bi, 
                'multa_data_raw': resultados_multa_final.get('multa_data_raw', []),
                
                # --- INICIO: DATOS DE REDUCCIÓN PARA APP ---
                'aplica_reduccion': aplica_reduccion_str,
                'porcentaje_reduccion': porcentaje_str,
                'multa_con_reduccion_uit': multa_con_reduccion_uit,
                'multa_reducida_uit': multa_reducida_uit,
                'multa_final_aplicada': multa_final_del_hecho_uit
                # --- FIN: DATOS DE REDUCCIÓN PARA APP ---
            },
            'tabla_personal_data': tabla_personal_data_app 
        },
        'texto_explicacion_prorrateo': '',
        'tabla_detalle_personal': tabla_detalle_personal_subdoc,
        'usa_capacitacion': True,
        'es_extemporaneo': False,
        'anexos_ce_generados': anexos_ce_generados,
        'ids_anexos': list(todos_los_ids_anexos), 
        'tabla_personal_data': tabla_personal_data_app,
        # --- INICIO: DATOS DE REDUCCIÓN PARA APP ---
        'aplica_reduccion': aplica_reduccion_str,
        'porcentaje_reduccion': porcentaje_str,
        'multa_reducida_uit': multa_reducida_uit
        # --- FIN: DATOS DE REDUCCIÓN PARA APP ---
    }

# --- 5. FUNCIÓN DE RUTEADOR (SIN CAMBIOS) ---
def procesar_infraccion(datos_comunes, datos_especificos):
    num_extremos = len(datos_especificos.get('extremos', []))
    if num_extremos == 0:
        return {'error': 'No se ha registrado ningún extremo para este hecho.'}
    elif num_extremos == 1:
        return _procesar_hecho_simple(datos_comunes, datos_especificos)
    else:
        return _procesar_hecho_multiple(datos_comunes, datos_especificos)