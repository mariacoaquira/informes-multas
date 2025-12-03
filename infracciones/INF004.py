import streamlit as st
import pandas as pd
import io
from babel.dates import format_date
from num2words import num2words
from docx import Document
from docxcompose.composer import Composer
from docxtpl import DocxTemplate, RichText
from datetime import date, timedelta
import holidays
from jinja2 import Environment
from textos_manager import obtener_fuente_formateada
from funciones import create_main_table_subdoc, create_table_subdoc, texto_con_numero, create_footnotes_subdoc, format_decimal_dinamico, redondeo_excel
from sheets import calcular_beneficio_ilicito, calcular_multa, descargar_archivo_drive, \
    calcular_beneficio_ilicito_extemporaneo


# --- CÁLCULO DEL COSTO EVITADO PARA INF004 ---
# Reemplaza tu antigua función de costo evitado con esta

def _calcular_costo_evitado_parcial(datos_comunes, horas_para_este_extremo, items_a_calcular, fecha_final_calculo):
    """
    Motor de Cálculo de CE para INF004.
    Recibe las HORAS ya calculadas y los ítems, y devuelve los costos.
    """
    result = {'items_calculados': [], 'error': None, 'fuente_salario': '', 'pdf_salario': '',
              'sustento_item_profesional': '', 'fuente_coti': '', 'placeholders_dinamicos': {},
              'fi_mes': '', 'fi_ipc': 0.0, 'fi_tc': 0.0}
    try:
        # 1. Unpack and validate data
        df_items_infracciones = datos_comunes.get('df_items_infracciones')
        df_costos_items = datos_comunes.get('df_costos_items')
        df_coti_general = datos_comunes.get('df_coti_general')
        df_salarios_general = datos_comunes.get('df_salarios_general')
        df_indices = datos_comunes.get('df_indices')
        id_rubro = datos_comunes.get('id_rubro_seleccionado')
        id_infraccion = datos_comunes.get('id_infraccion') # 'INF004'

        if any(df is None for df in [df_items_infracciones, df_costos_items, df_coti_general, df_salarios_general, df_indices]):
             result['error'] = "Faltan DataFrames esenciales para el cálculo del CE."
             return result

        # 2. HORAS (YA VIENEN CALCULADAS)
        # Se elimina el cálculo de horas que estaba aquí
        horas_calculadas_extremo = redondeo_excel(horas_para_este_extremo, 3)

        # 3. Get IPC/TC for the final calculation date
        # ... (Esta sección no cambia) ...
        try:
            if isinstance(fecha_final_calculo, str): fecha_final_dt = pd.to_datetime(fecha_final_calculo, errors='coerce')
            elif isinstance(fecha_final_calculo, date): fecha_final_dt = pd.to_datetime(fecha_final_calculo)
            else: fecha_final_dt = pd.NaT
            if pd.isna(fecha_final_dt): raise ValueError("Fecha final de cálculo inválida o Nula")
        except ValueError as e:
            result['error'] = f"Fecha final de cálculo inválida: {fecha_final_calculo} ({e})"
            return result
        if not pd.api.types.is_datetime64_any_dtype(df_indices['Indice_Mes']):
             try:
                 df_indices['Indice_Mes'] = pd.to_datetime(df_indices['Indice_Mes'], errors='coerce')
                 df_indices.dropna(subset=['Indice_Mes'], inplace=True)
             except Exception as e:
                  result['error'] = f"Error en formato de fechas de Índices: {e}"
                  return result
        ipc_row_inc = df_indices[df_indices['Indice_Mes'].dt.to_period('M') == fecha_final_dt.to_period('M')]
        if ipc_row_inc.empty:
            result['error'] = f"No se encontró IPC/TC para {fecha_final_dt.strftime('%B %Y')}"
            return result 
        ipc_incumplimiento = ipc_row_inc.iloc[0]['IPC_Mensual']
        tipo_cambio_incumplimiento = ipc_row_inc.iloc[0]['TC_Mensual']
        if pd.isna(ipc_incumplimiento) or pd.isna(tipo_cambio_incumplimiento) or tipo_cambio_incumplimiento == 0:
             result['error'] = f"Valores IPC/TC inválidos o faltantes para {fecha_final_dt.strftime('%B %Y')}."
             return result
        result['fi_mes'] = format_date(fecha_final_dt, "MMMM 'de' yyyy", locale='es')
        result['fi_ipc'] = float(ipc_incumplimiento)
        result['fi_tc'] = float(tipo_cambio_incumplimiento)
        # ... (Fin de la sección IPC/TC) ...

        # 4. Inicializar variables locales (no cambia)
        items_calculados_final = []
        fuente_salario_local, pdf_salario_local = '', ''
        sustentos_coti_local = []
        sustento_profesional_local = ''
        placeholders_dinamicos_local = {}
        salario_capturado = False

        # 5. Main loop - Receta INF004
        receta_df = df_items_infracciones[df_items_infracciones['ID_Infraccion'] == id_infraccion]
        if receta_df.empty:
             result['error'] = f"No se encontró receta para la infracción {id_infraccion}."
             return result

        for _, item_receta in receta_df.iterrows():
            id_item_a_buscar = item_receta['ID_Item_Infraccion']
            descripcion_insumo_receta = item_receta.get('Nombre_Item', 'N/A')

            # --- Cost Search Logic (no cambia) ---
            # ... (todo el bloque de búsqueda de costos se queda igual) ...
            posibles_costos = df_costos_items[df_costos_items['ID_Item_Infraccion'] == id_item_a_buscar].copy()
            if posibles_costos.empty: continue
            tipo_item_receta = item_receta.get('Tipo_Item'); df_candidatos = pd.DataFrame()
            if tipo_item_receta == 'Variable':
                 id_rubro_str = str(id_rubro) if id_rubro is not None else ''
                 if id_rubro_str:
                     posibles_costos['ID_Rubro'] = posibles_costos['ID_Rubro'].astype(str).fillna('')
                     df_candidatos = posibles_costos[posibles_costos['ID_Rubro'].str.contains(fr'\b{id_rubro_str}\b', regex=True, na=False)].copy()
                 if df_candidatos.empty:
                      df_candidatos = posibles_costos[posibles_costos['ID_Rubro'].isin(['', 'nan', None])].copy()
            elif tipo_item_receta == 'Fijo':
                df_candidatos = posibles_costos.copy()
            if df_candidatos.empty: continue
            fechas_fuente = []
            for _, candidato in df_candidatos.iterrows():
                id_general = candidato['ID_General']; fecha_fuente = pd.NaT
                if pd.notna(id_general):
                    if 'SAL' in id_general:
                        fuente = df_salarios_general[df_salarios_general['ID_Salario'] == id_general]
                        if not fuente.empty and 'Costeo_Salario' in fuente.columns and pd.notna(fuente.iloc[0]['Costeo_Salario']):
                             try: year_sal = int(fuente.iloc[0]['Costeo_Salario']); fecha_fuente = pd.to_datetime(f"{year_sal}-12-31", errors='coerce')
                             except (ValueError, TypeError): fecha_fuente = pd.NaT
                    elif 'COT' in id_general:
                        fuente = df_coti_general[df_coti_general['ID_Cotizacion'] == id_general]
                        if not fuente.empty and 'Fecha_Costeo' in fuente.columns and pd.notna(fuente.iloc[0]['Fecha_Costeo']):
                            fecha_fuente = pd.to_datetime(fuente.iloc[0]['Fecha_Costeo'], errors='coerce')
                fechas_fuente.append(fecha_fuente)
            df_candidatos['Fecha_Fuente'] = fechas_fuente; df_candidatos.dropna(subset=['Fecha_Fuente'], inplace=True)
            if df_candidatos.empty: continue
            fecha_final_dt_naive = fecha_final_dt.tz_localize(None) if fecha_final_dt.tzinfo is not None else fecha_final_dt
            df_candidatos['Fecha_Fuente_Naive'] = df_candidatos['Fecha_Fuente'].apply(lambda x: x.tz_localize(None) if pd.notna(x) and x.tzinfo is not None else x)
            df_candidatos['Diferencia_Dias'] = (df_candidatos['Fecha_Fuente_Naive'] - fecha_final_dt_naive).dt.days.abs()
            fila_costo_final = df_candidatos.loc[df_candidatos['Diferencia_Dias'].idxmin()]
            # --- Fin Cost Search ---

            id_general = fila_costo_final['ID_General']; fecha_fuente_dt = fila_costo_final['Fecha_Fuente']; ipc_costeo, tc_costeo = 0.0, 0.0
            if pd.notna(id_general) and 'SAL' in id_general:
                indices_del_anio = df_indices[df_indices['Indice_Mes'].dt.year == fecha_fuente_dt.year]
                if not indices_del_anio.empty: ipc_costeo = float(indices_del_anio['IPC_Mensual'].mean()); tc_costeo = float(indices_del_anio['TC_Mensual'].mean())
            elif pd.notna(id_general) and 'COT' in id_general:
                ipc_costeo_row = df_indices[df_indices['Indice_Mes'].dt.to_period('M') == fecha_fuente_dt.to_period('M')]
                if not ipc_costeo_row.empty: ipc_costeo = float(ipc_costeo_row.iloc[0]['IPC_Mensual']); tc_costeo = float(ipc_costeo_row.iloc[0]['TC_Mensual'])
            if ipc_costeo == 0 or pd.isna(ipc_costeo): continue

            # 6. Captura de placeholders (no cambia)
            # ... (se mantiene la lógica de captura de fuentes) ...
            if pd.notna(id_general):
                if 'COT' in id_general:
                    fuente_row_cot = df_coti_general[df_coti_general['ID_Cotizacion'] == id_general]
                    if not fuente_row_cot.empty: sustento_cot = fuente_row_cot.iloc[0].get('Fuente_Cotizacion'); sustentos_coti_local.append(sustento_cot) if sustento_cot else None
                elif 'SAL' in id_general and not salario_capturado:
                    fuente_row_sal = df_salarios_general[df_salarios_general['ID_Salario'] == id_general]
                    if not fuente_row_sal.empty: fuente_salario_local = fuente_row_sal.iloc[0].get('Fuente_Salario', ''); pdf_salario_local = fuente_row_sal.iloc[0].get('PDF_Salario', ''); salario_capturado = True
            if "Profesional" in descripcion_insumo_receta: sustento_profesional_local = fila_costo_final.get('Sustento_Item', '')
            try:
                 key_placeholder = f"fuente_{descripcion_insumo_receta.split()[0].lower().replace(':','')}"
                 fecha_formateada = format_date(fecha_fuente_dt, 'MMMM yyyy', locale='es').lower()
                 texto_fuente = f"{descripcion_insumo_receta}:\n{fecha_formateada}, IPC={ipc_costeo:,.3f}"
                 placeholders_dinamicos_local[key_placeholder] = texto_fuente
            except Exception as e: pass

            # 7. Cálculo de Montos (con horas_calculadas_extremo)
            try: costo_original = float(fila_costo_final['Costo_Unitario_Item'])
            except (ValueError, TypeError): costo_original = 0.0
            moneda_original = fila_costo_final['Moneda_Item']
            if moneda_original != 'S/' and (tc_costeo == 0 or pd.isna(tc_costeo)): continue
            precio_base_soles = costo_original if moneda_original == 'S/' else costo_original * tc_costeo
            factor_ajuste = redondeo_excel(ipc_incumplimiento / ipc_costeo, 3) if ipc_costeo > 0 else 0
            try: cantidad_recursos = float(item_receta.get('Cantidad_Recursos', 1.0))
            except (ValueError, TypeError): cantidad_recursos = 1.0

            # --- INICIO: CAMBIO CLAVE ---
            # 'horas_calculadas_extremo' ya es el total de horas para este extremo (ej: 24 horas)
            monto_soles = redondeo_excel(cantidad_recursos * horas_calculadas_extremo * precio_base_soles * factor_ajuste, 3)
            # --- FIN: CAMBIO CLAVE ---
            
            monto_dolares = redondeo_excel(monto_soles / tipo_cambio_incumplimiento if tipo_cambio_incumplimiento > 0 else 0, 3)
            
            items_calculados_final.append({
                "descripcion": descripcion_insumo_receta,
                "cantidad": cantidad_recursos,
                "horas": horas_calculadas_extremo, # Guardar el total de horas
                "precio_soles": precio_base_soles,
                "factor_ajuste": factor_ajuste,
                "monto_soles": monto_soles,
                "monto_dolares": monto_dolares,
                "id_anexo": fila_costo_final.get('ID_Anexo_Drive')
            })

        # 8. Update del resultado final (no cambia)
        result['items_calculados'] = items_calculados_final
        result['fuente_salario'] = fuente_salario_local
        # ... (resto de la función) ...
        result['pdf_salario'] = pdf_salario_local
        result['sustento_item_profesional'] = sustento_profesional_local
        result['fuente_coti'] = "\n".join(list(set(sustentos_coti_local)))
        result['placeholders_dinamicos'] = placeholders_dinamicos_local
        if items_calculados_final:
             result['error'] = None
        elif not result['error']:
             result['error'] = "No se generaron ítems de CE, verificar receta o costos base."
        return result

    except Exception as e:
        import traceback
        traceback.print_exc()
        result['error'] = f"Error crítico en _calcular_costo_evitado_parcial: {e}"
        return result


def renderizar_inputs_especificos(i, df_dias_no_laborables=None):
    st.markdown("##### Detalles del Requerimiento de Información")
    datos_hecho = st.session_state.imputaciones_data[i]
    if 'extremos' not in datos_hecho:
        datos_hecho['extremos'] = []

    # --- INICIO: LÓGICA DE CÁLCULO DE DÍAS HÁBILES ---
    def calcular_dias_habiles(fecha_inicio, fecha_fin, df_dnl=None):
        if not fecha_inicio or not fecha_fin or fecha_fin <= fecha_inicio:
            return 0
        feriados_pe = holidays.PE()
        dnl_set = set()
        if df_dnl is not None and 'Fecha_No_Laborable' in df_dnl.columns:
            fechas_nl = pd.to_datetime(df_dnl['Fecha_No_Laborable'], format='%d/%m/%Y', errors='coerce').dt.date
            dnl_set = set(fechas_nl.dropna())
        dias_habiles = 0
        dia_actual = fecha_inicio + timedelta(days=1)
        while dia_actual <= fecha_fin:
            es_habil = (dia_actual.weekday() < 5 and dia_actual not in feriados_pe and dia_actual not in dnl_set)
            if es_habil:
                dias_habiles += 1
            dia_actual += timedelta(days=1)
        return dias_habiles

    def calcular_fecha_incumplimiento(fecha_maxima, df_dnl=None):
        if not fecha_maxima: return None
        feriados_pe = holidays.PE()
        dnl_set = set()
        if df_dnl is not None and 'Fecha_No_Laborable' in df_dnl.columns:
            fechas_nl = pd.to_datetime(df_dnl['Fecha_No_Laborable'], format='%d/%m/%Y', errors='coerce').dt.date
            dnl_set = set(fechas_nl.dropna())
        fecha_inc = fecha_maxima
        while True:
            fecha_inc += timedelta(days=1)
            if fecha_inc.weekday() < 5 and fecha_inc not in feriados_pe and fecha_inc not in dnl_set: 
                return fecha_inc
    # --- FIN: LÓGICA DE CÁLCULO DE DÍAS HÁBILES ---

    
    # --- SECCIÓN 1: REQUERIMIENTO ORIGINAL (GLOBAL) ---
    st.markdown("###### 1. Requerimiento Original (Global)")
    datos_hecho['doc_req_num'] = st.text_input(
        "N.° de Documento (Carta, Oficio, etc.)", 
        key=f"doc_req_num_{i}", 
        value=datos_hecho.get('doc_req_num', '')
    )
    
    total_items = st.number_input("Número **total** de requerimientos de información solicitados", min_value=1, step=1,
                                  key=f"num_total_{i}", value=datos_hecho.get('num_items_solicitados', 1))
    datos_hecho['num_items_solicitados'] = total_items

    col1, col2, col3 = st.columns(3)
    with col1:
        fecha_solicitud = st.date_input("Fecha del requerimiento", key=f"fecha_sol_{i}", format="DD/MM/YYYY",
                                        value=datos_hecho.get('fecha_solicitud'))
    with col2:
        fecha_max_entrega_orig = st.date_input("Fecha máxima de entrega (Original)", min_value=fecha_solicitud, key=f"fecha_ent_orig_{i}",
                                      format="DD/MM/YYYY", value=datos_hecho.get('fecha_max_entrega_orig'))
    
    dias_habiles_orig = calcular_dias_habiles(fecha_solicitud, fecha_max_entrega_orig, df_dias_no_laborables)
    fecha_incumplimiento_orig = calcular_fecha_incumplimiento(fecha_max_entrega_orig, df_dias_no_laborables)
    
    with col3:
        st.metric(label="Plazo Original (Días Hábiles)", value=dias_habiles_orig)
        
    datos_hecho['fecha_solicitud'] = fecha_solicitud
    datos_hecho['fecha_max_entrega_orig'] = fecha_max_entrega_orig
    datos_hecho['dias_habiles_orig'] = dias_habiles_orig
    datos_hecho['fecha_incumplimiento_orig'] = fecha_incumplimiento_orig 

    st.divider()

# --- SECCIÓN 2: AMPLIACIÓN DE PLAZO (GLOBAL) ---
    st.markdown("###### 2. Ampliación de Plazo (Global)")
    aplica_ampliacion = st.radio(
        "¿Se otorgó ampliación de plazo?",
        ["No", "Sí"],
        key=f"aplica_ampliacion_{i}",
        index=0 if datos_hecho.get('aplica_ampliacion', 'No') == 'No' else 1,
        horizontal=True
    )
    datos_hecho['aplica_ampliacion'] = aplica_ampliacion
    
    dias_habiles_amp = 0
    num_items_amp = 0 # Valor por defecto
    
    if aplica_ampliacion == "Sí":
        col_amp1, col_amp2 = st.columns(2)
        with col_amp1:
            datos_hecho['doc_amp_num'] = st.text_input("N.° de Documento (Ampliación)", key=f"doc_amp_num_{i}", value=datos_hecho.get('doc_amp_num', ''))
        with col_amp2:
            datos_hecho['doc_amp_fecha'] = st.date_input("Fecha del Documento (Ampliación)", key=f"doc_amp_fecha_{i}", value=datos_hecho.get('doc_amp_fecha'), format="DD/MM/YYYY")

        col_amp3, col_amp4, col_amp5 = st.columns(3)
        with col_amp3:
            fecha_max_ampliacion = st.date_input("Nueva fecha máxima de entrega", min_value=fecha_max_entrega_orig, key=f"fecha_ent_amp_{i}", format="DD/MM/YYYY", value=datos_hecho.get('fecha_max_ampliacion'))
            datos_hecho['fecha_max_ampliacion'] = fecha_max_ampliacion
        
        # Calcular días adicionales (desde la fecha original hasta la nueva)
        fecha_inicio_calculo_amp = datos_hecho.get('doc_amp_fecha')
        dias_habiles_amp = calcular_dias_habiles(fecha_inicio_calculo_amp, fecha_max_ampliacion, df_dias_no_laborables)
        
        with col_amp4:
            # --- INICIO: CAMBIO CLAVE (NUEVO CAMPO) ---
            num_items_amp = st.number_input(
                f"N.° de ítems en ampliación", 
                min_value=1, 
                max_value=total_items, 
                value=datos_hecho.get('num_items_ampliacion', total_items), # Default: todos
                key=f"num_items_amp_{i}",
                help=f"Indique para cuántos ítems (del total de {total_items}) se otorgó esta ampliación."
            )
            datos_hecho['num_items_ampliacion'] = num_items_amp
            # --- FIN: CAMBIO CLAVE ---
            
        with col_amp5:
            st.metric(label="Plazo Adicional (Días Hábiles)", value=dias_habiles_amp)
    else:
        # Limpiar datos si el usuario cambia a "No"
        datos_hecho['doc_amp_num'] = ''
        datos_hecho['doc_amp_fecha'] = None
        datos_hecho['fecha_max_ampliacion'] = None
        # --- CORRECCIÓN ---
        # Reseteamos al total de items, no a 0, para evitar el error de min_value
        datos_hecho['num_items_ampliacion'] = total_items

    datos_hecho['dias_habiles_amp'] = dias_habiles_amp
    
    st.divider()

    # --- SECCIÓN 3: EXTREMOS DEL INCUMPLIMIENTO ---
    st.markdown("###### 3. Extremos del incumplimiento")
    
    items_asignados_total = sum(ext.get('cantidad_items', 0) for ext in datos_hecho['extremos'])
    items_restantes_total = total_items - items_asignados_total
    
    st.markdown(f"Resumen de ítems: **{items_asignados_total}** asignados / **{items_restantes_total}** restantes de un total de **{total_items}**.")
    if items_restantes_total < 0:
        st.error(f"¡Error! Se han asignado {items_asignados_total} ítems, superando el total de {total_items}.")

    boton_deshabilitado = (items_restantes_total <= 0)
    if st.button("+ Añadir Extremo", key=f"add_extremo_{i}", disabled=boton_deshabilitado):
        datos_hecho['extremos'].append({'cantidad_items': items_restantes_total, 'plazo_aplicado': 'Plazo Original'}) # Default
        st.rerun()

    
    for j, extremo in enumerate(datos_hecho['extremos']):
        with st.container(border=True):
            
            col_titulo, col_boton_eliminar = st.columns([0.85, 0.15])
            with col_titulo:
                st.markdown(f"**Extremo n.° {j + 1}**")
            with col_boton_eliminar:
                if st.button(f"🗑️", key=f"del_extremo_{i}_{j}", help="Eliminar este extremo"):
                    datos_hecho['extremos'].pop(j)
                    st.rerun()

            tipo_extremo = st.radio("Tipo de incumplimiento", 
                                    ["No remitió información / Remitió incompleto", "Remitió fuera de plazo"],
                                    key=f"tipo_extremo_{i}_{j}", 
                                    index=0 if extremo.get('tipo_extremo') == "No remitió información / Remitió incompleto" else 1 if extremo.get('tipo_extremo') == "Remitió fuera de plazo" else None)
            extremo['tipo_extremo'] = tipo_extremo
            
            items_asignados_por_otros = sum(ext.get('cantidad_items', 0) for k, ext in enumerate(datos_hecho['extremos']) if k != j)
            max_items_para_este_extremo = total_items - items_asignados_por_otros
            
            cantidad_items = st.number_input("Cantidad de ítems en este extremo", 
                                             min_value=1, 
                                             max_value=max_items_para_este_extremo,
                                             step=1, 
                                             key=f"cantidad_items_{i}_{j}",
                                             value=extremo.get('cantidad_items', 1))
            extremo['cantidad_items'] = cantidad_items
            
            # --- INICIO REQ 7: Lógica de Asignación de Plazo ---
            st.markdown("Asignación de Plazo (para este extremo)")
            
            # El radio solo muestra "Plazo Ampliado" si la ampliación fue activada
            opciones_plazo = ["Plazo Original"]
            if aplica_ampliacion == "Sí":
                opciones_plazo.append("Plazo Ampliado")
            
            plazo_aplicado = st.radio(
                "¿Qué plazo se aplica a estos ítems?",
                opciones_plazo,
                key=f"plazo_aplicado_{i}_{j}",
                index=0 if extremo.get('plazo_aplicado') == "Plazo Original" else (1 if extremo.get('plazo_aplicado') == "Plazo Ampliado" and aplica_ampliacion == "Sí" else 0),
                horizontal=True
            )
            extremo['plazo_aplicado'] = plazo_aplicado
            
            # Calcular la fecha máxima y de incumplimiento para ESTE extremo
            fecha_max_extremo = None
            if plazo_aplicado == "Plazo Ampliado":
                fecha_max_extremo = datos_hecho.get('fecha_max_ampliacion')
            else:
                fecha_max_extremo = datos_hecho.get('fecha_max_entrega_orig')
            
            fecha_inc_extremo = calcular_fecha_incumplimiento(fecha_max_extremo, df_dias_no_laborables)
            extremo['fecha_incumplimiento_extremo'] = fecha_inc_extremo # Fecha para BI
            
            # --- FIN REQ 7 ---

            if tipo_extremo == "Remitió fuera de plazo":
                extremo['fecha_extemporanea'] = st.date_input("Fecha de cumplimiento extemporáneo",
                                                              min_value=fecha_max_extremo, # Usar la fecha final del extremo
                                                              key=f"fecha_ext_{i}_{j}",
                                                              value=extremo.get('fecha_extemporanea'),
                                                              format="DD/MM/YYYY")

    st.divider()
    hubo_alegatos = st.radio("¿Hubo alegatos a la multa?", ["No", "Sí"], index=0, key=f"hubo_alegatos_{i}",
                             horizontal=True)
    datos_hecho['doc_adjunto_hecho'] = st.file_uploader("Adjuntar análisis de alegatos (.docx)", type=['docx'],
                                                        key=f"upload_analisis_{i}") if hubo_alegatos == "Sí" else None
    return datos_hecho


def validar_inputs(datos_hecho):
    """
    Valida la nueva estructura de inputs de INF004 (Req. 7).
    """
    
    # 1. Validar datos globales (Requerimiento Original)
    if not all([
        datos_hecho.get('doc_req_num'),
        datos_hecho.get('num_items_solicitados', 0) > 0,
        datos_hecho.get('fecha_solicitud'),
        datos_hecho.get('fecha_max_entrega_orig'),
        datos_hecho.get('fecha_incumplimiento_orig')
    ]):
        st.warning("Debe completar todos los campos del 'Requerimiento Original' (Sección 1).")
        return False

    # 2. Validar Ampliación (si aplica)
    if datos_hecho.get('aplica_ampliacion') == 'Sí':
        if not all([
            datos_hecho.get('doc_amp_num'),
            datos_hecho.get('doc_amp_fecha'),
            datos_hecho.get('fecha_max_ampliacion')
        ]):
            st.warning("Debe completar todos los datos de la 'Ampliación de Plazo' (Sección 2).")
            return False

    # 3. Validar que haya extremos
    if not datos_hecho.get('extremos'):
        st.warning("Debe añadir al menos un extremo (Sección 3).")
        return False
    
    # 4. Validar Total vs. Asignados
    total_items = datos_hecho.get('num_items_solicitados', 0)
    items_asignados = sum(ext.get('cantidad_items', 0) for ext in datos_hecho.get('extremos', []))
    
    if items_asignados > total_items: 
        st.warning(f"Error: Los ítems asignados ({items_asignados}) superan el total de ítems ({total_items}).")
        return False
    # (Permitir que sea menor, para "remisión incompleta")

    # 5. Validar CADA extremo
    for j, extremo in enumerate(datos_hecho.get('extremos', [])):
        
        if not all([
            extremo.get('tipo_extremo'),
            extremo.get('cantidad_items', 0) > 0,
            extremo.get('plazo_aplicado'), # Asegura que el radio 'Plazo' fue seleccionado
            extremo.get('fecha_incumplimiento_extremo') # Asegura que el cálculo del extremo se hizo
        ]):
            st.warning(f"Extremo {j+1}: Faltan datos básicos (Tipo, Cantidad, Plazo o cálculo de fecha).")
            return False
        
        if extremo.get('tipo_extremo') == "Remitió fuera de plazo" and not extremo.get('fecha_extemporanea'):
            st.warning(f"Extremo {j+1}: Debe ingresar la 'Fecha de cumplimiento extemporáneo'.")
            return False
    
    return True


def procesar_infraccion(datos_comunes, datos_hecho):
    num_extremos = len(datos_hecho.get('extremos', []))
    if num_extremos == 1:
        return _procesar_hecho_simple(datos_comunes, datos_hecho)
    elif num_extremos > 1:
        return _procesar_hecho_multiple(datos_comunes, datos_hecho)
    else:
        return {'error': 'No se ha registrado ningún extremo para este hecho.'}


def _procesar_hecho_simple(datos_comunes, datos_hecho):
    """
    Procesa un hecho con un único extremo.
    Calcula las horas según Req. 7 y las pasa al motor _calcular_costo_evitado_parcial.
    """
    try:
        # 1. Extraer datos del hecho y extremo
        id_infraccion = datos_comunes['id_infraccion']
        extremo = datos_hecho['extremos'][0]
        items_afectados = extremo.get('cantidad_items', 0)
        tipo_incumplimiento = extremo.get('tipo_extremo')
        numero_hecho = datos_comunes['numero_hecho_actual']

        # --- INICIO: LÓGICA DE HORAS (REQ. 8 - PRORRATEO PARCIAL) ---
        
        # 1. Horas Originales (Prorrateadas entre TODOS)
        num_items_total = datos_hecho.get('num_items_solicitados', 1)
        dias_habiles_orig = datos_hecho.get('dias_habiles_orig', 0)
        
        # --- CAMBIO: Redondeo de Horas Unitarias ---
        horas_item_orig_raw = (dias_habiles_orig * 8) / num_items_total if num_items_total > 0 else 0
        horas_item_orig = redondeo_excel(horas_item_orig_raw, 3) # <-- REDONDEO APLICADO
        
        # Horas originales que le corresponden a ESTE extremo
        horas_orig_del_extremo = horas_item_orig * items_afectados
        
        # 2. Horas de Ampliación (Prorrateadas SÓLO entre los ítems de la ampliación)
        horas_amp_del_extremo = 0
        
        fecha_calculo_ce = None
        fecha_incumplimiento_bi = extremo.get('fecha_incumplimiento_extremo')
        dias_habiles_amp_aplicados = 0
        
        if extremo.get('plazo_aplicado') == 'Plazo Ampliado':
            # Caso A: Este extremo SÍ tuvo ampliación
            dias_habiles_amp = datos_hecho.get('dias_habiles_amp', 0)
            num_items_en_ampliacion = datos_hecho.get('num_items_ampliacion', 1) 
            if num_items_en_ampliacion <= 0: num_items_en_ampliacion = 1
            
            # --- CAMBIO: Redondeo de Horas Unitarias (Ampliación) ---
            horas_item_amp_raw = (dias_habiles_amp * 8) / num_items_en_ampliacion
            horas_item_amp = redondeo_excel(horas_item_amp_raw, 3) # <-- REDONDEO APLICADO
            
            # Horas de ampliación que le corresponden a ESTE extremo
            horas_amp_del_extremo = horas_item_amp * items_afectados
            
            dias_habiles_amp_aplicados = dias_habiles_amp
            fecha_calculo_ce = fecha_incumplimiento_bi
        else:
            # Caso B: Este extremo NO tuvo ampliación
            fecha_calculo_ce = fecha_incumplimiento_bi
        
        # 3. Total de Horas para el Extremo
        # Ahora sumamos valores que ya vienen de un cálculo redondeado
        horas_finales_para_extremo = horas_orig_del_extremo + horas_amp_del_extremo
        
        if not fecha_calculo_ce:
            fecha_calculo_ce = fecha_incumplimiento_bi
        # --- FIN: LÓGICA DE HORAS ---

        # 2. Calcular el Costo Evitado (CE) - Pasando las horas calculadas
        # --- CAMBIO CLAVE: Se pasa el TOTAL de horas, no horas/item ---
        res_ce = _calcular_costo_evitado_parcial(
            datos_comunes, 
            horas_finales_para_extremo,  # <-- Se pasan las HORAS TOTALES del extremo (ej: 80)
            items_afectados, # Cantidad de items
            fecha_calculo_ce
        )
        if res_ce.get('error'): return {'error': res_ce['error']}
        
        ce_data_raw = res_ce.get('items_calculados', [])
        if not ce_data_raw: return {'error': "No se pudo calcular el Costo Evitado para el hecho."}

        total_soles = sum(item.get('monto_soles', 0) for item in ce_data_raw)
        total_dolares = sum(item.get('monto_dolares', 0) for item in ce_data_raw)

        # 3. Calcular BI y Multa (No cambia)
        # ... (código idéntico) ...
        texto_hecho_bi = f"{datos_hecho.get('texto_hecho', 'Hecho no especificado')}"
        datos_bi_base = { **datos_comunes, 'ce_soles': total_soles, 'ce_dolares': total_dolares, 'fecha_incumplimiento': fecha_incumplimiento_bi, 'texto_del_hecho': texto_hecho_bi }
        res_bi = None
        es_extemporaneo = (tipo_incumplimiento == "Remitió fuera de plazo")
        if es_extemporaneo:
            fecha_extemporanea = extremo.get('fecha_extemporanea')
            pre_calculo_bi = calcular_beneficio_ilicito(datos_bi_base)
            if pre_calculo_bi.get('error'): return pre_calculo_bi
            datos_bi_ext = {**datos_bi_base, 'fecha_cumplimiento_extemporaneo': fecha_extemporanea, **pre_calculo_bi}
            res_bi = calcular_beneficio_ilicito_extemporaneo(datos_bi_ext)
        else:
            res_bi = calcular_beneficio_ilicito(datos_bi_base)
        if not res_bi or res_bi.get('error'): return res_bi or {'error': 'Error desconocido al calcular el BI.'}
        beneficio_ilicito_uit = res_bi.get('beneficio_ilicito_uit', 0)
        
        # --- CORRECCIÓN: Factor de Graduación ---
        factor_f = datos_hecho.get('factor_f_calculado', 1.0)
        
        res_multa = calcular_multa({
            **datos_comunes, 
            'beneficio_ilicito': beneficio_ilicito_uit,
            'factor_f': factor_f # <--- AÑADIDO
        })
        multa_uit = res_multa.get('multa_final_uit', 0)

        # 4. Lógica de Reducción y Tope (No cambia)
        # ... (código idéntico) ...
        datos_hecho_completos = datos_comunes.get('datos_hecho_completos', {})
        aplica_reduccion_str = datos_hecho_completos.get('aplica_reduccion', 'No')
        porcentaje_str = datos_hecho_completos.get('porcentaje_reduccion', '0%')
        multa_con_reduccion_uit = multa_uit
        if aplica_reduccion_str == 'Sí':
            reduccion_factor = 0.5 if porcentaje_str == '50%' else 0.7
            multa_con_reduccion_uit = redondeo_excel(multa_uit * reduccion_factor, 3)
        infraccion_info = datos_comunes['df_tipificacion'][datos_comunes['df_tipificacion']['ID_Infraccion'] == id_infraccion]
        tope_multa_uit = float('inf')
        if not infraccion_info.empty and pd.notna(infraccion_info.iloc[0].get('Tope_Multa_Infraccion')):
            tope_multa_uit = float(infraccion_info.iloc[0]['Tope_Multa_Infraccion'])
        multa_final_del_hecho_uit = min(multa_con_reduccion_uit, tope_multa_uit)
        se_aplica_tope = multa_con_reduccion_uit > tope_multa_uit
        multa_reducida_uit = multa_con_reduccion_uit if aplica_reduccion_str == 'Sí' else multa_uit
        
        # 5. Preparar tablas y textos para Word

        # --- INICIO: Carga de Plantilla Específica (Lógica INF008) ---
        jinja_env = Environment(trim_blocks=True, lstrip_blocks=True)
        df_tipificacion = datos_comunes['df_tipificacion']
        # 'id_infraccion' ya se definió al inicio de esta función
        
        filas_inf = df_tipificacion[df_tipificacion['ID_Infraccion'] == id_infraccion]
        if filas_inf.empty: 
            return {'error': f"No se encontró ID '{id_infraccion}' en Tipificación."}
        
        fila_inf = filas_inf.iloc[0]
        id_tpl_bi = fila_inf.get('ID_Plantilla_BI') # Plantilla para hecho simple
        
        if not id_tpl_bi: 
            return {'error': f'Falta ID_Plantilla_BI para {id_infraccion} en Tipificación.'}
            
        buf_bi = descargar_archivo_drive(id_tpl_bi)
        
        if not buf_bi: 
            return {'error': f'Fallo descarga plantilla BI simple para {id_infraccion} (ID: {id_tpl_bi}).'}
        
        # Esta es la línea clave: 'doc_tpl' ahora es la plantilla BI de INF004
        doc_tpl = DocxTemplate(buf_bi) 
        # --- FIN: Carga de Plantilla Específica ---

        ce_table_formatted = []
        # ... (El resto de la creación de la tabla continúa sin cambios) ...
        for item in ce_data_raw:
            # --- INICIO: CORRECCIÓN DE SINTAXIS ---
            try: 
                horas_val = float(item.get('horas', 0))
            except (ValueError, TypeError): 
                horas_val = 0
                
            try: 
                cantidad_val = float(item.get('cantidad', 0))
            except (ValueError, TypeError): 
                cantidad_val = 0
            # --- FIN: CORRECCIÓN DE SINTAXIS ---
            descripcion_original = item.get('descripcion', ''); texto_adicional = ""
            if "Profesional" in descripcion_original: texto_adicional = "1/ "
            elif "Alquiler de laptop" in descripcion_original: texto_adicional = "2/ "
            ce_table_formatted.append({
                'descripcion': f"{descripcion_original}{texto_adicional}",
                'cantidad': format_decimal_dinamico(cantidad_val), 
                'horas': format_decimal_dinamico(horas_val), 
                'precio_soles': f"S/ {item.get('precio_soles', 0):,.3f}", 
                'factor_ajuste': f"{item.get('factor_ajuste', 0):,.3f}",
                'monto_soles': f"S/ {item.get('monto_soles', 0):,.3f}",
                'monto_dolares': f"US$ {item.get('monto_dolares', 0):,.3f}"
            })
        ce_table_formatted.append({
            'descripcion': 'Total', 'cantidad': '', 'horas': '', 'precio_soles': '', 'factor_ajuste': '',
            'monto_soles': f"S/ {total_soles:,.3f}",
            'monto_dolares': f"US$ {total_dolares:,.3f}"
        })
        tabla_ce_subdoc = create_table_subdoc(
            doc_tpl,
            headers=["Descripción", "Cantidad", "Horas", "Precio asociado (S/)", "Factor de ajuste 3/", "Monto (*) (S/)", "Monto (*) (US$) 4/"],
            data=ce_table_formatted,
            keys=['descripcion', 'cantidad', 'horas', 'precio_soles', 'factor_ajuste', 'monto_soles', 'monto_dolares']
        )
        filas_bi_crudas, footnote_mapping, datos_para_fuentes = res_bi.get('table_rows', []), res_bi.get('footnote_mapping', {}), res_bi.get('footnote_data', {})
        footnotes_list = [f"({letra}) {obtener_fuente_formateada(ref_key, datos_para_fuentes, id_infraccion=id_infraccion, es_extemporaneo=es_extemporaneo)}" for letra, ref_key in sorted(footnote_mapping.items())]
        filas_bi_para_tabla = []
        for fila in filas_bi_crudas:
            super_base = fila.get('descripcion_superindice', ''); ref_letra = fila.get('ref')
            if ref_letra: super_base += f"({ref_letra})"
            filas_bi_para_tabla.append({
                'descripcion_texto': fila.get('descripcion_texto', ''), 
                'descripcion_superindice': super_base, 
                'monto': fila.get('monto', '')
            })
        footnotes_data = {'list': footnotes_list, 'elaboration': 'Elaboración: Subdirección de Sanción y Gestión Incentivos (SSAG) - DFAI.', 'style': 'FuenteTabla'}
        tabla_bi_subdoc = create_main_table_subdoc(doc_tpl, ["Descripción", "Monto"], filas_bi_para_tabla, ['descripcion_texto', 'monto'], footnotes_data=footnotes_data, column_widths=(5, 1))
        tabla_multa_subdoc = create_main_table_subdoc(doc_tpl, ["Componentes", "Monto"], res_multa.get('multa_data_raw', []), ['Componentes', 'Monto'], texto_posterior="Elaboración: Subdirección de Sanción y Gestión Incentivos (SSAG) - DFAI.", estilo_texto_posterior='FuenteTabla', column_widths=(5, 1))
        
        # --- INICIO: LÓGICA DE TEXTO DE RAZONABILIDAD (REQ 7 - CORREGIDO) ---
        dias_plazo_texto = texto_con_numero(datos_hecho.get('dias_habiles_orig', 0), genero='m')
        total_items_texto = texto_con_numero(datos_hecho.get('num_items_solicitados', 0))
        items_afectados_texto = texto_con_numero(items_afectados)
        
        # --- CAMBIO CLAVE: 'horas_finales_para_extremo' YA ES el total de horas ---
        horas_texto_formato = texto_con_numero(horas_finales_para_extremo, genero='f') # Horas totales del extremo
        dias_equiv_texto = texto_con_numero(horas_finales_para_extremo / 8, genero='m')
        
        texto_razonabilidad = ""
        if tipo_incumplimiento == "No remitió información / Remitió incompleto":
            texto_razonabilidad = (f"Toda vez que en el presente hecho se le otorgaron {dias_plazo_texto} días para la realización de {total_items_texto} actividades; siendo que no remitió {items_afectados_texto}, por lo tanto, se considerará un (01) profesional por un periodo de {horas_texto_formato} horas de trabajo ({dias_equiv_texto} días de trabajo), ello en virtud al principio de razonabilidad.")
        elif tipo_incumplimiento == "Remitió fuera de plazo":
            texto_razonabilidad = (f"Toda vez que en el presente hecho se le otorgaron {dias_plazo_texto} días para la realización de {total_items_texto} actividades; siendo que remitió tardíamente {items_afectados_texto}, por lo tanto, se considerará un (01) profesional por un periodo de {horas_texto_formato} horas de trabajo ({dias_equiv_texto} días de trabajo), ello en virtud al principio de razonabilidad.")

        if extremo.get('plazo_aplicado') == 'Plazo Ampliado':
            dias_amp_texto = texto_con_numero(datos_hecho.get('dias_habiles_amp', 0), genero='m')
            texto_razonabilidad += f" Dicho periodo de trabajo incluye una ampliación de plazo de {dias_amp_texto} días hábiles."
        # --- FIN: LÓGICA DE TEXTO DE RAZONABILIDAD ---

        # --- INICIO: Formateo de Plazos (Req. Usuario) ---
        
        # 1. Plazos del Extremo Específico (Horas/Item)
        horas_item_final = horas_finales_para_extremo # (ej: 24)
        dias_item_final = horas_item_final / 8       # (ej: 3)
        
        # Aplicar formato (Req 1)
        ph_horas_item = f"{texto_con_numero(horas_item_final, genero='f')} horas"
        ph_dias_item = f"{texto_con_numero(dias_item_final, genero='m')} días hábiles"

        # 2. Plazos Totales (Globales)
        dias_total_global = datos_hecho.get('dias_habiles_orig', 0) + datos_hecho.get('dias_habiles_amp', 0)
        horas_total_global = dias_total_global * 8

        # Aplicar formato (Req 2)
        ph_horas_total = f"{texto_con_numero(horas_total_global, genero='f')} horas"
        ph_dias_total = f"{texto_con_numero(dias_total_global, genero='m')} días hábiles"
        
        # --- FIN: Formateo de Plazos ---

        # 6. Construcción del Contexto Final
        placeholders_dinamicos = res_ce.get('placeholders_dinamicos', {})
        fecha_max_original_fmt = format_date(datos_hecho.get('fecha_max_entrega_orig'), "d 'de' MMMM 'de' yyyy", locale='es') if datos_hecho.get('fecha_max_entrega_orig') else "N/A"
        fecha_max_final_fmt = format_date(fecha_incumplimiento_bi, "d 'de' MMMM 'de' yyyy", locale='es') if fecha_incumplimiento_bi else "N/A" # Fecha final es la del BI
        fecha_extemporanea_fmt = format_date(extremo.get('fecha_extemporanea'), "d 'de' MMMM 'de' yyyy", locale='es') if extremo.get('fecha_extemporanea') else "N/A"
        doc_amp_fecha_fmt = format_date(datos_hecho.get('doc_amp_fecha'), "d 'de' MMMM 'de' yyyy", locale='es') if datos_hecho.get('doc_amp_fecha') else ''
        
        contexto_final_word = {
            **datos_comunes['context_data'],
            **placeholders_dinamicos,
            'acronyms': datos_comunes['acronym_manager'],
            'total_items_requeridos': datos_hecho.get('num_items_solicitados', 1),
            'items_extremo_actual': items_afectados,
            'fecha_requerimiento': format_date(datos_hecho_completos.get('fecha_solicitud'), "d 'de' MMMM 'de' yyyy", locale='es') if datos_hecho_completos.get('fecha_solicitud') else "N/A",
            'hecho': {
                'numero_imputado': numero_hecho,
                'descripcion': RichText(datos_hecho.get('texto_hecho', '')),
                'tabla_ce': tabla_ce_subdoc,
                'tabla_bi': tabla_bi_subdoc,
                'tabla_multa': tabla_multa_subdoc,
            },
            'numeral_hecho': f"IV.{numero_hecho + 1}",
            'texto_condicional_razonabilidad': texto_razonabilidad,
            'multa_original_uit': f"{multa_uit:,.3f} UIT",
            'mh_uit': f"{multa_final_del_hecho_uit:,.3f} UIT",
            'bi_uit': f"{beneficio_ilicito_uit:,.3f} UIT",
            'fuente_cos': res_bi.get('fuente_cos', ''),
            'fuente_salario': res_ce.get('fuente_salario', ''),
            'pdf_salario': res_ce.get('pdf_salario', ''),
            'sustento_item_profesional': res_ce.get('sustento_item_profesional', ''),
            'fuente_coti': res_ce.get('fuente_coti', ''),
            'fi_mes': res_ce.get('fi_mes', ''), 'fi_ipc': res_ce.get('fi_ipc', 0), 'fi_tc': res_ce.get('fi_tc', 0),

            # --- INICIO: (REQ 3) PLACEHOLDERS AMPLIACIÓN (FORMATEADOS) ---
            'fecha_max_presentacion': fecha_max_final_fmt,
            'fecha_max_original': fecha_max_original_fmt,
            'fecha_extemporanea': fecha_extemporanea_fmt,
            'aplica_ampliacion': datos_hecho_completos.get('aplica_ampliacion', 'No') == 'Sí',
            'doc_req_num': datos_hecho_completos.get('doc_req_num', ''),
            'doc_amp_num': datos_hecho_completos.get('doc_amp_num', ''),
            'doc_amp_fecha': doc_amp_fecha_fmt,
            
            'plazo_final_dias_extremo': ph_dias_item,     # (Req 1)
            'plazo_final_horas_extremo': ph_horas_item,   # (Req 1)
            'plazo_total_dias': ph_dias_total,             # (Req 2)
            'plazo_total_horas': ph_horas_total,           # (Req 2)
            
            'dias_habiles_orig': f"{texto_con_numero(datos_hecho.get('dias_habiles_orig', 0), genero='m')} días hábiles",
            'dias_habiles_amp': f"{texto_con_numero(datos_hecho.get('dias_habiles_amp', 0), genero='m')} días hábiles",
            # --- FIN: (REQ 3) ---

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
        }

        # 7. Renderizar y Guardar (Cuerpo y Anexo)
        doc_tpl.render(contexto_final_word, autoescape=True, jinja_env=jinja_env)
        buf_final_hecho = io.BytesIO()
        doc_tpl.save(buf_final_hecho)

        anexos_ce_generados = []
        fila_infraccion_anexo = datos_comunes['df_tipificacion'][datos_comunes['df_tipificacion']['ID_Infraccion'] == id_infraccion]
        id_plantilla_anexo_ce = fila_infraccion_anexo.iloc[0].get('ID_Plantilla_CE')
        if id_plantilla_anexo_ce:
            buffer_anexo = descargar_archivo_drive(id_plantilla_anexo_ce)
            if buffer_anexo:
                anexo_tpl = DocxTemplate(buffer_anexo)
                anexo_tpl.render(contexto_final_word, autoescape=True, jinja_env=jinja_env) 
                buffer_final_anexo = io.BytesIO()
                anexo_tpl.save(buffer_final_anexo)
                anexos_ce_generados.append(buffer_final_anexo)

        # 8. Devolver resultados
        resultados_app = {
             'extremos': [{
                  'tipo': tipo_incumplimiento, 'ce_data': ce_data_raw, 
                  'ce_soles': total_soles, 'ce_dolares': total_dolares,
                  'bi_data': res_bi.get('table_rows', []), 'bi_uit': beneficio_ilicito_uit,
             }],
             'totales': {
                  'ce_total_soles': total_soles, 'ce_total_dolares': total_dolares,
                  'beneficio_ilicito_uit': beneficio_ilicito_uit,
                  'multa_final_uit': multa_uit, 'bi_data_raw': res_bi.get('table_rows', []),
                  'multa_data_raw': res_multa.get('multa_data_raw', []),
                  'aplica_reduccion': aplica_reduccion_str,
                  'porcentaje_reduccion': porcentaje_str,
                  'multa_con_reduccion_uit': multa_con_reduccion_uit, 
                  'multa_reducida_uit': multa_reducida_uit,
                  'multa_final_aplicada': multa_final_del_hecho_uit 
             }
        }
        
        return {
            'contexto_final_word': contexto_final_word, 
            'doc_pre_compuesto': buf_final_hecho,
            'resultados_para_app': resultados_app,
            'es_extemporaneo': es_extemporaneo,
            'usa_capacitacion': False, 
            'anexos_ce_generados': anexos_ce_generados,
            'ids_anexos': list(filter(None, set(item.get('id_anexo') for item in ce_data_raw))),
            'tabla_detalle_personal': None,
            'tabla_personal_data': []
        }
    
    except Exception as e:
        import traceback
        traceback.print_exc()
        return {'error': f"Error crítico en _procesar_hecho_simple INF004: {e}"}
    

# ---------------------------------------------------------------------
# FUNCIÓN 6: PROCESAR HECHO MÚLTIPLE (INF004 - CORREGIDA CON SUPERÍNDICES)
# ---------------------------------------------------------------------
def _procesar_hecho_multiple(datos_comunes, datos_hecho):
    """
    Procesa INF004 con múltiples extremos, usando la lógica Req. 7.
    """
    try:
        # --- 1. Cargar Plantillas (No cambia) ---
        df_tipificacion, id_infraccion = datos_comunes['df_tipificacion'], datos_comunes['id_infraccion']
        filas_inf = df_tipificacion[df_tipificacion['ID_Infraccion'] == id_infraccion]
        if filas_inf.empty: return {'error': f"No se encontró ID '{id_infraccion}' en Tipificación."}
        fila_inf = filas_inf.iloc[0]
        id_tpl_principal = fila_inf.get('ID_Plantilla_BI_Extremo')
        id_tpl_anx = fila_inf.get('ID_Plantilla_CE_Extremo')
        if not id_tpl_principal or not id_tpl_anx:
             return {'error': f'Faltan IDs de plantilla (BI_Extremo o CE_Extremo) para {id_infraccion}.'}
        buffer_plantilla = descargar_archivo_drive(id_tpl_principal)
        buffer_anexo = descargar_archivo_drive(id_tpl_anx)
        if not buffer_plantilla or not buffer_anexo:
            return {'error': f'Fallo descarga plantilla BI Extremo {id_tpl_principal} o anexo {id_tpl_anx}.'}
        jinja_env = Environment(trim_blocks=True, lstrip_blocks=True)
        tpl_principal = DocxTemplate(buffer_plantilla)
        # --- FIN CARGA PLANTILLAS ---

        # 2. Inicializar acumuladores y datos globales
        total_ce_soles = 0.0; total_ce_dolares = 0.0; total_bi_uit = 0.0
        lista_bi_resultados_completos = [] 
        anexos_ids = set()
        num_hecho = datos_comunes['numero_hecho_actual']
        anexos_ce = []
        lista_extremos_plantilla_word = []
        ce_total_raw_para_app = []
        resultados_app = {'extremos': [], 'totales': {'ce_total_soles': 0.0, 'ce_total_dolares': 0.0}, 'bi_data_raw': []}

        # --- INICIO: LÓGICA DE HORAS (REQ 8 - PRORRATEO PARCIAL GLOBAL) ---
        num_items_total = datos_hecho.get('num_items_solicitados', 1)
        dias_habiles_orig = datos_hecho.get('dias_habiles_orig', 0)
        
        # 1. Horas Originales (Prorrateadas entre TODOS)
        # --- CAMBIO: Redondeo de Horas Unitarias ---
        horas_item_orig_raw = (dias_habiles_orig * 8) / num_items_total if num_items_total > 0 else 0
        horas_item_orig = redondeo_excel(horas_item_orig_raw, 3) # <-- REDONDEO APLICADO
        
        # 2. Horas de Ampliación (Prorrateadas SÓLO entre los ítems de la ampliación)
        dias_habiles_amp = datos_hecho.get('dias_habiles_amp', 0)
        num_items_en_ampliacion = datos_hecho.get('num_items_ampliacion', 1) 
        if num_items_en_ampliacion <= 0: num_items_en_ampliacion = 1
        
        horas_item_amp = 0
        if dias_habiles_amp > 0:
            # --- CAMBIO: Redondeo de Horas Unitarias (Ampliación) ---
            horas_item_amp_raw = (dias_habiles_amp * 8) / num_items_en_ampliacion
            horas_item_amp = redondeo_excel(horas_item_amp_raw, 3) # <-- REDONDEO APLICADO
        # --- FIN: LÓGICA DE HORAS ---

        # 3. Iterar sobre cada extremo para PREPARAR DATOS
        for j, extremo in enumerate(datos_hecho['extremos']):
            
            # a. Calcular Horas para ESTE extremo (CORREGIDO)
            items_afectados = extremo.get('cantidad_items', 0)
            tipo_incumplimiento = extremo.get('tipo_extremo')
            
            fecha_calculo_ce = None
            fecha_incumplimiento_bi = extremo.get('fecha_incumplimiento_extremo')
            fecha_max_entrega_final_extremo = None 
            
            # Horas originales que le corresponden a ESTE extremo
            horas_orig_del_extremo = horas_item_orig * items_afectados
            horas_amp_del_extremo = 0
            
            if extremo.get('plazo_aplicado') == 'Plazo Ampliado':
                # Caso A: Este extremo SÍ tuvo ampliación
                # Horas de ampliación que le corresponden a ESTE extremo
                horas_amp_del_extremo = horas_item_amp * items_afectados
                
                fecha_calculo_ce = fecha_incumplimiento_bi
                fecha_max_entrega_final_extremo = datos_hecho.get('fecha_max_ampliacion')
            else:
                # Caso B: Este extremo NO tuvo ampliación
                # horas_amp_del_extremo se queda en 0
                fecha_calculo_ce = fecha_incumplimiento_bi
                fecha_max_entrega_final_extremo = datos_hecho.get('fecha_max_entrega_orig')
            
            # --- CAMBIO CLAVE: CÁLCULO DE HORAS TOTALES ---
            horas_totales_para_extremo = horas_orig_del_extremo + horas_amp_del_extremo
            
            if not fecha_calculo_ce: fecha_calculo_ce = fecha_incumplimiento_bi
            
            # b. Calcular CE del extremo (Pasando horas TOTALES)
            res_ce_parcial = _calcular_costo_evitado_parcial(
                datos_comunes, 
                horas_totales_para_extremo, # <-- Pasa el TOTAL de horas del extremo
                items_afectados, 
                fecha_calculo_ce
            )
            if res_ce_parcial.get('error'): st.warning(f"Error CE Extremo {j+1}: {res_ce_parcial['error']}. Saltando."); continue
            
            ce_parcial_raw = res_ce_parcial.get('items_calculados', [])
            if not ce_parcial_raw: st.warning(f"No se generaron ítems CE para extremo {j+1}. Saltando."); continue
            ce_soles_parcial = sum(item.get('monto_soles', 0) for item in ce_parcial_raw)
            ce_dolares_parcial = sum(item.get('monto_dolares', 0) for item in ce_parcial_raw)

            # c. Calcular BI del extremo
            es_extemporaneo_extremo = (tipo_incumplimiento == "Remitió fuera de plazo")
            fecha_extemporanea = extremo.get('fecha_extemporanea') if es_extemporaneo_extremo else None
            texto_hecho_bi = f"{datos_hecho.get('texto_hecho', 'Hecho no especificado')} - Extremo {j + 1}"
            datos_bi_base = {**datos_comunes, 'ce_soles': ce_soles_parcial, 'ce_dolares': ce_dolares_parcial, 'fecha_incumplimiento': fecha_incumplimiento_bi, 'texto_del_hecho': texto_hecho_bi}
            res_bi_parcial = None
            if es_extemporaneo_extremo:
                pre_calculo_bi = calcular_beneficio_ilicito(datos_bi_base)
                if pre_calculo_bi.get('error'): st.warning(f"Error pre-BI Extremo {j+1}: {pre_calculo_bi['error']}. Saltando."); continue
                datos_bi_ext = {**datos_bi_base, 'fecha_cumplimiento_extemporaneo': fecha_extemporanea, **pre_calculo_bi}; res_bi_parcial = calcular_beneficio_ilicito_extemporaneo(datos_bi_ext)
            else: res_bi_parcial = calcular_beneficio_ilicito(datos_bi_base)
            if not res_bi_parcial or res_bi_parcial.get('error'): st.warning(f"Error BI Extremo {j+1}: {res_bi_parcial.get('error', 'Error')}. Saltando."); continue

            # d. Acumular totales
            bi_parcial_uit = res_bi_parcial.get('beneficio_ilicito_uit', 0.0)
            total_ce_soles += ce_soles_parcial; total_ce_dolares += ce_dolares_parcial; total_bi_uit += bi_parcial_uit
            for item in ce_parcial_raw: anexos_ids.add(item.get('id_anexo'))
            ce_total_raw_para_app.extend(ce_parcial_raw)
            lista_bi_resultados_completos.append(res_bi_parcial) 
            resultados_app['extremos'].append({ 'tipo': tipo_incumplimiento, 'ce_data': ce_parcial_raw, 'ce_soles': ce_soles_parcial, 'ce_dolares': ce_dolares_parcial, 'bi_data': res_bi_parcial.get('table_rows', []), 'bi_uit': bi_parcial_uit })
            resultados_app['totales']['ce_total_soles'] = total_ce_soles; resultados_app['totales']['ce_total_dolares'] = total_ce_dolares

            # e. Generar Anexo CE del extremo
            tpl_anx_loop = DocxTemplate(io.BytesIO(buffer_anexo.getvalue()))
            ce_anexo_formatted = []
            for item in ce_parcial_raw:
                descripcion_original = item.get('descripcion', ''); texto_adicional = ""
                if "Profesional" in descripcion_original: texto_adicional = "1/ "
                elif "Alquiler de laptop" in descripcion_original: texto_adicional = "2/ "
                ce_anexo_formatted.append({
                    'descripcion': f"{descripcion_original}{texto_adicional}", 
                    'cantidad': format_decimal_dinamico(item.get('cantidad', 0)), 
                    'horas': format_decimal_dinamico(item.get('horas', 0)), 
                    'precio_soles': f"S/ {item.get('precio_soles', 0):,.3f}", 
                    'factor_ajuste': f"{item.get('factor_ajuste', 0):,.3f}", 
                    'monto_soles': f"S/ {item.get('monto_soles', 0):,.3f}", 
                    'monto_dolares': f"US$ {item.get('monto_dolares', 0):,.3f}"
                })
            ce_anexo_formatted.append({'descripcion': 'Total', 'monto_soles': f"S/ {ce_soles_parcial:,.3f}", 'monto_dolares': f"US$ {ce_dolares_parcial:,.3f}"})
            tabla_ce_anexo_subdoc = create_table_subdoc( tpl_anx_loop, ["Descripción", "Cantidad", "Horas", "Precio asociado (S/)", "Factor de ajuste", "Monto (S/)", "Monto (US$)"], ce_anexo_formatted, ['descripcion', 'cantidad', 'horas', 'precio_soles', 'factor_ajuste', 'monto_soles', 'monto_dolares'] )
            
            contexto_anexo_extremo = { 
                **datos_comunes['context_data'], 
                **(res_ce_parcial.get('placeholders_dinamicos', {})), 
                'acronyms': datos_comunes['acronym_manager'], 
                'hecho': {'numero_imputado': num_hecho}, 
                'extremo': {
                    'numeral': j+1, 'tipo': tipo_incumplimiento, 
                    'fecha_incumplimiento': format_date(fecha_incumplimiento_bi, "d 'de' MMMM 'de' yyyy", locale='es'), 
                    'fecha_max_presentacion': format_date(fecha_max_entrega_final_extremo, "d 'de' MMMM 'de' yyyy", locale='es') if fecha_max_entrega_final_extremo else "N/A",
                    'fecha_max_original': format_date(datos_hecho.get('fecha_max_entrega_orig'), "d 'de' MMMM 'de' yyyy", locale='es') if datos_hecho.get('fecha_max_entrega_orig') else "N/A",
                    'fecha_extemporanea': format_date(fecha_extemporanea, "d 'de' MMMM 'de' yyyy", locale='es') if fecha_extemporanea else "N/A"
                }, 
                'tabla_ce': tabla_ce_anexo_subdoc, 'fuente_salario': res_ce_parcial.get('fuente_salario', ''), 'pdf_salario': res_ce_parcial.get('pdf_salario', ''), 'sustento_item_profesional': res_ce_parcial.get('sustento_item_profesional', ''), 'fuente_coti': res_ce_parcial.get('fuente_coti', ''), 'fi_mes': res_ce_parcial.get('fi_mes', ''), 'fi_ipc': res_ce_parcial.get('fi_ipc', 0), 'fi_tc': res_ce_parcial.get('fi_tc', 0)
            }
            tpl_anx_loop.render(contexto_anexo_extremo, autoescape=True, jinja_env=jinja_env); buf_anx_final = io.BytesIO(); tpl_anx_loop.save(buf_anx_final); anexos_ce.append(buf_anx_final)

            # f. Generar tablas y texto para el CUERPO
            tabla_ce_cuerpo = create_table_subdoc( tpl_principal, ["Descripción", "Cantidad", "Horas", "Precio asociado (S/)", "Factor de ajuste", "Monto (S/)", "Monto (US$)"], ce_anexo_formatted, ['descripcion', 'cantidad', 'horas', 'precio_soles', 'factor_ajuste', 'monto_soles', 'monto_dolares'] )
            filas_bi_crudas_ext, fn_map_ext, fn_data_ext = res_bi_parcial.get('table_rows', []), res_bi_parcial.get('footnote_mapping', {}), res_bi_parcial.get('footnote_data', {})
            fn_list_ext = [f"({l}) {obtener_fuente_formateada(k, fn_data_ext, id_infraccion, es_extemporaneo_extremo)}" for l, k in sorted(fn_map_ext.items())]
            fn_data_dict_ext = {'list': fn_list_ext, 'elaboration': 'Elaboración: Subdirección de Sanción y Gestión de Incentivos (SSAG) - DFAI.', 'style': 'FuenteTabla'}
            filas_bi_con_superindice = []
            for fila in filas_bi_crudas_ext:
                nueva_fila = fila.copy(); ref_letra = nueva_fila.get('ref')
                texto_base = str(nueva_fila.get('descripcion_texto', nueva_fila.get('descripcion', '')))
                super_existente = str(nueva_fila.get('descripcion_superindice', ''))
                if ref_letra: super_existente += f"({ref_letra})"
                nueva_fila['descripcion_texto'] = texto_base; nueva_fila['descripcion_superindice'] = super_existente
                filas_bi_con_superindice.append(nueva_fila)
            tabla_bi_cuerpo = create_main_table_subdoc(
                tpl_principal, ["Descripción", "Monto"], filas_bi_con_superindice,
                keys=['descripcion_texto', 'monto'], footnotes_data=fn_data_dict_ext, column_widths=(5, 1)
            )

            # --- INICIO: LÓGICA DE TEXTO DE RAZONABILIDAD (REQ 7 - CORREGIDO) ---
            dias_plazo_texto_orig = texto_con_numero(datos_hecho.get('dias_habiles_orig', 0), genero='m')
            total_items_texto = texto_con_numero(datos_hecho.get('num_items_solicitados', 0))
            items_afectados_texto = texto_con_numero(items_afectados)
            
            # --- CAMBIO CLAVE: 'horas_totales_para_extremo' YA ES el total de horas ---
            horas_texto_formato_ext = texto_con_numero(horas_totales_para_extremo, genero='f')
            dias_equiv_texto_ext = texto_con_numero(horas_totales_para_extremo / 8, genero='m')

            texto_razonabilidad_parcial = ""
            if tipo_incumplimiento == "No remitió información / Remitió incompleto":
                texto_razonabilidad_parcial = (f"Toda vez que en el presente hecho se le otorgaron {dias_plazo_texto_orig} días para la realización de {total_items_texto} actividades; siendo que no remitió {items_afectados_texto}, por lo tanto, se considerará un (01) profesional por un periodo de {horas_texto_formato_ext} horas de trabajo ({dias_equiv_texto_ext} días de trabajo), ello en virtud al principio de razonabilidad.")
            elif tipo_incumplimiento == "Remitió fuera de plazo":
                texto_razonabilidad_parcial = (f"Toda vez que en el presente hecho se le otorgaron {dias_plazo_texto_orig} días para la realización de {total_items_texto} actividades; siendo que remitió tardíamente {items_afectados_texto}, por lo tanto, se considerará un (01) profesional por un periodo de {horas_texto_formato_ext} horas de trabajo ({dias_equiv_texto_ext} días de trabajo), ello en virtud al principio de razonabilidad.")
            
            if extremo.get('plazo_aplicado') == 'Plazo Ampliado':
                dias_amp_texto = texto_con_numero(datos_hecho.get('dias_habiles_amp', 0), genero='m')
                texto_razonabilidad_parcial += f" Dicho periodo de trabajo incluye una ampliación de plazo de {dias_amp_texto} días hábiles."
            # --- FIN: LÓGICA DE TEXTO DE RAZONABILIDAD ---

            # --- INICIO: Formateo de Plazos (Req. Usuario - Bucle) ---
            horas_item_final_loop = horas_totales_para_extremo # Horas/Item (e.g., 8 or 24)
            dias_item_final_loop = horas_item_final_loop / 8
            
            ph_horas_item_loop = f"{texto_con_numero(horas_item_final_loop, genero='f')} horas"
            ph_dias_item_loop = f"{texto_con_numero(dias_item_final_loop, genero='m')} días hábiles"
            # --- FIN: Formateo de Plazos ---

            lista_extremos_plantilla_word.append({
                'loop_index': j + 1,
                'numeral': f"{num_hecho}.{j + 1}",
                'descripcion': f"Cálculo para los {items_afectados_texto} ítems '{tipo_incumplimiento}'.",
                'tabla_ce': tabla_ce_cuerpo, 'tabla_bi': tabla_bi_cuerpo,
                'bi_uit_extremo': f"{bi_parcial_uit:,.3f} UIT",
                'texto_razonabilidad': RichText(texto_razonabilidad_parcial),
                'fecha_max_presentacion': format_date(fecha_max_entrega_final_extremo, "d 'de' MMMM 'de' yyyy", locale='es') if fecha_max_entrega_final_extremo else "N/A",
                'fecha_max_original': format_date(datos_hecho.get('fecha_max_entrega_orig'), "d 'de' MMMM 'de' yyyy", locale='es') if datos_hecho.get('fecha_max_entrega_orig') else "N/A",
                'fecha_extemporanea': format_date(fecha_extemporanea, "d 'de' MMMM 'de' yyyy", locale='es') if fecha_extemporanea else "N/A",
                'items_extremo_actual': items_afectados,
                
                # --- INICIO: PLACEHOLDERS REQUERIDOS (FORMATEADOS) ---
                'plazo_final_dias_extremo': ph_dias_item_loop,
                'plazo_final_horas_extremo': ph_horas_item_loop,
                # --- FIN: PLACEHOLDERS REQUERIDOS ---
            })
        # --- FIN DEL BUCLE DE EXTREMOS ---

        # 5. Post-Cálculo: Multa Final
        if not lista_bi_resultados_completos: return {'error': 'No se pudo calcular BI para ningún extremo.'}
        
        # --- CORRECCIÓN: Factor de Graduación ---
        factor_f = datos_hecho.get('factor_f_calculado', 1.0)
        
        res_multa_final = calcular_multa({
            **datos_comunes, 
            'beneficio_ilicito': total_bi_uit,
            'factor_f': factor_f # <--- AÑADIDO
        })
        multa_final_uit = res_multa_final.get('multa_final_uit', 0.0)       
        
        # ... (El resto de la función: Lógica de Reducción, Contexto Final, Return) ...
        # (El código de reducción y los contextos que ya te di se mantienen igual)
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
        tabla_multa_final_subdoc = create_main_table_subdoc( tpl_principal, ["Componentes", "Monto"], res_multa_final.get('multa_data_raw', []), ['Componentes', 'Monto'], texto_posterior="Elaboración: Subdirección de Sanción y Gestión de Incentivos (SSAG) - DFAI.", estilo_texto_posterior='FuenteTabla', column_widths=(5, 1) )


        # --- INICIO: Formateo de Plazos (Req. Usuario - Global) ---
        dias_habiles_orig_global = datos_hecho.get('dias_habiles_orig', 0)
        dias_habiles_amp_global = datos_hecho.get('dias_habiles_amp', 0)
        dias_total_global = dias_habiles_orig_global + dias_habiles_amp_global
        horas_total_global = dias_total_global * 8

        ph_horas_total_global = f"{texto_con_numero(horas_total_global, genero='f')} horas"
        ph_dias_total_global = f"{texto_con_numero(dias_total_global, genero='m')} días hábiles"
        ph_dias_orig_global = f"{texto_con_numero(dias_habiles_orig_global, genero='m')} días hábiles"
        ph_dias_amp_global = f"{texto_con_numero(dias_habiles_amp_global, genero='m')} días hábiles"
        # --- FIN: Formateo de Plazos ---

        # 6. Contexto Final y Renderizado
        contexto_final = {
            **datos_comunes['context_data'],
            'acronyms': datos_comunes['acronym_manager'],
            'total_items_requeridos': datos_hecho.get('num_items_solicitados', 1),
            'fecha_requerimiento': format_date(datos_hecho.get('fecha_solicitud'), "d 'de' MMMM 'de' yyyy", locale='es') if datos_hecho.get('fecha_solicitud') else "N/A",
            'hecho': {
                'numero_imputado': num_hecho,
                'descripcion': RichText(datos_hecho.get('texto_hecho', '')),
                'lista_extremos': lista_extremos_plantilla_word,
             },
            'numeral_hecho': f"IV.{num_hecho + 1}",
            'bi_uit_total': f"{total_bi_uit:,.3f} UIT",
            'multa_original_uit': f"{multa_final_uit:,.3f} UIT",
            'mh_uit': f"{multa_final_del_hecho_uit:,.3f} UIT",
            'tabla_multa_final': tabla_multa_final_subdoc,
            'texto_explicacion_prorrateo': '',
            
            # --- INICIO: (REQ 3) PLACEHOLDERS AMPLIACIÓN (FORMATEADOS) ---
            'aplica_ampliacion': any(ext.get('plazo_aplicado') == 'Plazo Ampliado' for ext in datos_hecho['extremos']),
            'doc_req_num': datos_hecho.get('doc_req_num', ''),
            'doc_amp_num': next((ext.get('doc_amp_num') for ext in datos_hecho['extremos'] if ext.get('doc_amp_num')), ''), # Primer N° de ampliación
            'doc_amp_fecha': format_date(next((ext.get('doc_amp_fecha') for ext in datos_hecho['extremos'] if ext.get('doc_amp_fecha')), None), "d 'de' MMMM 'de' yyyy", locale='es'),
            
            'plazo_total_dias': ph_dias_total_global,
            'plazo_total_horas': ph_horas_total_global,
            
            'dias_habiles_orig': ph_dias_orig_global,
            'dias_habiles_amp': ph_dias_amp_global,
            # --- FIN: (REQ 3) ---
            
            'aplica_reduccion': aplica_reduccion_str == 'Sí',
            'porcentaje_reduccion': porcentaje_str,
            'texto_reduccion': datos_hecho_completos.get('texto_reduccion', ''),
            'memo_num': datos_hecho_completos.get('memo_num', ''),
            'memo_fecha': format_date(datos_hecho_completos.get('memo_fecha'), "d 'de' MMMM 'de' yyyy", locale='es') if datos_hecho_completos.get('memo_fecha') else '',
            'escrito_num': datos_hecho_completos.get('escrito_num', ''),
            'escrito_fecha': format_date(datos_hecho_completos.get('escrito_fecha'), "d 'de' MMMM 'de' yyyy", locale='es') if datos_hecho_completos.get('escrito_fecha') else '',
            'multa_con_reduccion_uit': f"{multa_con_reduccion_uit:,.3f} UIT",
            'se_aplica_tope': se_aplica_tope,
            'tope_multa_uit': f"{tope_multa_uit:,.3f} UIT"
        }
        
        tpl_principal.render(contexto_final, autoescape=True, jinja_env=jinja_env);
        buf_final = io.BytesIO(); tpl_principal.save(buf_final)

        # 7. Preparar datos para App
        resultados_app['totales'] = {
            **resultados_app['totales'], 
            'beneficio_ilicito_uit': total_bi_uit, 
            'multa_data_raw': res_multa_final.get('multa_data_raw', []), 
            'multa_final_uit': multa_final_uit, 
            'bi_data_raw': lista_bi_resultados_completos,
            'aplica_reduccion': aplica_reduccion_str,
            'porcentaje_reduccion': porcentaje_str,
            'multa_con_reduccion_uit': multa_con_reduccion_uit, 
            'multa_reducida_uit': multa_reducida_uit,
            'multa_final_aplicada': multa_final_del_hecho_uit 
        }

        # 8. Devolver resultados
        return {
            'doc_pre_compuesto': buf_final,
            'resultados_para_app': resultados_app,
            'texto_explicacion_prorrateo': '',
            'tabla_detalle_personal': None,
            'usa_capacitacion': False,
            'es_extemporaneo': any(ext.get('tipo_extremo') == 'Remitió fuera de plazo' for ext in datos_hecho['extremos']),
            'anexos_ce_generados': anexos_ce,
            'ids_anexos': list(filter(None, anexos_ids)),
            'tabla_personal_data': [],
            'aplica_reduccion': aplica_reduccion_str,
            'porcentaje_reduccion': porcentaje_str,
            'multa_reducida_uit': multa_reducida_uit
        }
    except Exception as e:
        import traceback; traceback.print_exc()
        try: st.error(f"Error fatal en _procesar_hecho_multiple INF004: {e}")
        except ImportError: print(f"Error fatal en _procesar_hecho_multiple INF004: {e}")
        return {'error': f"Error fatal en _procesar_hecho_multiple INF004: {e}"}