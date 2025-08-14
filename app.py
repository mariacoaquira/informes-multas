import streamlit as st
import io
import locale
from datetime import datetime, date, timedelta
from dateutil.relativedelta import relativedelta
import pandas as pd
import holidays
from num2words import num2words
import mammoth                            
from docx import Document                   
from docxcompose.composer import Composer 
from docxtpl import DocxTemplate, RichText
import importlib

from sheets import (conectar_gsheet, cargar_hoja_a_df, RUTA_CREDENCIALES_GCP, 
                    NOMBRE_GSHEET_ASIGNACIONES, NOMBRE_GSHEET_MAESTRO, 
                    get_person_details_by_base_name, descargar_archivo_drive,
                    calcular_costo_evitado, calcular_beneficio_ilicito, calcular_multa)

from funciones import (
    combinar_con_composer, 
    create_table_subdoc, 
    create_main_table_subdoc,
    create_summary_table_subdoc
)


# --------------------------------------------------------------------
#  INICIALIZACIÓN DE LA APLICACIÓN
# --------------------------------------------------------------------
st.set_page_config(layout="wide", page_title="Asistente de Multas")
st.title("🤖 Asistente para Generación de Informes de Multa")

# Actualizamos la lista de claves en session_state
for key in [
    'info_expediente', 'rubro_seleccionado', 'id_rubro_seleccionado', 'tipo_infraccion_seleccionado',
    'subtipo_infraccion_seleccionado', 'id_infraccion_actual', 'dias_habiles_plazo',
    'fecha_incumplimiento', 'costo_evitado_total_soles', 'costo_evitado_total_dolares',
    'paso3_completo', 'beneficio_ilicito_uit',
    'df_final_ce_data', 'lista_items_infraccion_usados', 'df_resumen_bi_data', 'df_multa_data',
    'multa_final_uit', 'num_expediente_formateado', 'id_cotizacion_activa',
    'fi_ipc_valor', 'fi_tc_valor',
    'numero_rsd', 'fecha_rsd',
    'texto_hecho_imputado', 'doc_adjunto_file'
]:
    if key not in st.session_state:
        st.session_state[key] = None

cliente_gspread = conectar_gsheet() # Sin argumentos

# --------------------------------------------------------------------
#  CUERPO DE LA APLICACIÓN
# --------------------------------------------------------------------
if cliente_gspread:

    # --- PASO 1: BÚSQUEDA DE EXPEDIENTE ---
    st.header("Paso 1: Búsqueda del Expediente")
    col1, col2 = st.columns([1, 2])
    with col1:
        try:
            locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
        except locale.Error:
            locale.setlocale(locale.LC_TIME, '')

        fecha_actual = datetime.now()
        fecha_anterior = fecha_actual - relativedelta(months=1)
        nombre_hoja_actual = fecha_actual.strftime("%B %Y").capitalize()
        nombre_hoja_anterior = fecha_anterior.strftime("%B %Y").capitalize()
        hojas_disponibles = [nombre_hoja_actual, nombre_hoja_anterior]
        mes_seleccionado = st.selectbox("Selecciona el mes de la asignación:", options=hojas_disponibles)

    df_asignaciones = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_ASIGNACIONES, mes_seleccionado)

    with col2:
        if df_asignaciones is not None:
            num_expediente_simple = st.text_input("Ingresa el N° de Expediente (formato XXXX-XXXX):",
                                                  placeholder="Ej: 3345-2023")
            if st.button("Buscar Expediente", type="primary"):
                for key in st.session_state.keys():
                    if key != 'info_expediente':
                        st.session_state[key] = None

                if num_expediente_simple:
                    try:
                        if "-" not in num_expediente_simple or len(num_expediente_simple.split('-')) != 2:
                            raise ValueError("El formato debe ser XXXX-XXXX.")
                        num_formateado = f"{num_expediente_simple}-OEFA/DFAI/PAS"
                        st.session_state.num_expediente_formateado = num_formateado
                        resultado = df_asignaciones[df_asignaciones['EXPEDIENTE'] == num_formateado]
                        if not resultado.empty:
                            st.success(f"¡Expediente '{num_formateado}' encontrado!")
                            st.session_state.info_expediente = resultado.iloc[0].to_dict()
                            if 'imputaciones_data' not in st.session_state:
                                num_imputaciones = int(st.session_state.info_expediente.get('IMPUTACIONES', 1))
                                # Crea la lista para guardar la información de cada hecho.
                                st.session_state.imputaciones_data = [{} for _ in range(num_imputaciones)]
                                st.info(f"Se han preparado {num_imputaciones} bloque(s) para los hechos imputados.")
                        else:
                            st.error(
                                f"No se encontró el expediente '{num_formateado}' en la hoja de '{mes_seleccionado}'.")
                            st.session_state.info_expediente = None
                    except ValueError as e:
                        st.error(f"Error en el formato: {e}")
                else:
                    st.warning("Por favor, ingresa un número de expediente.")
    st.write("---")

    # --- PASO 2: DETALLES DEL EXPEDIENTE ---
    if st.session_state.info_expediente:
        st.header("Paso 2: Detalles del Expediente")
        info_caso = st.session_state.info_expediente

        st.subheader("Datos del Expediente")
        nombre_completo_analista = "No encontrado"
        df_analistas = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Analistas")
        if df_analistas is not None:
            nombre_base_analista = info_caso.get('ANALISTA ECONÓMICO')
            if nombre_base_analista:
                analista_encontrado = df_analistas[df_analistas['Nombre_Base_Analista'] == nombre_base_analista]
                if not analista_encontrado.empty:
                    nombre_completo_analista = analista_encontrado.iloc[0]['Nombre_Analista']

        col_info1, col_info2 = st.columns(2)
        with col_info1:
            st.text_input("Nombre o Razón Social", value=info_caso.get('ADMINISTRADO'), disabled=True)
            st.text_input("Producto", value=info_caso.get('PRODUCTO'), disabled=True)

        with col_info2:
            st.text_input("Analista Económico", value=nombre_completo_analista, disabled=True)
            st.text_input("Sector", value=info_caso.get('SECTOR'), disabled=True)

            df_sector_subdireccion = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Sector_Subdireccion")
            if df_sector_subdireccion is not None and 'ID_Rubro' in df_sector_subdireccion.columns:
                sector_del_caso = info_caso.get('SECTOR')
                if sector_del_caso:
                    rubros_filtrados_df = df_sector_subdireccion[
                        df_sector_subdireccion['Sector_Base'] == sector_del_caso]
                    if not rubros_filtrados_df.empty:
                        lista_rubros = rubros_filtrados_df['Sector_Rubro'].tolist()

                        nombre_rubro_seleccionado = st.selectbox("Elige el rubro", options=lista_rubros, index=None,
                                                                 placeholder="Selecciona una opción...")

                        if nombre_rubro_seleccionado:
                            st.session_state.rubro_seleccionado = nombre_rubro_seleccionado
                            # Buscamos y guardamos el ID del rubro seleccionado
                            fila_rubro = rubros_filtrados_df[
                                rubros_filtrados_df['Sector_Rubro'] == nombre_rubro_seleccionado]
                            if not fila_rubro.empty:
                                id_rubro = fila_rubro.iloc[0]['ID_Rubro']
                                st.session_state.id_rubro_seleccionado = id_rubro
                                st.info(f"ID del Rubro seleccionado: **{id_rubro}**")
                    else:
                        st.warning(f"No hay rubros para el sector '{sector_del_caso}'.")
            else:
                st.error("No se pudo cargar la hoja 'Sector_Subdireccion'.")

            st.subheader("Resolución Subdirectoral (RSD)")
            col_rsd1, col_rsd2 = st.columns([2, 1])
            with col_rsd1:
                st.session_state.numero_rsd = st.text_input("N.º de Resolución Subdirectoral",
                                                            placeholder="Ej: 0242-2024-OEFA/DFAI-SFIS")
            with col_rsd2:
                st.session_state.fecha_rsd = st.date_input("Fecha de notificación de RSD", value=None,
                                                           format="DD/MM/YYYY")

    rubro_ok = st.session_state.get('rubro_seleccionado') is not None
    rsd_ok = st.session_state.get('numero_rsd') and st.session_state.get('numero_rsd').strip() != ""
    fecha_ok = st.session_state.get('fecha_rsd') is not None

    if rubro_ok and rsd_ok and fecha_ok:
        st.session_state.paso2_completo = True
    else:
        st.session_state.paso2_completo = False

else:
    st.warning("Completa todos los campos del Paso 2 para continuar.")

st.write("---")

if (st.session_state.get('paso2_completo') and
        'imputaciones_data' in st.session_state and
        st.session_state.imputaciones_data is not None):

    st.header("Detalles de Hechos Imputados")

    for i in range(len(st.session_state.imputaciones_data)):
        with st.expander(f"Hecho imputado n.° {i + 1}", expanded=(i == 0)):

            # --- PASO 3: DETALLES DEL HECHO IMPUTADO ---
            if st.session_state.rubro_seleccionado:
                st.subheader(f"Paso 3: Detalles del Hecho {i + 1}")
                st.session_state.imputaciones_data[i]['texto_hecho'] = st.text_area(
                    label="Redacta aquí el hecho imputado tal como debe aparecer en el informe:",
                    placeholder="Ejemplo: El administrado incumplió con presentar el Informe de Monitoreo Ambiental...",
                    key=f"texto_hecho_{i}",  # La clave única es fundamental
                    height=150
                )
                st.markdown("---")  # Pequeño separador visual
                df_tipificacion = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Tipificacion_Infracciones")
                if df_tipificacion is not None:
                    try:
                        lista_tipos_infraccion = df_tipificacion['Tipo_Infraccion'].unique().tolist()
                        st.session_state.imputaciones_data[i]['tipo_seleccionado'] = st.radio(
                            "**Selecciona el tipo de infracción:**",
                            options=lista_tipos_infraccion, index=None,
                            horizontal=True, key=f"radio_tipo_infraccion_{i}")

                        if st.session_state.imputaciones_data[i].get('tipo_seleccionado'):
                            subtipos_df = df_tipificacion[
                                df_tipificacion['Tipo_Infraccion'] == st.session_state.imputaciones_data[i]['tipo_seleccionado']]
                            lista_subtipos = subtipos_df['Descripcion_Infraccion'].tolist()
                            st.session_state.imputaciones_data[i]['subtipo_seleccionado'] = st.selectbox(
                                "**Selecciona la descripción de la infracción:**", options=lista_subtipos, index=None,
                                placeholder="Elige una descripción específica...", key=f"subtipo_infraccion_{i}")

                            if st.session_state.imputaciones_data[i].get('subtipo_seleccionado'):
                                descripcion_seleccionada = st.session_state.imputaciones_data[i]['subtipo_seleccionado']
                                id_infraccion = subtipos_df[subtipos_df['Descripcion_Infraccion'] == descripcion_seleccionada].iloc[0]['ID_Infraccion']

                                # Guarda el ID para este hecho específico
                                st.session_state.imputaciones_data[i]['id_infraccion'] = id_infraccion

                                if st.session_state.imputaciones_data[i]['id_infraccion'] == 'INF004':
                                    st.markdown("##### Detalles del Requerimiento")
                                    col_req1, col_req2 = st.columns(2)
                                    with col_req1:
                                        fecha_solicitud = st.date_input("Fecha de solicitud", value=None,
                                                                        format="DD/MM/YYYY",
                                                                        key=f"fecha_sol_{i}")
                                        fecha_entrega = st.date_input("Fecha máxima de entrega", value=None,
                                                                      min_value=fecha_solicitud, format="DD/MM/YYYY",
                                                                      key=f"fecha_ent_{i}")
                                        if fecha_solicitud and fecha_entrega:
                                            holidays_pe = holidays.PE()
                                            dias_habiles = sum(1 for d in
                                                               pd.date_range(fecha_solicitud + timedelta(days=1),
                                                                             fecha_entrega) if
                                                               d.weekday() < 5 and d not in holidays_pe)
                                            st.session_state.dias_habiles_plazo = dias_habiles
                                            st.metric(label="Días Hábiles de Plazo", value=dias_habiles)
                                            fecha_incumplimiento = fecha_entrega + timedelta(days=1)
                                            # Guarda la fecha específica para este hecho 'i'
                                            st.session_state.imputaciones_data[i]['fecha_incumplimiento'] = fecha_incumplimiento
                                            st.info(
                                                f"Fecha de Incumplimiento: **{fecha_incumplimiento.strftime('%d/%m/%Y')}**")
                                    with col_req2:
                                        num_requerimientos = st.number_input("Número de requerimientos de información",
                                                                             min_value=1, step=1,
                                                                             key=f"num_requerimientos_{i}")
                                        estado_remision = st.radio("Estado de la remisión:",
                                                                   options=["No remitió la información",
                                                                            "Remitió completo pero tardío",
                                                                            "Remitió parcial",
                                                                            "Remitió parcial pero tardío"], index=None,
                                                                   key=f"estado_remision_{i}")

                                    if fecha_solicitud and fecha_entrega and num_requerimientos and estado_remision:
                                        st.subheader("Anexar análisis económico del beneficio ilícito")

                                        st.session_state.imputaciones_data[i]['doc_adjunto_hecho'] = st.file_uploader(
                                            "Sube el Word con el análisis económico para este hecho:",
                                            type=["docx"],
                                            key=f"doc_adjunto_{i}"  # Mantenemos la key única
                                        )

                                elif st.session_state.imputaciones_data[i]['id_infraccion'] == 'INF003':
                                    st.markdown("##### Detalles de la Supervisión para Capacitación")

                                    if st.session_state.get('fecha_incumplimiento') and st.session_state.get(
                                            'num_requerimientos') and st.session_state.get('estado_remision'):
                                        st.session_state.paso3_completo = True

                                    col1, col2 = st.columns(2)
                                    with col1:
                                        fecha_supervision = st.date_input(
                                            "Fecha de supervisión",
                                            value=None,
                                            format="DD/MM/YYYY",
                                            key=f"fecha_supervision_{i}"
                                        )
                                        if fecha_supervision:
                                            # Fecha de incumplimiento igual a la de supervisión
                                            st.session_state.imputaciones_data[i]['fecha_incumplimiento'] = fecha_supervision
                                            st.info(f"Fecha de Incumplimiento: **{fecha_supervision.strftime('%d/%m/%Y')}**")
                                    
                                    with col2:
                                        num_personal = st.number_input("Número de personal para capacitación", min_value=1, step=1, key=f"num_personal_{i}")
                                        st.session_state.num_personal_capacitacion = num_personal # Se mantiene para el cálculo
                                        st.session_state.imputaciones_data[i]['num_personal_capacitacion'] = num_personal # CLAVE
                                    
                                    # Subida de documentos
                                    st.subheader("Anexar documentos requeridos")
                                    
                                    doc_sustento = st.file_uploader(
                                        "Sube el documento de sustento de capacitación:",
                                        type=["docx", "pdf"],
                                        key=f"doc_sustento_{i}"
                                    )
                                    
                                    doc_analisis = st.file_uploader(
                                        "Sube el Word con el análisis económico para este hecho:",
                                        type=["docx"],
                                        key=f"doc_analisis_{i}"
                                    )
                                    
                                    # Guardar en session_state
                                    st.session_state.imputaciones_data[i]['doc_sustento'] = doc_sustento
                                    st.session_state.imputaciones_data[i]['doc_adjunto_hecho'] = doc_analisis  # mismo campo que INF004 para compatibilidad
                                    
                                    # Validación de completitud
                                    if fecha_supervision and num_personal and doc_sustento and doc_analisis:
                                        st.session_state.imputaciones_data[i]['paso3_completo'] = True
                                    else:
                                        st.session_state.imputaciones_data[i]['paso3_completo'] = False
                            else:
                                st.session_state.paso3_completo = True
                    except KeyError as e:
                        st.error(f"Error en la hoja 'Tipificacion_Infracciones'. Falta la columna: {e}")
                else:
                    st.error("Error crítico: No se pudo cargar la hoja 'Tipificacion_Infracciones'.")

            if (st.session_state.get(f"subtipo_infraccion_{i}") and
                    st.session_state.get(f"texto_hecho_{i}") and
                    st.session_state.imputaciones_data[i].get('doc_adjunto_hecho')):
                st.session_state.imputaciones_data[i]['paso3_completo'] = True
            else:
                st.session_state.imputaciones_data[i]['paso3_completo'] = False

            st.write("---")  # La línea divisoria

            # app.py

            # --- PASO 4: CÁLCULO DEL COSTO EVITADO (CE) ---
            if st.session_state.imputaciones_data[i].get('paso3_completo'):
                st.subheader(f"Paso 4: Cálculo del CE para el Hecho {i + 1}")

                # 1. Cargar todos los DataFrames necesarios desde Google Sheets
                df_items_infracciones = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Items_Infracciones")
                df_costos_items = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Costos_Items")
                df_coti_general = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Cotizaciones_General")
                df_salarios_general = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Salarios_General")
                df_indices = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Indices_BCRP")

                # Verifica que todas las hojas se hayan cargado antes de continuar
                if all(df is not None for df in [df_items_infracciones, df_costos_items, df_coti_general, df_salarios_general, df_indices]):
                    
                    # 2. Preparar el diccionario de datos de entrada para la función de cálculo
                    datos_para_calculo = {
                        'df_items_infracciones': df_items_infracciones,
                        'df_costos_items': df_costos_items,
                        'df_coti_general': df_coti_general,
                        'df_salarios_general': df_salarios_general,
                        'df_indices': df_indices,
                        'id_infraccion': st.session_state.imputaciones_data[i].get('id_infraccion'),
                        'fecha_incumplimiento': st.session_state.imputaciones_data[i].get('fecha_incumplimiento'),
                        'id_rubro': st.session_state.id_rubro_seleccionado,
                        'dias_habiles': st.session_state.dias_habiles_plazo,
                        'num_personal_capacitacion': st.session_state.get('num_personal_capacitacion', 0)
                    }

                    all_dfs_loaded = all(df is not None for df in [df_items_infracciones, df_costos_items, df_coti_general, df_salarios_general, df_indices])
                    if all_dfs_loaded:
                        resultados_ce = calcular_costo_evitado(datos_para_calculo)
                        if not resultados_ce.get('error'):
                            # --- CORRECCIÓN: Guardar todos los datos necesarios en la sesión ---
                            st.session_state.imputaciones_data[i]['ce_data_raw'] = resultados_ce.get('ce_data_raw', [])
                            st.session_state.imputaciones_data[i]['ce_total_soles'] = resultados_ce.get('total_soles', 0)
                            st.session_state.imputaciones_data[i]['ce_total_dolares'] = resultados_ce.get('total_dolares', 0)
                            st.session_state.imputaciones_data[i]['sustentos'] = resultados_ce.get('sustentos', [])
                            st.session_state.imputaciones_data[i]['ids_anexos'] = resultados_ce.get('ids_anexos', [])
                            st.session_state.imputaciones_data[i]['fuente_coti'] = resultados_ce.get('fuente_coti', '')
                            st.session_state.imputaciones_data[i]['fuente_salario'] = resultados_ce.get('fuente_salario', '')
                            st.session_state.imputaciones_data[i]['pdf_salario'] = resultados_ce.get('pdf_salario', '')
                            st.session_state.imputaciones_data[i]['fi_mes'] = resultados_ce.get('fi_mes', '')
                            st.session_state.imputaciones_data[i]['fi_ipc'] = resultados_ce.get('fi_ipc', '')
                            st.session_state.imputaciones_data[i]['fi_tc'] = resultados_ce.get('fi_tc', '')
                            st.session_state.imputaciones_data[i]['resumen_fuentes_costo'] = resultados_ce.get('resumen_fuentes_costo', '')
                            st.session_state.imputaciones_data[i]['incluye_igv'] = resultados_ce.get('incluye_igv', '')
                            
                            # CLAVE: Guardar datos del primer item para placeholders específicos
                            ce_data = resultados_ce.get('ce_data_raw', [])
                            if ce_data:
                                st.session_state.imputaciones_data[i]['costo_original'] = ce_data[0].get('costo_original', 0)
                                st.session_state.imputaciones_data[i]['moneda_original'] = ce_data[0].get('moneda_original', 'S/')
                        
                        # Reconstruir el DataFrame para mostrarlo en la tabla formateada
                        df_presentacion_ce = pd.DataFrame(resultados_ce['ce_data_raw'])

                        # --- NUEVO: Configuración condicional de la tabla ---
                        id_infraccion_actual = st.session_state.imputaciones_data[i].get('id_infraccion')

                        if id_infraccion_actual == 'INF003':
                            # Configuración específica para INF003 (como en el Word)
                            column_config = {
                                'descripcion': 'Descripción',
                                'precio_dolares': 'Precio asociado (US$)',
                                'precio_soles': 'Precio asociado (S/)',
                                'factor_ajuste': 'Factor de ajuste',
                                'monto_soles': 'Monto (S/)',
                                'monto_dolares': 'Monto (US$)'
                            }
                            formatters_ce = {
                                'Precio asociado (US$)': "US$ {:,.3f}",
                                'Precio asociado (S/)': "{:,.3f}",
                                'Factor de ajuste': "{:,.3f}",
                                'Monto (S/)': "S/ {:,.3f}",
                                'Monto (US$)': "US$ {:,.3f}"
                            }
                        else:
                            # Configuración general para las demás infracciones
                            column_config = {
                                'descripcion': 'Descripción',
                                'cantidad': 'Cantidad',
                                'horas': 'Horas',
                                'precio_soles': 'Precio (S/)',
                                'precio_dolares': 'Precio (US$)',
                                'factor_ajuste': 'Factor de Ajuste',
                                'monto_soles': 'Monto (S/)',
                                'monto_dolares': 'Monto (US$)'
                            }
                            formatters_ce = {
                                "Cantidad": "{:.0f}",
                                "Horas": "{:.0f}",
                                "Precio (S/)": "{:,.3f}",
                                "Precio (US$)": "US$ {:,.3f}",
                                "Factor de Ajuste": "{:,.3f}",
                                "Monto (S/)": "S/ {:,.3f}",
                                "Monto (US$)": "US$ {:,.3f}"
                            }

                        # --- El resto del código se mantiene y usa la configuración elegida ---

                        # Filtrar el DataFrame para quedarnos solo con las columnas que existen y queremos mostrar
                        cols_to_display = [col for col in column_config.keys() if col in df_presentacion_ce.columns]
                        df_display = df_presentacion_ce[cols_to_display].rename(columns=column_config)

                        # Crear la fila de Total usando los nombres de columna correctos
                        total_row = pd.DataFrame([{
                            column_config['descripcion']: "Total", 
                            column_config['monto_soles']: resultados_ce['total_soles'],
                            column_config['monto_dolares']: resultados_ce['total_dolares']
                        }])

                        # Unir los datos con el total y limpiar valores nulos
                        df_final_display = pd.concat([df_display, total_row], ignore_index=True)

                        # Mostrar el DataFrame final y formateado
                        st.dataframe(df_final_display.style.format(formatters_ce, na_rep='').hide(axis="index"),
                                    use_container_width=True)

                        st.session_state.imputaciones_data[i]['paso4_completo'] = True
                else:
                    st.error("Faltan una o más hojas de datos para el cálculo del CE. Revisa la conexión y los nombres de las hojas en Google Sheets.")

            # --- PASO 5: CÁLCULO DEL BENEFICIO ILÍCITO ---
            if st.session_state.imputaciones_data[i].get('paso4_completo'):
                st.subheader(f"Paso 5: Cálculo del BI para el Hecho {i + 1}")

                # 1. Cargar DataFrames necesarios
                df_cos = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "COS")
                df_uit = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "UIT")
                df_indices_bi = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Indices_BCRP")
                
                if all(df is not None for df in [df_cos, df_uit, df_indices_bi]):
                    # 2. Preparar datos de entrada para la función
                    datos_para_bi = {
                        'df_cos': df_cos,
                        'df_uit': df_uit,
                        'df_indices_bi': df_indices_bi,
                        'rubro': st.session_state.rubro_seleccionado,
                        'ce_soles': st.session_state.imputaciones_data[i].get('ce_total_soles', 0),
                        'ce_dolares': st.session_state.imputaciones_data[i].get('ce_total_dolares', 0),
                        'fecha_incumplimiento': st.session_state.imputaciones_data[i].get('fecha_incumplimiento')
                    }

                    # 3. Llamar a la función de cálculo
                    resultados_bi = calcular_beneficio_ilicito(datos_para_bi)

                    # 4. Procesar y mostrar resultados
                    if resultados_bi.get('error'):
                        st.error(resultados_bi['error'])
                    else:
                        # Guardar resultados en session_state
                        st.session_state.imputaciones_data[i]['beneficio_ilicito_uit'] = resultados_bi['beneficio_ilicito_uit']
                        st.session_state.imputaciones_data[i]['bi_data_raw'] = resultados_bi['bi_data_raw']
                        st.session_state.imputaciones_data[i]['fuente_cos'] = resultados_bi['fuente_cos']
                        
                        # Mostrar la tabla de resumen
                        df_resumen = pd.DataFrame(resultados_bi['bi_data_raw'])
                        st.dataframe(df_resumen.style.hide(axis="index"), use_container_width=True)
                        
                        st.session_state.imputaciones_data[i]['paso5_completo'] = True
                else:
                    st.error("Faltan las hojas 'COS' o 'UIT' para el cálculo del BI.")

            # --- PASO 6: CÁLCULO DE LA MULTA ---
            st.write("---")
            if st.session_state.imputaciones_data[i].get('paso5_completo'):
                st.subheader(f"Paso 6: Cálculo de la Multa para el Hecho {i + 1}")

                df_tipificacion_multa = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Tipificacion_Infracciones")

                if df_tipificacion_multa is not None:
                    # 1. Preparar datos de entrada para la función
                    datos_para_multa = {
                        'df_tipificacion': df_tipificacion_multa,
                        'id_infraccion': st.session_state.imputaciones_data[i].get('id_infraccion'),
                        'beneficio_ilicito': st.session_state.imputaciones_data[i].get('beneficio_ilicito_uit', 0)
                    }
                    
                    # 2. Llamar a la función de cálculo
                    resultados_multa = calcular_multa(datos_para_multa)

                    # 3. Procesar y mostrar resultados
                    if resultados_multa.get('error'):
                        st.warning(resultados_multa['error']) # Usamos st.warning para errores no críticos
                    
                    # Guardar y mostrar incluso si hay advertencia
                    multa_calculada = resultados_multa.get('multa_final_uit', 0)
                    st.session_state.imputaciones_data[i]['multa_final_uit'] = multa_calculada
                    st.session_state.imputaciones_data[i]['multa_data_raw'] = resultados_multa.get('multa_data_raw', [])
                    
                    df_multa = pd.DataFrame(resultados_multa.get('multa_data_raw', []))
                    st.dataframe(df_multa.style.format({"Monto": "{} "}).hide(axis="index"), use_container_width=True)

                    st.success(f"**Monto de la Multa Propuesta: {multa_calculada:,.3f} UIT**")
                else:
                    st.error("No se pudo cargar la hoja 'Tipificacion_Infracciones' para el cálculo final.")

    # --------------------------------------------------------------------
    #  GENERAR INFORME FINAL
    # --------------------------------------------------------------------
    # Esta variable comprueba que todos los expanders hayan completado sus cálculos.
    all_steps_complete = False
    if 'imputaciones_data' in st.session_state and st.session_state.imputaciones_data:
        all_steps_complete = all(d.get('paso5_completo', False) for d in st.session_state.imputaciones_data)

    # Si todo está listo, muestra la sección para generar el informe.
    if all_steps_complete:
        st.header("Paso 7: Generar Informe Final")

        # Lógica para encontrar y descargar la plantilla maestra correcta.
        df_plantillas = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Productos")
        id_plantilla = None
        if df_plantillas is not None:
            producto_caso = st.session_state.info_expediente.get('PRODUCTO', '')
            plantilla_row = df_plantillas[df_plantillas['Producto'] == producto_caso]
            if not plantilla_row.empty:
                id_plantilla = plantilla_row.iloc[0]['Producto_Plantilla']
                st.info(f"Usando plantilla específica para el producto '{producto_caso}'.")
            else:
                st.warning(f"No se encontró una plantilla para '{producto_caso}'. Buscando plantilla por defecto...")
                default_row = df_plantillas[df_plantillas['Producto'] == 'DEFAULT']
                if not default_row.empty:
                    id_plantilla = default_row.iloc[0]['Producto_Plantilla']
                    st.info("Usando plantilla por defecto.")
                else:
                    st.error("No se encontró la plantilla específica ni la de 'DEFAULT' en la hoja 'Plantillas'.")
        else:
            st.error("No se pudo cargar la hoja 'Plantillas'.")

        template_file_buffer = None
        if id_plantilla:
            with st.spinner(f"Cargando plantilla ({id_plantilla[:10]}...) desde Google Drive..."):
                template_file_buffer = descargar_archivo_drive(id_plantilla, RUTA_CREDENCIALES_GCP)

        # Muestra el botón final, que activará toda la lógica de creación de documentos.
        if template_file_buffer and st.button("🚀 Generar Informe", type="primary"):
            with st.spinner("Generando informe... Este proceso puede tardar un momento."):
                try:
                    # ==============================================================================
                    # ETAPA 1: PREPARAR DATOS Y PLANTILLA MAESTRA INICIAL
                    # ==============================================================================
                    st.write("🔄 Preparando datos generales del informe...")
                    
                    doc_maestra = DocxTemplate(template_file_buffer)
                    
                    # --- INICIO DEL CÓDIGO QUE PREPARA CONTEXT_DATA ---
                    
                    # 1. Cargar todos los DataFrames una sola vez
                    info_caso = st.session_state.info_expediente
                    df_analistas = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Analistas")
                    df_subdirecciones = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Subdirecciones")
                    df_sector_sub = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Sector_Subdireccion")
                    df_producto_asunto = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Producto_Asunto")
                    df_administrados = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Administrados")

                    # --- Datos de anexo CE --- #
                    df_indices_final = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Indices_BCRP")

                    # 2. Crear el diccionario de datos GENERALES ('context_data')
                    context_data = {'fecha_hoy': date.today().strftime('%d de %B de %Y')}

                    # Obtener nombre completo del administrado
                    nombre_base_administrado = info_caso.get('ADMINISTRADO', '')
                    nombre_final_administrado = nombre_base_administrado
                    if df_administrados is not None:
                        admin_info = df_administrados[df_administrados['Nombre_Administrado_Base'] == nombre_base_administrado]
                        if not admin_info.empty:
                            nombre_final_administrado = admin_info.iloc[0].get('Nombre_Administrado', nombre_base_administrado)
                    context_data['administrado'] = nombre_final_administrado

                    # Añadir detalles del analista y revisor
                    analista_details = get_person_details_by_base_name(info_caso.get('ANALISTA ECONÓMICO'), df_analistas)
                    revisor_details = get_person_details_by_base_name(info_caso.get('REVISOR'), df_analistas)
                    
                    context_data.update({
                        'titulo_analista': analista_details['titulo'],
                        'nombre_analista': analista_details['nombre'],
                        'cargo_analista': analista_details['cargo'],
                        'colegiatura_analista': analista_details['colegiatura'],
                        'titulo_revisor': revisor_details['titulo'],
                        'nombre_revisor': revisor_details['nombre'],
                        'cargo_revisor': revisor_details['cargo'],
                        'colegiatura_revisor': revisor_details['colegiatura']
                    })

                    # Añadir texto sobre el número de imputaciones
                    num_imputaciones = int(info_caso.get('IMPUTACIONES', 1))
                    inf_numero = num_imputaciones
                    inf_texto = num2words(num_imputaciones, lang='es')
                    if inf_numero == 1: inf_texto = "una"
                    inf_oracion = "presunta infracción administrativa" if num_imputaciones == 1 else "presuntas infracciones administrativas"
                    
                    context_data['inf_numero'] = inf_numero
                    context_data['inf_texto'] = inf_texto
                    context_data['inf_oracion'] = inf_oracion

                    context_data['expediente'] = st.session_state.get('num_expediente_formateado', '')
                    context_data['ht'] = info_caso.get('HT', '')
                    context_data['numero_rsd'] = st.session_state.get('numero_rsd', '')
                    context_data['fecha_rsd_texto'] = (st.session_state.get('fecha_rsd') or date.today()).strftime('%d de %B de %Y')

                    # Lógica para obtener datos de la subdirección y SSAG
                    id_sub_row = df_sector_sub[df_sector_sub['Sector_Rubro'] == st.session_state.rubro_seleccionado]
                    if not id_sub_row.empty:
                        id_sub = id_sub_row.iloc[0].get('ID_Subdireccion')
                        sub_row = df_subdirecciones[df_subdirecciones['ID_Subdireccion'] == id_sub]
                        if not sub_row.empty:
                            context_data['nombre_encargado_sub1'] = sub_row.iloc[0].get('Encargado_Sub', '')
                            context_data['cargo_encargado_sub1'] = sub_row.iloc[0].get('Cargo_Encargado_Sub', '')
                            context_data['titulo_encargado_sub1'] = sub_row.iloc[0].get('Titulo_Encargado_Sub', '')
                            context_data['subdireccion'] = sub_row.iloc[0].get('Subdireccion', '')
                            context_data['id_subdireccion'] = sub_row.iloc[0].get('ID_Subdireccion', '')

                    ssag_row = df_subdirecciones[df_subdirecciones['ID_Subdireccion'].astype(str).str.strip().str.upper() == 'SSAG']
                    if not ssag_row.empty:
                        nombre_enc_ssag = ssag_row.iloc[0].get('Encargado_Sub')
                        context_data['nombre_encargado_sub2'] = nombre_enc_ssag
                        context_data['titulo_encargado_sub2'] = ssag_row.iloc[0].get('Titulo_Encargado_Sub', '')
                        context_data['cargo_encargado_sub2'] = ssag_row.iloc[0].get('Cargo_Encargado_Sub', '')
                        if nombre_enc_ssag and df_analistas is not None:
                            enc_ssag_analista_row = df_analistas[df_analistas['Nombre_Analista'] == nombre_enc_ssag]
                            if not enc_ssag_analista_row.empty:
                                context_data['colegiatura_encargado_sub2'] = enc_ssag_analista_row.iloc[0].get('Colegiatura_Analista', '')

                    # Lógica para el "asunto" del producto
                    producto_caso = info_caso.get('PRODUCTO', '')
                    if producto_caso and df_producto_asunto is not None:
                        asunto_row = df_producto_asunto[df_producto_asunto['Producto'] == producto_caso]
                        if not asunto_row.empty:
                            context_data['asunto'] = asunto_row.iloc[0].get('Producto_Asunto', '')
                    
                    # --- Datos de anexo CE --- #
                    # Busca el último mes del índice IPC
                    mes_indice_texto = "No disponible"
                    if df_indices_final is not None and not df_indices_final.empty:
                        try:
                            df_indices_final['Indice_Mes_dt'] = pd.to_datetime(df_indices_final['Indice_Mes'],
                                                                            dayfirst=True, errors='coerce')
                            latest_date = df_indices_final['Indice_Mes_dt'].max()
                            if pd.notna(latest_date):
                                mes_indice_texto = latest_date.strftime('%B %Y').lower()
                        except Exception:
                            pass  # Si hay error, se queda como "No disponible"
                    context_data['mes_indice'] = mes_indice_texto

                    # --- FIN DEL CÓDIGO QUE PREPARA CONTEXT_DATA ---

                    # ==============================================================================
                    # ETAPA 2: PROCESAR CADA HECHO IMPUTADO EN UN BUCLE
                    # ==============================================================================
                    st.write("🔄 **Etapa 2:** Construyendo cada sección de infracción...")
                    lista_secciones_finales = []
                    lista_anexos_ce_finales = []
                    lista_hechos_para_plantilla = []
                    df_tipificacion = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Tipificacion_Infracciones")

                    # --- UN SOLO BUCLE PARA PROCESAR TODO ---
                    for i, datos_hecho in enumerate(st.session_state.imputaciones_data, start=1):
                        st.write(f"   - Procesando Hecho n.° {i}...")

                        # Tarea 1: Preparar la lista para el resumen que va en la plantilla maestra
                        lista_hechos_para_plantilla.append({
                            'numero_imputado': i,
                            'descripcion': datos_hecho.get('texto_hecho', '')
                        })

                        # Tarea 2: Unir y rellenar la sección de la infracción
                        id_infraccion_actual = datos_hecho.get('id_infraccion')
                        fila_infraccion = df_tipificacion[df_tipificacion['ID_Infraccion'] == id_infraccion_actual]
                        
                        id_plantilla_infraccion = fila_infraccion.iloc[0].get('ID_Plantilla_BI')
                        buffer_plantilla_infraccion = descargar_archivo_drive(id_plantilla_infraccion, RUTA_CREDENCIALES_GCP)
                        archivo_analisis_subido = datos_hecho.get('doc_adjunto_hecho')
                        
                        buffer_seccion_unida = io.BytesIO()
                        if archivo_analisis_subido and buffer_plantilla_infraccion:
                            # La función de funciones.py se usa aquí correctamente
                            combinar_con_composer(buffer_plantilla_infraccion, archivo_analisis_subido, buffer_seccion_unida)
                        elif buffer_plantilla_infraccion:
                            buffer_seccion_unida = buffer_plantilla_infraccion
                        else:
                            continue
                        
                        buffer_seccion_unida.seek(0)
                        
                        # Rellenar la sección con su contexto específico
                        doc_seccion_tpl = DocxTemplate(buffer_seccion_unida)
                        contexto_del_hecho = {}
                        try:
                            nombre_modulo = f"infracciones.{id_infraccion_actual}"
                            modulo_especialista = importlib.import_module(nombre_modulo)
                            datos_generales = {
                                'context_data': context_data, 
                                'numero_hecho_actual': i
                            }
                            st.json(datos_hecho) 
                            contexto_del_hecho = modulo_especialista.preparar_contexto_especifico(doc_seccion_tpl, datos_hecho, datos_generales)
                            contexto_del_hecho['mp_uit'] = f"{datos_hecho.get('multa_final_uit', 0):,.3f} UIT"
                            placeholders_a_pasar = [
                                'fuente_salario', 'pdf_salario', 'fuente_coti',
                                'fi_mes', 'fi_ipc', 'fi_tc',
                                'resumen_fuentes_costo',
                                'incluye_igv', 'precio_dol'
                            ]
                            for key in placeholders_a_pasar:
                                contexto_del_hecho[key] = datos_hecho.get(key, '')
                        except ImportError:
                            st.error(f"No se encontró el módulo de lógica para '{id_infraccion_actual}'.")
                            continue
                        
                        doc_seccion_tpl.render(contexto_del_hecho, autoescape=True)
                        buffer_seccion_rellenada = io.BytesIO()
                        doc_seccion_tpl.save(buffer_seccion_rellenada)
                        lista_secciones_finales.append(buffer_seccion_rellenada)

                        # Tarea 3: Preparar el Anexo de Costo Evitado para este hecho
                        id_plantilla_anexo_ce = fila_infraccion.iloc[0].get('ID_Plantilla_CE')
                        if id_plantilla_anexo_ce:
                            buffer_anexo_ce = descargar_archivo_drive(id_plantilla_anexo_ce, RUTA_CREDENCIALES_GCP)
                            if buffer_anexo_ce:
                                anexo_ce_tpl = DocxTemplate(buffer_anexo_ce)
                                anexo_ce_tpl.render(contexto_del_hecho, autoescape=True)
                                buffer_anexo_ce_rellenado = io.BytesIO()
                                anexo_ce_tpl.save(buffer_anexo_ce_rellenado)
                                lista_anexos_ce_finales.append(buffer_anexo_ce_rellenado)

                        st.write(f"   ✓ Hecho n.° {i} listo.")


                    # ==============================================================================
                    # ETAPA 3: ENSAMBLAJE FINAL DEL INFORME MAESTRO
                    # ==============================================================================
                    st.write("\n🧩 **Etapa 3:** Ensamblando el informe maestro...")

                    # --- 3a: Rellenar los datos de la plantilla maestra ---
                    # 1. Crear la tabla de RESUMEN FINAL
                    #    Esta tabla se construye después del bucle para tener los totales de todas las infracciones.
                    summary_rows = []
                    multa_total_uit = 0
                    for i, datos_hecho in enumerate(st.session_state.imputaciones_data, 1):
                        multa_de_este_hecho = datos_hecho.get('multa_final_uit', 0)
                        multa_total_uit += multa_de_este_hecho
                        summary_rows.append({
                            'numeral': f"IV.{i+1}", 
                            'infraccion': f"Hecho imputado n.° {i}",
                            'multa': f"{multa_de_este_hecho:,.3f} UIT"
                        })
                    summary_rows.append({'numeral': 'Total', 'infraccion': '', 'multa': f"{multa_total_uit:,.3f} UIT"})
                    
                    sub_resumen_final = create_summary_table_subdoc(doc_maestra,
                                                                    ["Numeral", "Infracciones", "Multa"],
                                                                    summary_rows,
                                                                    ['numeral', 'infraccion', 'multa'])
                    
                    bi_total = sum(
                    (h.get('beneficio_ilicito_uit') or 0)
                    for h in st.session_state.imputaciones_data
                    )

                    # 2. Construcción final del diccionario 'contexto_maestro'
                    contexto_maestro = {
                        # Se añade la lista de secciones que el bucle {% for %} en tu Word utilizará
                        'lista_hechos_imputados': lista_hechos_para_plantilla, 
                        
                        # Se añade la tabla de resumen final que acabamos de crear
                        'tabla_resumen_final': sub_resumen_final,
                        'multa_final_total': f"{multa_total_uit:,.3f} UIT",
                        'mt_uit': f"{multa_total_uit:,.3f} UIT",
                        
                        # Se añade el resto de los datos generales que ya habíamos preparado en 'context_data'
                        **context_data 
                    }

                    contexto_maestro['bi_uit'] = f"{bi_total:,.3f} UIT"
                    
                    # Si estás usando el método de inserción con docxcompose, "protege" el marcador
                    contexto_maestro['INSERTAR_CONTENIDO_AQUI'] = '{{INSERTAR_CONTENIDO_AQUI}}'

                    doc_maestra.render(contexto_maestro, autoescape=True)
                    buffer_maestra_rellenada = io.BytesIO()
                    doc_maestra.save(buffer_maestra_rellenada)
                    buffer_maestra_rellenada.seek(0)
                    
                    # --- 3b: Usar Composer para insertar las secciones preparadas ---
                    doc_final_base = Document(buffer_maestra_rellenada)
                    compositor_final = Composer(doc_final_base)

                    # Buscamos el índice del párrafo que contiene el marcador
                    marcador_idx = -1
                    for i, p in enumerate(compositor_final.doc.paragraphs):
                        if '{{INSERTAR_CONTENIDO_AQUI}}' in p.text:
                            marcador_idx = i
                            break

                    if marcador_idx != -1:
                        # Iteramos la lista en orden INVERSO para que se apilen correctamente
                        for i, seccion_buffer in enumerate(reversed(lista_secciones_finales)):
                            seccion_buffer.seek(0)
                            
                            # --- INICIO DEL NUEVO CÓDIGO ---
                            # Si NO es la primera plantilla que insertamos (la última de la lista original),
                            # entonces añadimos nuestro separador primero.
                            if i > 0:
                                # Creamos un documento separador con un solo párrafo vacío
                                separador_doc = Document()
                                separador_doc.add_paragraph() # Esto es el "enter"
                                # Insertamos el separador en la posición del marcador
                                compositor_final.insert(marcador_idx, separador_doc)
                            # --- FIN DEL NUEVO CÓDIGO ---
                            
                            # Ahora, insertamos la plantilla de la infracción en el mismo lugar.
                            # Esto la colocará ANTES del separador que acabamos de poner.
                            compositor_final.insert(marcador_idx, Document(seccion_buffer))

                        # Finalmente, eliminamos el párrafo que contenía el marcador original
                        for p in compositor_final.doc.paragraphs:
                            if '{{INSERTAR_CONTENIDO_AQUI}}' in p.text:
                                p._element.getparent().remove(p._element)
                                break

                    # ==============================================================================
                    # ETAPA 4: AÑADIR ANEXOS AL FINAL
                    # ==============================================================================
                    st.write("📑 **Etapa 4:** Añadiendo anexos al final del informe...")

                    # --- Añadir los anexos de Costo Evitado ---
                    if lista_anexos_ce_finales:
                        compositor_final.doc.add_page_break()
                        compositor_final.doc.add_heading("Anexo n.° 1: Detalle del Costo Evitado", level=1)
                        
                        # Bucle para añadir cada anexo de CE
                        for i, anexo_ce_buffer in enumerate(lista_anexos_ce_finales):
                            anexo_ce_buffer.seek(0)
                            compositor_final.append(Document(anexo_ce_buffer))
                            
                            # INICIO DEL CAMBIO: Añadir salto de página si no es el último
                            if i < len(lista_anexos_ce_finales) - 1:
                                compositor_final.doc.add_page_break()
                            # FIN DEL CAMBIO
                                
                        st.write("   ✓ Anexos de Costo Evitado añadidos.")

                    # --- Añadir los anexos de Sustento ---
                    lista_ids_anexos = []
                    for hecho in st.session_state.imputaciones_data:
                        anexos_del_hecho = hecho.get('ids_anexos', [])
                        if anexos_del_hecho:
                            lista_ids_anexos.extend(anexos_del_hecho)

                    if lista_ids_anexos:
                        # Elimina duplicados para no añadir el mismo anexo varias veces
                        lista_ids_anexos = list(dict.fromkeys(lista_ids_anexos))
                        st.write(f"   - Se encontraron {len(lista_ids_anexos)} anexo(s) de sustento para añadir.")
                        
                        # Añadimos un salto de página para separar esta sección de la anterior
                        compositor_final.doc.add_page_break()
                        compositor_final.doc.add_heading("Anexo: Sustento de Costos", level=1)

                        # Bucle para descargar y añadir cada anexo de Drive
                        for i, file_id in enumerate(lista_ids_anexos):
                            try:
                                anexo_drive_buffer = descargar_archivo_drive(file_id, RUTA_CREDENCIALES_GCP)
                                if anexo_drive_buffer:
                                    compositor_final.append(Document(anexo_drive_buffer))
                                    st.write(f"    ✓ Anexo de Drive `{file_id[:10]}...` añadido.")
                                    
                                    # INICIO DEL CAMBIO: Añadir salto de página si no es el último
                                    if i < len(lista_ids_anexos) - 1:
                                        compositor_final.doc.add_page_break()
                                    # FIN DEL CAMBIO
                            except Exception as drive_error:
                                st.error(f"No se pudo añadir el anexo de Drive `{file_id}`. Error: {drive_error}")
                    else:
                        st.write("   - No se encontraron anexos de sustento para añadir.")

                    # ==============================================================================
                    # ETAPA 5: GUARDAR Y DESCARGAR
                    # ==============================================================================
                    final_buffer = io.BytesIO()
                    compositor_final.save(final_buffer)
                    final_buffer.seek(0)

                    # Previsualización opcional con Mammoth
                    with st.expander("✨ Previsualización del Documento Final (versión HTML)"):
                        # Se usa una copia del buffer para la previsualización para no agotar el original
                        preview_buffer = io.BytesIO(final_buffer.getvalue())
                        result = mammoth.convert_to_html(preview_buffer)
                        st.markdown(result.value, unsafe_allow_html=True)

                    # Botón de descarga final
                    st.download_button(
                        label="✅ Descargar Informe Final Compuesto",
                        data=final_buffer.getvalue(),
                        file_name=f"Informe_Multa_{st.session_state.num_expediente_formateado.replace('/', '-')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary"
                    )
                
                    st.success("¡Informe final generado con éxito!")

                except Exception as e:
                    st.error(f"Ocurrió un error al generar el documento: {e}")
                    st.exception(e)

if not cliente_gspread:
    st.error(
        "🔴 No se pudo establecer la conexión con Google Sheets. Revisa el archivo de credenciales y la conexión a internet.")