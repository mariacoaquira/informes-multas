import streamlit as st
import io
import locale
from babel.dates import format_date 
from datetime import datetime, date
import pandas as pd
import importlib
from docx import Document                   
from docxcompose.composer import Composer 
from docxtpl import DocxTemplate, RichText
import mammoth
from num2words import num2words

# Importaciones de nuestros módulos
from sheets import (conectar_gsheet, cargar_hoja_a_df, 
                    get_person_details_by_base_name, descargar_archivo_drive,
                    calcular_beneficio_ilicito, calcular_multa)
from funciones import (combinar_con_composer, create_table_subdoc, 
                     create_main_table_subdoc, create_summary_table_subdoc)

# --- INICIALIZACIÓN DE LA APLICACIÓN ---
st.set_page_config(layout="wide", page_title="Asistente de Multas")
st.title("🤖 Asistente para Generación de Informes de Multa")

if 'app_inicializado' not in st.session_state:
    st.session_state.clear()
    st.session_state.app_inicializado = True

cliente_gspread = conectar_gsheet()
NOMBRE_GSHEET_MAESTRO = "Base de datos"
NOMBRE_GSHEET_ASIGNACIONES = "Base de asignaciones de multas"

# --- CUERPO DE LA APLICACIÓN ---
if cliente_gspread:
    # --- PASO 1: BÚSQUEDA DE EXPEDIENTE ---
    st.header("Paso 1: Búsqueda del Expediente")
    col1, col2 = st.columns([1, 2])
    with col1:
        try:
            locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
        except locale.Error:
            locale.setlocale(locale.LC_TIME, '')
        hojas_disponibles = [format_date(datetime.now() - pd.DateOffset(months=i), "MMMM yyyy", locale='es').capitalize() for i in range(3)]
        mes_seleccionado = st.selectbox("Selecciona el mes de la asignación:", options=hojas_disponibles)
    with col2:
        df_asignaciones = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_ASIGNACIONES, mes_seleccionado)
        if df_asignaciones is not None:
            num_expediente_input = st.text_input("Ingresa el N° de Expediente:", placeholder="Ej: 1234-2023 o 1234-2023-OEFA/DFAI/PAS")
            
            if st.button("Buscar Expediente", type="primary"):
                # --- MEJORA: Limpiamos solo los datos relevantes, no toda la sesión ---
                if 'info_expediente' in st.session_state:
                    del st.session_state['info_expediente']
                if 'imputaciones_data' in st.session_state:
                    del st.session_state['imputaciones_data']
            
            if num_expediente_input:
                num_formateado = ""
                if "OEFA" in num_expediente_input.upper():
                    num_formateado = num_expediente_input
                elif "-" in num_expediente_input:
                    num_formateado = f"{num_expediente_input}-OEFA/DFAI/PAS"

                if num_formateado:
                    resultado = df_asignaciones[df_asignaciones['EXPEDIENTE'] == num_formateado]
                    if not resultado.empty:
                        # Guardamos los datos del expediente si no los teníamos ya
                        if 'info_expediente' not in st.session_state:
                            st.success(f"¡Expediente '{num_formateado}' encontrado!")
                            st.session_state.num_expediente_formateado = num_formateado
                            st.session_state.info_expediente = resultado.iloc[0].to_dict()
                        
                        # --- CORRECCIÓN CLAVE ---
                        # Solo inicializamos la lista de hechos si no existe previamente
                        if 'imputaciones_data' not in st.session_state:
                            num_imputaciones = int(st.session_state.info_expediente.get('IMPUTACIONES', 1))
                            st.session_state.imputaciones_data = [{} for _ in range(num_imputaciones)]
                    else:
                        st.error(f"No se encontró el expediente '{num_expediente_input}'.")
                else:
                    st.warning("Ingresa un número de expediente en un formato válido.")
    st.divider()

    # --- PASO 2: DETALLES DEL EXPEDIENTE ---
    if st.session_state.get('info_expediente'):
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
        
        # Lógica de validación final del Paso 2
        rubro_ok = st.session_state.get('rubro_seleccionado') is not None
        resolucion_ok = False
        info_exp = st.session_state.get('info_expediente')
        producto_caso = info_exp.get('PRODUCTO', '') if info_exp else ''
        if producto_caso == 'RD':
            if st.session_state.get('numero_ifi') and st.session_state.get('fecha_ifi'):
                resolucion_ok = True
        else:
            if st.session_state.get('numero_rsd') and st.session_state.get('fecha_rsd'):
                resolucion_ok = True
        
        if rubro_ok and resolucion_ok:
            st.session_state.paso2_completo = True
        else:
            st.session_state.paso2_completo = False

    # --- PREPARACIÓN DE DATOS GENERALES (SE HACE UNA SOLA VEZ) ---
    if st.session_state.get('paso2_completo') and 'context_data' not in st.session_state:
        with st.spinner("Preparando datos y plantilla general..."):
            # 1. Cargar la Plantilla Maestra
            df_plantillas = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Productos")
            id_plantilla = None
            if df_plantillas is not None:
                producto_caso = st.session_state.info_expediente.get('PRODUCTO', '')
                plantilla_row = df_plantillas[df_plantillas['Producto'] == producto_caso]
                if not plantilla_row.empty:
                    id_plantilla = plantilla_row.iloc[0]['Producto_Plantilla']
                else:
                    default_row = df_plantillas[df_plantillas['Producto'] == 'DEFAULT']
                    if not default_row.empty:
                        id_plantilla = default_row.iloc[0]['Producto_Plantilla']
            
            template_file_buffer = None
            if id_plantilla:
                template_file_buffer = descargar_archivo_drive(id_plantilla)
            
            st.session_state.template_file_buffer = template_file_buffer
            
            # 2. Preparar el diccionario context_data
            if st.session_state.get('template_file_buffer'):
                st.session_state.template_file_buffer.seek(0)
                doc_maestra_obj = DocxTemplate(st.session_state.template_file_buffer)
    
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
                fecha_actual = date.today()
                fecha_rsd_dt = st.session_state.get('fecha_rsd') or fecha_actual

                context_data = {
                    # --- LÍNEAS MODIFICADAS ---
                    'fecha_hoy': format_date(fecha_actual, "d 'de' MMMM 'de' yyyy", locale='es'),
                    'mes_hoy': format_date(fecha_actual, "MMMM 'de' yyyy", locale='es').lower(),
                    'fecha_rsd_texto': format_date(fecha_rsd_dt, "d 'de' MMMM 'de' yyyy", locale='es'),
                }

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
                            mes_indice_texto = format_date(latest_date, 'MMMM yyyy', locale='es').lower()
                    except Exception:
                        pass  # Si hay error, se queda como "No disponible"
                context_data['mes_indice'] = mes_indice_texto

                # --- FIN DEL CÓDIGO QUE PREPARA CONTEXT_DATA ---
                st.session_state.context_data = context_data
                st.success("Datos generales preparados.")

    st.divider()

    # --- BUCLE MODULAR PARA HECHOS IMPUTADOS ---
    if st.session_state.get('context_data'):
        st.header("Paso 3: Detalles de Hechos Imputados")
        df_tipificacion = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Tipificacion_Infracciones")
        
        for i in range(len(st.session_state.get('imputaciones_data', []))):
            with st.expander(f"Hecho imputado n.° {i + 1}", expanded=(i == 0)):
                
                st.subheader(f"Detalles del Hecho {i + 1}")

                # Inputs comunes que siempre se muestran
                st.session_state.imputaciones_data[i]['texto_hecho'] = st.text_area(
                    "Redacta aquí el hecho imputado:",
                    key=f"texto_hecho_{i}",
                    height=150
                )

                st.divider()

                if df_tipificacion is not None:
                    try:
                        lista_tipos_infraccion = df_tipificacion['Tipo_Infraccion'].unique().tolist()
                        tipo_seleccionado = st.radio(
                            "**Selecciona el tipo de infracción:**",
                            options=lista_tipos_infraccion,
                            index=None,
                            horizontal=True,
                            key=f"radio_tipo_infraccion_{i}"
                        )
                        st.session_state.imputaciones_data[i]['tipo_seleccionado'] = tipo_seleccionado

                        if tipo_seleccionado:
                            subtipos_df = df_tipificacion[df_tipificacion['Tipo_Infraccion'] == tipo_seleccionado]
                            lista_subtipos = subtipos_df['Descripcion_Infraccion'].tolist()
                            subtipo_seleccionado = st.selectbox(
                                "**Selecciona la descripción de la infracción:**",
                                options=lista_subtipos,
                                index=None,
                                placeholder="Elige una descripción específica...",
                                key=f"subtipo_infraccion_{i}"
                            )
                            st.session_state.imputaciones_data[i]['subtipo_seleccionado'] = subtipo_seleccionado

                            if subtipo_seleccionado:
                                fila_infraccion = subtipos_df[subtipos_df['Descripcion_Infraccion'] == subtipo_seleccionado].iloc[0]
                                id_infraccion = fila_infraccion['ID_Infraccion']
                                # Guarda el ID para este hecho específico
                                st.session_state.imputaciones_data[i]['id_infraccion'] = id_infraccion
                    
                    except KeyError as e:
                        st.error(f"Error en la hoja 'Tipificacion_Infracciones'. Falta la columna: {e}")
                else:
                    st.error("Error crítico: No se pudo cargar la hoja 'Tipificacion_Infracciones'.")
                    
                id_infraccion = st.session_state.imputaciones_data[i].get('id_infraccion')
                if id_infraccion:
                    try:
                        modulo_especialista = importlib.import_module(f"infracciones.{id_infraccion}")
                        datos_especificos = modulo_especialista.renderizar_inputs_especificos(i)
                        st.session_state.imputaciones_data[i].update(datos_especificos)
                        
                        # --- INICIO DE LA MODIFICACIÓN ---
                        
                        # 1. Validar los datos generales y específicos
                        datos_generales_ok = st.session_state.imputaciones_data[i].get('texto_hecho') and st.session_state.imputaciones_data[i].get('subtipo_seleccionado')
                        datos_especificos_ok = modulo_especialista.validar_inputs(st.session_state.imputaciones_data[i])
                        
                        # 2. El botón solo se habilita si todo está OK
                        boton_habilitado = datos_generales_ok and datos_especificos_ok
                        
                        st.divider()
                        if st.button(f"Calcular Hecho {i+1}", key=f"calc_btn_{i}", disabled=(not boton_habilitado)):
                            
                        # --- FIN DE LA MODIFICACIÓN ---
                            with st.spinner(f"Calculando hecho {i+1}..."):
                                # Cargar DFs y convertirlos a fecha
                                df_coti_general = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Cotizaciones_General")
                                df_indices = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Indices_BCRP")
                                if df_indices is not None:
                                    df_indices['Indice_Mes'] = pd.to_datetime(df_indices['Indice_Mes'], dayfirst=True, errors='coerce')
                                if df_coti_general is not None:
                                    df_coti_general['Fecha_Costeo'] = pd.to_datetime(df_coti_general['Fecha_Costeo'], dayfirst=True, errors='coerce')
                                
                                # Preparar datos comunes, incluyendo el context_data ya creado
                                datos_comunes = {
                                    'df_items_infracciones': cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Items_Infracciones"),
                                    'df_costos_items': cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Costos_Items"),
                                    'df_salarios_general': cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Salarios_General"),
                                    'df_cos': cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "COS"),
                                    'df_uit': cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "UIT"),
                                    'df_coti_general': df_coti_general,
                                    'df_indices': df_indices,
                                    'df_tipificacion': df_tipificacion,
                                    'id_infraccion': id_infraccion,
                                    'rubro': st.session_state.rubro_seleccionado,
                                    'id_rubro_seleccionado': st.session_state.get('id_rubro_seleccionado'),
                                    'numero_hecho_actual': i + 1,
                                    'doc_tpl': DocxTemplate(st.session_state.template_file_buffer),
                                    'context_data': st.session_state.get('context_data', {})
                                }
                                
                                resultados_completos = modulo_especialista.procesar_infraccion(
                                    datos_comunes, 
                                    st.session_state.imputaciones_data[i]
                                )
                                # --- INICIO DE LA CORRECCIÓN ---
                                # Revisa si el especialista reportó un error
                                if resultados_completos.get('error'):
                                    st.error(f"Error en el cálculo del Hecho {i+1}: {resultados_completos['error']}")
                                else:
                                    # Si no hay error, guarda los resultados y muestra el mensaje de éxito
                                    st.session_state.imputaciones_data[i]['resultados'] = resultados_completos
                                    st.success(f"Hecho {i+1} calculado.")

                                    # --- AÑADE ESTA LÍNEA PARA GUARDAR LOS ANEXOS DEL HECHO ---
                                    st.session_state.imputaciones_data[i]['anexos_ce'] = resultados_completos.get('anexos_ce_generados', [])
                                # --- FIN DE LA CORRECCIÓN ---
                    except ImportError:
                        st.error(f"El módulo para '{id_infraccion}' no está implementado.")

                # Sección para mostrar resultados ya calculados
                if 'resultados' in st.session_state.imputaciones_data[i]:
                    # Extraemos los resultados que preparamos para la app
                    resultados_app = st.session_state.imputaciones_data[i]['resultados']['resultados_para_app']
                    st.subheader(f"Resultados del Cálculo para el Hecho {i + 1}")
                    
                    # --- Mostrar Tabla de Costo Evitado ---
                    st.markdown("###### Costo Evitado (CE)")
                    df_presentacion_ce = pd.DataFrame(resultados_app.get('ce_data_raw', []))
                    
                    # Lógica de tabla condicional (la que ya tenías)
                    id_infraccion_actual = st.session_state.imputaciones_data[i].get('id_infraccion')
                    if id_infraccion_actual == 'INF003':
                        column_config = {
                            'descripcion': 'Descripción', 'precio_dolares': 'Precio asociado (US$)',
                            'precio_soles': 'Precio asociado (S/)', 'factor_ajuste': 'Factor de ajuste',
                            'monto_soles': 'Monto (S/)', 'monto_dolares': 'Monto (US$)'
                        }
                        formatters_ce = {
                            "Cantidad": "{:.0f}",
                            # Esta función muestra el entero si no hay decimales, o redondea a 3 si los hay
                            "Horas": lambda x: f"{int(x)}" if pd.notna(x) and x == int(x) else f"{x:,.3f}",
                            "Precio (S/)": "S/ {:,.3f}",
                            "Precio (US$)": "US$ {:,.3f}",
                            "Factor de Ajuste": "{:,.3f}",
                            "Monto (S/)": "S/ {:,.3f}",
                            "Monto (US$)": "US$ {:,.3f}"
                        }
                    else:
                        # Configuración general para las demás
                        column_config = {
                            'descripcion': 'Descripción', 'cantidad': 'Cantidad', 'horas': 'Horas',
                            'precio_soles': 'Precio (S/)', 'precio_dolares': 'Precio (US$)',
                            'factor_ajuste': 'Factor de Ajuste', 'monto_soles': 'Monto (S/)', 'monto_dolares': 'Monto (US$)'
                        }
                        formatters_ce = {
                            "Cantidad": "{:.0f}", "Horas": lambda x: f"{int(x)}" if pd.notna(x) and x == int(x) else f"{x:,.3f}", "Precio (S/)": "{:,.3f}",
                            "Precio (US$)": "US$ {:,.3f}", "Factor de Ajuste": "{:,.3f}",
                            "Monto (S/)": "S/ {:,.3f}", "Monto (US$)": "US$ {:,.3f}"
                        }

                    cols_to_display = [col for col in column_config.keys() if col in df_presentacion_ce.columns]
                    df_display = df_presentacion_ce[cols_to_display].rename(columns=column_config)
                    total_row = pd.DataFrame([{'Descripción': "Total", 'Monto (S/)': resultados_app['ce_total_soles'], 'Monto (US$)': resultados_app['ce_total_dolares']}])
                    df_final_display = pd.concat([df_display, total_row], ignore_index=True)
                    st.dataframe(df_final_display.style.format(formatters_ce, na_rep='').hide(axis="index"), use_container_width=True)
                    
                    # --- Mostrar Tabla de Beneficio Ilícito ---
                    st.markdown("###### Beneficio Ilícito (BI)")
                    df_bi_crudo = pd.DataFrame(resultados_app.get('bi_data_raw', []))

                    # --- INICIO DE LA MODIFICACIÓN ---
                    if not df_bi_crudo.empty:
                        # 1. Seleccionamos solo las columnas que queremos mostrar
                        columnas_a_mostrar = ['descripcion', 'monto']
                        
                        # 2. Creamos un nuevo DataFrame solo con esas columnas
                        df_bi_display = df_bi_crudo[columnas_a_mostrar]
                        
                        # 3. Renombramos las columnas para una mejor presentación
                        df_bi_display = df_bi_display.rename(columns={
                            'descripcion': 'Descripción',
                            'monto': 'Monto'
                        })
                        
                        # 4. Mostramos el DataFrame limpio
                        st.dataframe(df_bi_display.style.hide(axis="index"), use_container_width=True)
                    else:
                        # Si no hay datos, muestra una tabla vacía como antes
                        st.dataframe(df_bi_crudo)
                    # --- FIN DE LA MODIFICACIÓN ---

                    # --- Mostrar Tabla de Multa ---
                    st.markdown("###### Multa Propuesta")
                    df_multa = pd.DataFrame(resultados_app.get('multa_data_raw', []))
                    st.dataframe(df_multa.style.format({"Monto": "{} "}).hide(axis="index"), use_container_width=True)

                    pass

# --------------------------------------------------------------------
#  PASO 7: GENERAR INFORME FINAL
# --------------------------------------------------------------------
# Variable que comprueba que todos los hechos han sido calculados en la sesión
all_steps_complete = False
if 'imputaciones_data' in st.session_state and st.session_state.imputaciones_data:
    all_steps_complete = all('resultados' in d for d in st.session_state.imputaciones_data)

# Si todo está listo, muestra la sección para generar el informe
if all_steps_complete:
    st.divider()
    st.header("Paso 4: Generar Informe Final")

    # El botón de generar informe solo se muestra si la plantilla maestra se cargó correctamente
    if st.session_state.get('template_file_buffer'):
        if st.button("🚀 Generar Informe", type="primary"):
            with st.spinner("Generando informe... Este proceso puede tardar un momento."):
                try:
                    # --- ETAPA 1: PREPARAR DATOS Y PLANTILLA MAESTRA ---
                    st.write("🔄 **Etapa 1:** Preparando datos maestros...")
                    st.session_state.template_file_buffer.seek(0)
                    doc_maestra = DocxTemplate(st.session_state.template_file_buffer)
                    context_data = st.session_state.get('context_data', {})

                    # Crear la tabla de RESUMEN FINAL de multas
                    summary_rows = []
                    multa_total_uit = 0
                    lista_hechos_para_plantilla = []

                    for i, datos_hecho in enumerate(st.session_state.imputaciones_data, 1):
                        lista_hechos_para_plantilla.append({
                            'numero_imputado': i,
                            'descripcion': datos_hecho.get('texto_hecho', '')
                        })

                        resultados_hecho = datos_hecho.get('resultados', {}).get('resultados_para_app', {})
                        
                        # 1. Obtenemos el valor de la multa con todos sus decimales
                        multa_de_este_hecho_float = resultados_hecho.get('multa_final_uit', 0)
                        
                        # 2. Redondeamos el valor individual a 3 decimales
                        multa_de_este_hecho_redondeada = round(multa_de_este_hecho_float, 3)
                        
                        # 3. Sumamos el valor YA redondeado al total
                        multa_total_uit += multa_de_este_hecho_redondeada
                        
                        summary_rows.append({
                            'numeral': f"IV.{i + 1}", 
                            'infraccion': f"Hecho imputado n.° {i}",
                            'multa': f"{multa_de_este_hecho_redondeada:,.3f} UIT" # Usamos el valor redondeado para mostrar
                        })

                    # Aseguramos que el total final también se formatee correctamente
                    summary_rows.append({'numeral': 'Total', 'infraccion': '', 'multa': f"{multa_total_uit:,.3f} UIT"})
                    # Define el texto de elaboración
                    texto_elaboracion = "Elaboración: Subdirección de Sanción y Gestión Incentivos (SSAG) - DFAI."
                    anchos_para_resumen = (1, 4, 1.5) # 1" para Numeral, 4" para Infracciones, 1.5" para Multa
                    # Pasa el texto al crear la tabla de resumen
                    sub_resumen_final = create_summary_table_subdoc(
                        doc_maestra, 
                        ["Numeral", "Infracciones", "Multa"], 
                        summary_rows, 
                        ['numeral', 'infraccion', 'multa'],
                        texto_posterior=texto_elaboracion,
                        column_widths=anchos_para_resumen # <-- Aquí pasas los anchos
                    )
                    
                    # Construcción final del diccionario 'contexto_maestro'
                    contexto_maestro = {
                        **context_data,
                        'lista_hechos_imputados': lista_hechos_para_plantilla, 
                        'tabla_resumen_final': sub_resumen_final,
                        'multa_final_total': f"{multa_total_uit:,.3f} UIT",
                        'mt_uit': f"{multa_total_uit:,.3f} UIT",
                        'INSERTAR_CONTENIDO_AQUI': '{{INSERTAR_CONTENIDO_AQUI}}'
                    }
                    
                    doc_maestra.render(contexto_maestro, autoescape=True)
                    buffer_maestra_rellenada = io.BytesIO()
                    doc_maestra.save(buffer_maestra_rellenada)
                    buffer_maestra_rellenada.seek(0)

                    # --- ETAPA 2: RECOGER SECCIONES DE CADA HECHO ---
                    st.write("🔄 **Etapa 2:** Construyendo cada sección de infracción...")
                    lista_secciones_finales = []
                    lista_anexos_ce_finales = []
                    lista_hechos_para_plantilla = [] 

                    df_tipificacion = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Tipificacion_Infracciones")
                    
                    for i, datos_hecho in enumerate(st.session_state.imputaciones_data, start=1):

                        # --- AÑADE ESTE BLOQUE PARA RECOLECTAR LOS ANEXOS ---
                        if 'anexos_ce' in datos_hecho and datos_hecho['anexos_ce']:
                            lista_anexos_ce_finales.extend(datos_hecho['anexos_ce'])
                    # ----------------------------------------------------
                        resultados = datos_hecho.get('resultados', {})
                        contexto_a_renderizar = resultados.get('contexto_final_word', {})
                        
                        id_infraccion_actual = datos_hecho.get('id_infraccion')
                        fila_infraccion = df_tipificacion[df_tipificacion['ID_Infraccion'] == id_infraccion_actual]
                        
                        id_plantilla_infraccion = fila_infraccion.iloc[0].get('ID_Plantilla_BI')
                        buffer_plantilla_infraccion = descargar_archivo_drive(id_plantilla_infraccion)
                        archivo_analisis_subido = datos_hecho.get('doc_adjunto_hecho')
                        
                        buffer_seccion_unida = io.BytesIO()
                        if archivo_analisis_subido and buffer_plantilla_infraccion:
                            combinar_con_composer(buffer_plantilla_infraccion, archivo_analisis_subido, buffer_seccion_unida)
                        elif buffer_plantilla_infraccion:
                            buffer_seccion_unida = buffer_plantilla_infraccion
                        else:
                            continue
                        
                        buffer_seccion_unida.seek(0)
                        doc_seccion_tpl = DocxTemplate(buffer_seccion_unida)
                        doc_seccion_tpl.render(contexto_a_renderizar, autoescape=True)
                        # --- INICIO DEL CÓDIGO A AÑADIR ---
                        # Elimina el último párrafo de la sección si está vacío
                        doc = doc_seccion_tpl.docx
                        if doc.paragraphs and not doc.paragraphs[-1].text.strip():
                            p_element = doc.paragraphs[-1]._element
                            p_element.getparent().remove(p_element)
                        # --- FIN DEL CÓDIGO A AÑADIR ---
                        buffer_seccion_rellenada = io.BytesIO()
                        doc_seccion_tpl.save(buffer_seccion_rellenada)
                        lista_secciones_finales.append(buffer_seccion_rellenada)

                    # --- ETAPA 3: ENSAMBLAJE FINAL DEL INFORME MAESTRO ---
                    st.write("\n🧩 **Etapa 3:** Ensamblando el informe maestro...")
                    doc_final_base = Document(buffer_maestra_rellenada)
                    compositor_final = Composer(doc_final_base)
                    
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
                        # --- LÍNEA CORREGIDA ---
                        anexos_del_hecho = hecho.get('resultados', {}).get('ids_anexos', [])
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
                                anexo_drive_buffer = descargar_archivo_drive(file_id)
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

                    # Primero obtenemos el nombre del expediente de forma segura
                    nombre_exp = st.session_state.get('num_expediente_formateado', 'EXPEDIENTE_SIN_NUMERO')

                    # Botón de descarga final
                    st.download_button(
                        label="✅ Descargar Informe Final Compuesto",
                        data=final_buffer.getvalue(),
                        file_name=f"Informe_Multa_{nombre_exp.replace('/', '-')}.docx",
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
