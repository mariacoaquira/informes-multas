import streamlit as st
import pandas as pd
import io
from datetime import timedelta, date
from docxtpl import DocxTemplate, RichText
from babel.dates import format_date

# --- Imports from your other modules ---
# Hacia la línea 13 de INF002.py
from funciones import (create_main_table_subdoc, create_table_subdoc,
                     create_consolidated_bi_table_subdoc, texto_con_numero,
                     create_footnotes_subdoc, create_detailed_ce_table_subdoc,
                     formatear_periodo_monitoreo, create_considerations_table_subdoc,
                     redondeo_excel, create_ce2_envio_table_subdoc, create_ce2_lab_table_subdoc,
                     create_graduation_table_subdoc) # <--- AÑADIR ESTA
from sheets import calcular_beneficio_ilicito, calcular_multa, descargar_archivo_drive
from textos_manager import obtener_fuente_formateada


# --- CÁLCULO DEL COSTO EVITADO ---
def _calcular_costo_evitado_monitoreo(datos_comunes, datos_extremo):
    """
    Calcula el CE, agrupa anexos y recopila sustentos de TODOS los grupos.
    """
    try:
        # 1. Desempaquetado de datos
        df_items_infracciones = datos_comunes['df_items_infracciones']
        df_costos_items = datos_comunes['df_costos_items']
        df_coti_general = datos_comunes['df_coti_general']
        df_salarios_general = datos_comunes['df_salarios_general']
        df_indices = datos_comunes['df_indices']
        id_rubro = datos_comunes.get('id_rubro_seleccionado')
        
        fecha_incumplimiento = datos_extremo.get('fecha_incumplimiento')
        tipo_monitoreo_sel = datos_extremo.get('tipo_monitoreo_sel')
        tipo_servicio = datos_extremo.get('tipo_servicio')
        parametros_seleccionados = datos_extremo.get('parametros_seleccionados', [])
        multiplicador_puntos = datos_extremo.get('cantidad', 1)

        # 2. Lógica de filtrado
        receta_df = df_items_infracciones[(df_items_infracciones['ID_Infraccion'] == 'INF002') & (df_items_infracciones['Descripcion_Item'].str.contains(tipo_monitoreo_sel, na=False))].copy()
        if receta_df.empty: return {'error': f"No se encontró receta para '{tipo_monitoreo_sel}'."}
        
        receta_requiere_parametros = any(rec['Nombre_Item'] == 'Parámetro' for rec in receta_df.to_dict('records'))
        if receta_requiere_parametros and not parametros_seleccionados:
             return {'error': 'Debe seleccionar al menos un parámetro.'}

        if tipo_servicio == "Solo análisis de parámetros":
            receta_df_filtrada_servicio = receta_df[receta_df['Tipo_Costo'].str.startswith('Análisis', na=False)]
        else:
            receta_df_filtrada_servicio = receta_df
        
        fecha_inc_dt = pd.to_datetime(fecha_incumplimiento)
        ipc_row_inc = df_indices[df_indices['Indice_Mes'].dt.to_period('M') == fecha_inc_dt.to_period('M')]
        if ipc_row_inc.empty: return {'error': f"No se encontró IPC para '{fecha_inc_dt.strftime('%B %Y')}'."}

        ipc_incumplimiento = ipc_row_inc.iloc[0]['IPC_Mensual']
        tc_incumplimiento = ipc_row_inc.iloc[0]['TC_Mensual']

        fecha_inc_texto = format_date(fecha_inc_dt, "MMMM 'de' yyyy", locale='es')
        texto_ipc_incumplimiento = f"{fecha_inc_texto.capitalize()}, IPC = {ipc_incumplimiento:,.6f}"
        texto_tc_incumplimiento = f"{fecha_inc_texto.capitalize()}, TC = {tc_incumplimiento:,.4f}"
        
        fuente_salario = ""
        pdf_salario = ""
        texto_ipc_costeo_salario = ""
        anio_salario = ""
        
        # --- INICIO MODIFICACIÓN: Conjuntos para sustentos por grupo ---
        sustentos_personal_set = set()
        sustentos_seguros_set = set()
        sustentos_epp_set = set()
        sustentos_movilidad_set = set()
        # --- FIN MODIFICACIÓN ---

        ce1_items, ce2_envio_items, ce2_lab_items = [], [], []
        ce1_soles, ce1_dolares, ce2_envio_soles, ce2_envio_dolares, ce2_lab_soles, ce2_lab_dolares = 0, 0, 0, 0, 0, 0
        
        anexos_agrupados = {}
        orden_fuentes = []

        # 3. Bucle principal
        for _, item_receta in receta_df_filtrada_servicio.iterrows():
            id_item_a_buscar = item_receta['ID_Item_Infraccion']
            nombre_item_receta = item_receta.get('Nombre_Item')
            tipo_item_receta = item_receta.get('Tipo_Item')

            posibles_costos = df_costos_items[df_costos_items['ID_Item_Infraccion'] == id_item_a_buscar].copy()
            if posibles_costos.empty: continue
            
            costos_a_procesar = pd.DataFrame()
            if nombre_item_receta == 'Parámetro':
                costos_a_procesar = posibles_costos[posibles_costos['Descripcion_Item'].isin(parametros_seleccionados)]
            else:
                candidatos_df = pd.DataFrame()
                if tipo_item_receta == 'Variable':
                    posibles_costos['ID_Rubro'] = posibles_costos['ID_Rubro'].astype(str)
                    candidatos_con_rubro = posibles_costos[posibles_costos['ID_Rubro'].str.contains(id_rubro, na=False)]
                    candidatos_df = candidatos_con_rubro if not candidatos_con_rubro.empty else posibles_costos[posibles_costos['ID_Rubro'].isnull() | (posibles_costos['ID_Rubro'] == 'nan') | (posibles_costos['ID_Rubro'] == '')]
                else: 
                    candidatos_df = posibles_costos
                
                if not candidatos_df.empty:
                    fechas_fuente = []
                    for _, candidato in candidatos_df.iterrows():
                        id_general_temp = candidato['ID_General']
                        fecha_fuente = pd.NaT
                        if pd.notna(id_general_temp):
                            if 'SAL' in id_general_temp:
                                fuente = df_salarios_general[df_salarios_general['ID_Salario'] == id_general_temp]
                                if not fuente.empty: fecha_fuente = pd.to_datetime(f"{int(fuente.iloc[0]['Costeo_Salario'])}-12-31")
                            elif 'COT' in id_general_temp:
                                fuente = df_coti_general[df_coti_general['ID_Cotizacion'] == id_general_temp]
                                if not fuente.empty: fecha_fuente = fuente.iloc[0]['Fecha_Costeo']
                        fechas_fuente.append(fecha_fuente)
                    candidatos_df['Fecha_Fuente'] = fechas_fuente
                    candidatos_df.dropna(subset=['Fecha_Fuente'], inplace=True)
                    if not candidatos_df.empty:
                        candidatos_df['Diferencia_Dias'] = (candidatos_df['Fecha_Fuente'] - fecha_inc_dt).dt.days.abs()
                        costos_a_procesar = candidatos_df.loc[[candidatos_df['Diferencia_Dias'].idxmin()]]
            
            if costos_a_procesar.empty: continue

            for _, costo_calculado in costos_a_procesar.iterrows():
                id_general = costo_calculado.get('ID_General')
                ipc_costeo, tc_costeo = 0, 0
                
                if pd.notna(id_general):
                    if id_general not in anexos_agrupados:
                        anexos_agrupados[id_general] = {'items': [], 'full': None}
                        orden_fuentes.append(id_general)

                    if 'SAL' in id_general:
                        fuente = df_salarios_general[df_salarios_general['ID_Salario'] == id_general]
                        if not fuente.empty:
                            anio = int(fuente.iloc[0]['Costeo_Salario'])
                            indices_anio = df_indices[df_indices['Indice_Mes'].dt.year == anio]
                            if not indices_anio.empty: 
                                ipc_costeo, tc_costeo = indices_anio['IPC_Mensual'].mean(), indices_anio['TC_Mensual'].mean()
                                if not fuente_salario:
                                    anio_salario = str(anio)
                                    fuente_salario = fuente.iloc[0].get('Fuente_Salario', '')
                                    pdf_salario = fuente.iloc[0].get('PDF_Salario', '')
                                    texto_ipc_costeo_salario = f"Promedio {anio}, IPC = {ipc_costeo:,.6f}"

                    elif 'COT' in id_general:
                        fuente = df_coti_general[df_coti_general['ID_Cotizacion'] == id_general]
                        if not fuente.empty:
                            fecha_fuente_dt = fuente.iloc[0]['Fecha_Costeo']
                            ipc_row = df_indices[df_indices['Indice_Mes'].dt.to_period('M') == fecha_fuente_dt.to_period('M')]
                            if not ipc_row.empty: ipc_costeo, tc_costeo = ipc_row.iloc[0]['IPC_Mensual'], ipc_row.iloc[0]['TC_Mensual']
                
                            id_anexo_completo = fuente.iloc[0].get('ID_Anexo_Drive')
                            if pd.notna(id_anexo_completo):
                                anexos_agrupados[id_general]['full'] = id_anexo_completo
                
                if pd.isna(ipc_costeo) or ipc_costeo == 0: continue
                
                if pd.notna(costo_calculado.get('ID_Anexo_Drive')):
                    if pd.notna(id_general):
                        anexos_agrupados[id_general]['items'].append(costo_calculado.get('ID_Anexo_Drive'))

                cantidad_recursos_base = pd.to_numeric(item_receta.get('Cantidad_Recursos'), errors='coerce'); cantidad_recursos_base = 1 if pd.isna(cantidad_recursos_base) or cantidad_recursos_base == 0 else cantidad_recursos_base
                cantidad_horas_base = pd.to_numeric(item_receta.get('Cantidad_Horas'), errors='coerce'); cantidad_horas_base = 1 if pd.isna(cantidad_horas_base) or cantidad_horas_base == 0 else cantidad_horas_base
                
                costo_unitario_orig = float(costo_calculado['Costo_Unitario_Item'])
                
                incluye_igv = costo_calculado.get('Incluye_IGV', 'SI')
                igv_multiplier = 1.00
                if incluye_igv.upper() == 'NO': igv_multiplier = 1.18
                precio_ajustado_igv = costo_unitario_orig * igv_multiplier
                precio_unitario_final = redondeo_excel(precio_ajustado_igv, 3)
                
                precio_base_soles = precio_unitario_final if costo_calculado['Moneda_Item'] == 'S/' else precio_unitario_final * tc_costeo
                factor_ajuste = redondeo_excel(ipc_incumplimiento / ipc_costeo, 3)

                tipo_costo = item_receta.get('Tipo_Costo', '')
                grupo_principal = str(tipo_costo).split('-')[0]

                if tipo_costo == 'Análisis-Laboratorio':
                    multiplicador_aplicable = multiplicador_puntos
                    display_cantidad = multiplicador_aplicable 
                    unidad_texto = costo_calculado.get('Unidad', 'Und')
                else:
                    multiplicador_aplicable = 1
                    display_cantidad = cantidad_recursos_base 
                    if grupo_principal in ['Personal', 'Movilidad']: 
                        unidad_texto = f"{int(cantidad_horas_base)} horas"
                    elif 'horas' in str(costo_calculado.get('Unidad', '')).lower():
                         unidad_texto = f"{int(cantidad_horas_base)} horas"
                    else:
                        unidad_texto = costo_calculado.get('Unidad', 'Und')

                monto_soles_raw = multiplicador_aplicable * cantidad_recursos_base * cantidad_horas_base * precio_base_soles * factor_ajuste
                monto_soles = redondeo_excel(monto_soles_raw, 3)
                monto_dolares_raw = monto_soles / tc_incumplimiento if tc_incumplimiento > 0 else 0
                monto_dolares = redondeo_excel(monto_dolares_raw, 3)

                grupo, subgrupo = (tipo_costo.split('-', 1) + [None])[:2] if '-' in tipo_costo else (tipo_costo, None)
                
                item_calculado = { "descripcion": costo_calculado.get('Descripcion_Item', 'N/A'), "unidad": unidad_texto, "cantidad": display_cantidad, "precio_unitario": precio_unitario_final, "factor_ajuste": factor_ajuste, "monto_soles": monto_soles, "monto_dolares": monto_dolares, "grupo": grupo, "subgrupo": subgrupo }
                
                # --- INICIO MODIFICACIÓN: Recolectar Sustentos por Grupo ---
                sustento_txt = costo_calculado.get('Sustento_Item')
                if pd.notna(sustento_txt) and str(sustento_txt).strip():
                    clean_txt = str(sustento_txt).strip()
                    if grupo == 'Personal':
                        sustentos_personal_set.add(clean_txt)
                    elif grupo == 'Seguro y certificaciones':
                        sustentos_seguros_set.add(clean_txt)
                    elif grupo == 'Equipo de protección personal':
                        sustentos_epp_set.add(clean_txt)
                    elif grupo == 'Movilidad':
                        sustentos_movilidad_set.add(clean_txt)
                # --- FIN MODIFICACIÓN ---

                if grupo in ['Personal', 'Seguro y certificaciones', 'Equipo de protección personal', 'Movilidad']:
                    ce1_items.append(item_calculado); ce1_soles += monto_soles; ce1_dolares += monto_dolares
                elif tipo_costo == 'Análisis-Envío':
                    ce2_envio_items.append(item_calculado); ce2_envio_soles += monto_soles; ce2_envio_dolares += monto_dolares
                elif tipo_costo == 'Análisis-Laboratorio':
                    ce2_lab_items.append(item_calculado); ce2_lab_soles += monto_soles; ce2_lab_dolares += monto_dolares
        
        todos_los_anexos_final = []
        for id_gen in orden_fuentes:
            grupo = anexos_agrupados[id_gen]
            todos_los_anexos_final.extend(grupo['items'])
            if grupo['full']:
                todos_los_anexos_final.append(grupo['full'])
        
        # --- INICIO: Helper Original (String simple) ---
        # Línea 228 aprox. en INF002.py
        def formatear_sustentos(conjunto_sustentos):
            if not conjunto_sustentos:
                return ""
            
            lista = sorted(list(conjunto_sustentos))
            # Usamos RichText para que Word reconozca el \n como un salto de párrafo real
            return RichText("\n".join([f"- {s}" for s in lista]))

        sustento_personal_texto = formatear_sustentos(sustentos_personal_set)
        sustento_seguros_texto = formatear_sustentos(sustentos_seguros_set)
        sustento_epp_texto = formatear_sustentos(sustentos_epp_set)
        sustento_movilidad_texto = formatear_sustentos(sustentos_movilidad_set)
        # --- FIN MODIFICACIÓN ---

        return {
            "ce1_items": ce1_items, "ce1_soles": ce1_soles, "ce1_dolares": ce1_dolares, 
            "ce2_envio_items": ce2_envio_items, "ce2_envio_soles": ce2_envio_soles, "ce2_envio_dolares": ce2_envio_dolares, 
            "ce2_lab_items": ce2_lab_items, "ce2_lab_soles": ce2_lab_soles, "ce2_lab_dolares": ce2_lab_dolares, 
            "total_general_soles": ce1_soles + ce2_envio_soles + ce2_lab_soles, 
            "total_general_dolares": ce1_dolares + ce2_envio_dolares + ce2_lab_dolares, 
            "ids_anexos": todos_los_anexos_final, 
            "error": None,
            "texto_ipc_incumplimiento": texto_ipc_incumplimiento,
            "texto_tc_incumplimiento": texto_tc_incumplimiento,
            "fuente_salario": fuente_salario,
            "pdf_salario": pdf_salario,
            "texto_ipc_costeo_salario": texto_ipc_costeo_salario,
            "anio_salario": anio_salario,
            
            # Nuevos campos en el retorno
            "sustento_personal": sustento_personal_texto,
            "sustento_seguros": sustento_seguros_texto,
            "sustento_epp": sustento_epp_texto,
            "sustento_movilidad": sustento_movilidad_texto,
        }

    except Exception as e:
        import traceback; traceback.print_exc()
        return {'error': f"Error crítico en cálculo de CE: {e}"}
    

# --- INTERFAZ DE USUARIO ---
def renderizar_inputs_especificos(i, df_dias_no_laborables=None):
    """
    Renderiza la interfaz de usuario en Streamlit para la infracción INF002,
    siguiendo un flujo lógico y con campos dinámicos.
    """
    st.markdown("##### Detalles del Monitoreo Ambiental Omitido")
    datos_hecho = st.session_state.imputaciones_data[i]
    if 'extremos' not in datos_hecho:
        datos_hecho['extremos'] = [{}]

    # Cargar los datos necesarios desde el estado de la sesión
    df_items_infracciones = st.session_state.datos_calculo['df_items_infracciones']
    df_costos_items = st.session_state.datos_calculo['df_costos_items']
    df_recetas_inf002 = df_items_infracciones[df_items_infracciones['ID_Infraccion'] == 'INF002']
    
    if df_recetas_inf002.empty:
        st.error("No se encontraron recetas para INF002 en la hoja 'Items_Infracciones'.")
        return datos_hecho
    
    if st.button("+ Añadir Extremo", key=f"add_extremo_{i}"):
        datos_hecho['extremos'].append({})
        st.rerun()

    # Iterar sobre cada extremo que el usuario haya añadido
    for j, extremo in enumerate(datos_hecho['extremos']):
        with st.container(border=True):
            st.markdown(f"**Extremo de Incumplimiento n.° {j + 1}**")
            
            # 1. TIPO DE MONITOREO
            # --- MODIFICACIÓN 1: Restringir tipos de monitoreo ---
            lista_tipos_monitoreo = ["Monitoreo de Aire", "Monitoreo de Agua", "Monitoreo de Ruido"]
            
            tipo_monitoreo_sel = st.selectbox(
                "1. Tipo de Monitoreo",
                options=lista_tipos_monitoreo,
                key=f"tipo_monitoreo_{i}_{j}",
                index=lista_tipos_monitoreo.index(extremo.get('tipo_monitoreo_sel')) if extremo.get('tipo_monitoreo_sel') in lista_tipos_monitoreo else None,
                placeholder="Seleccione..."
            )

            extremo['tipo_monitoreo_sel'] = tipo_monitoreo_sel
            
            # 2. PUNTOS DE MONITOREO
            extremo['cantidad'] = st.number_input(
                "2. Puntos de Monitoreo",
                min_value=1,
                step=1,
                key=f"cantidad_{i}_{j}",
                value=extremo.get('cantidad', 1)
            )

            # 2.1 CÓDIGOS DE LOS PUNTOS
            extremo['nombres_puntos'] = st.text_input( 
                "2.1. Nombres de los Puntos de Monitoreo", 
                key=f"nombres_puntos_{i}_{j}", 
                value=extremo.get('nombres_puntos', ''), 
                placeholder="Ej: MR-01, EFLU-02" 
            ) 

            # 2.2 FRECUENCIA DEL MONITOREO
            opciones_frecuencia = ["Trimestral", "Semestral", "Anual"]
            frecuencia_guardada = extremo.get('frecuencia_monitoreo')
            
            extremo['frecuencia_monitoreo'] = st.selectbox(
                "2.2. Frecuencia del Monitoreo",
                options=opciones_frecuencia,
                key=f"frecuencia_{i}_{j}",
                index=opciones_frecuencia.index(frecuencia_guardada) if frecuencia_guardada in opciones_frecuencia else None,
                placeholder="Seleccione la frecuencia..."
            )

            # 3. PARÁMETROS OMITIDOS (se muestra solo si se eligió un tipo de monitoreo)
            # --- MODIFICACIÓN 2: Nombres de parámetros desde Costos_Items ---
            if tipo_monitoreo_sel:
                receta_parametros = df_recetas_inf002[
                    (df_recetas_inf002['Descripcion_Item'].str.contains(tipo_monitoreo_sel, na=False)) &
                    (df_recetas_inf002['Nombre_Item'] == 'Parámetro')
                ]
                ids_items_parametros = receta_parametros['ID_Item_Infraccion'].tolist()
                
                # Filtramos la tabla de costos usando los IDs encontrados en la receta, 
                # pero extraemos los nombres (labels) de 'Descripcion_Item' de la tabla de costos.
                opciones_parametros = df_costos_items[
                    df_costos_items['ID_Item_Infraccion'].isin(ids_items_parametros)
                ]['Descripcion_Item'].unique().tolist()
                
                parametros_sel = st.multiselect(
                    "3. Parámetros Omitidos",
                    options=opciones_parametros,
                    key=f"parametros_{i}_{j}",
                    default=extremo.get('parametros_seleccionados', [])
                )
                extremo['parametros_seleccionados'] = parametros_sel
            
            # 4. TIPO DE SERVICIO
            tipo_servicio = st.radio(
                "4. Servicio Omitido",
                ["Monitoreo completo (Muestreo y Análisis)", "Solo análisis de parámetros"],
                key=f"tipo_servicio_{i}_{j}",
                index=["Monitoreo completo (Muestreo y Análisis)", "Solo análisis de parámetros"].index(extremo.get('tipo_servicio')) if extremo.get('tipo_servicio') else None
            )
            extremo['tipo_servicio'] = tipo_servicio

            # 5. FECHAS 
            if tipo_servicio:
                col_fecha1, col_fecha2 = st.columns(2)
                with col_fecha1:
                    today = date.today()
                    if tipo_servicio == "Monitoreo completo (Muestreo y Análisis)":
                        fecha_maxima = st.date_input(
                            "5. Fecha máxima de ejecución",
                            key=f"fecha_max_{i}_{j}",
                            value=extremo.get('fecha_base'),
                            format="DD/MM/YYYY",
                            max_value=today
                        )
                        if fecha_maxima:
                            extremo['fecha_base'] = fecha_maxima
                            extremo['fecha_incumplimiento'] = fecha_maxima + timedelta(days=1)
                    else: # Solo análisis
                        fecha_ejecucion = st.date_input(
                            "5. Fecha de ejecución",
                            key=f"fecha_ejec_{i}_{j}",
                            value=extremo.get('fecha_base'),
                            format="DD/MM/YYYY",
                            max_value=today
                        )
                        if fecha_ejecucion:
                            extremo['fecha_base'] = fecha_ejecucion
                            extremo['fecha_incumplimiento'] = fecha_ejecucion
                with col_fecha2:
                    if extremo.get('fecha_incumplimiento'):
                        st.metric("Fecha de Incumplimiento", extremo['fecha_incumplimiento'].strftime('%d/%m/%Y'))

            # Botón para eliminar el extremo
            if st.button(f"🗑️ Eliminar Extremo {j + 1}", key=f"del_extremo_{i}_{j}"):
                datos_hecho['extremos'].pop(j)
                st.rerun()
                
    return datos_hecho

def validar_inputs(datos_hecho):
    """
    Valida que todos los campos necesarios para cada extremo estén completos.
    """
    # 1. Verifica que la lista de extremos exista y no esté vacía.
    if not datos_hecho.get('extremos'):
        return False

    # 2. Itera sobre cada extremo para validar sus datos individualmente.
    for extremo in datos_hecho.get('extremos'):
        # 3. Comprueba que los campos básicos estén presentes y tengan valor.
        if not all([
            extremo.get('fecha_incumplimiento'),
            extremo.get('tipo_servicio'),
            extremo.get('tipo_monitoreo_sel'),
            extremo.get('cantidad', 0) > 0
        ]):
            return False

        # 4. Validación condicional para los parámetros.
        if extremo.get('tipo_monitoreo_sel'):
            # Busca en la "receta" si este tipo de monitoreo tiene items marcados como 'Parámetro'.
            receta_parametros_check = st.session_state.datos_calculo['df_items_infracciones'][
                (st.session_state.datos_calculo['df_items_infracciones']['Descripcion_Item'].str.contains(extremo['tipo_monitoreo_sel'], na=False)) &
                (st.session_state.datos_calculo['df_items_infracciones']['Nombre_Item'] == 'Parámetro')
            ]
            
            # Si la receta indica que hay parámetros para elegir, verifica que el usuario haya seleccionado al menos uno.
            if not receta_parametros_check.empty and not extremo.get('parametros_seleccionados'):
                st.warning(f"Para el monitoreo '{extremo['tipo_monitoreo_sel']}', debe seleccionar al menos un parámetro.")
                return False
                
    # 5. Si todos los extremos pasan todas las validaciones, retorna True.
    return True

def _procesar_hecho_simple(datos_comunes, datos_hecho):
    """
    Procesa un hecho simple:
    1. Genera un anexo de CE completo con todos los detalles.
    2. Genera un cuerpo de informe que contiene BI, Multa y la tabla de Consideraciones.
    """
    # 1. Cargar plantillas
    df_tipificacion, id_infraccion = datos_comunes['df_tipificacion'], datos_comunes['id_infraccion']
    fila_infraccion = df_tipificacion[df_tipificacion['ID_Infraccion'] == id_infraccion].iloc[0]
    id_tpl_cuerpo = fila_infraccion.get('ID_Plantilla_BI')
    id_tpl_anexo_ce = fila_infraccion.get('ID_Plantilla_CE')
    if not id_tpl_cuerpo: return {'error': 'Falta ID de plantilla BI en Tipificación para INF002.'}
    buffer_cuerpo = descargar_archivo_drive(id_tpl_cuerpo)
    if not buffer_cuerpo: return {'error': 'Fallo en descarga de la plantilla.'}

    doc_tpl_para_tablas = DocxTemplate(io.BytesIO(buffer_cuerpo.getvalue()))
    extremo = datos_hecho['extremos'][0]

    res_ce = _calcular_costo_evitado_monitoreo(datos_comunes, extremo)
    if res_ce.get('error'): return {'error': f"Error en el extremo 1: {res_ce['error']}"}

    # --- Lógica de Ordenamiento (Req 1: Planificación antes de Muestreo) ---
    orden_presentacion = ['Planificación', 'Muestreo', 'Seguro y certificaciones', 'Equipo de protección personal', 'Movilidad']
    
    def obtener_orden(item):
        subgrupo = item.get('subgrupo')
        try:
            return orden_presentacion.index(subgrupo)
        except ValueError:
            return len(orden_presentacion) # Enviar al final si no se encuentra
            
    # La lista de items a pasar a la función que genera la tabla debe ser la ordenada
    ce1_items_ordenados = sorted(res_ce['ce1_items'], key=obtener_orden)

    total_ce_soles = res_ce['total_general_soles']
    total_ce_dolares = res_ce['total_general_dolares']
    todos_los_anexos = res_ce['ids_anexos']

    # 2. Cálculos de BI y Multa
    texto_hecho_bi = f"{datos_hecho.get('texto_hecho', 'Incumplimiento de monitoreos')}"
    datos_bi = {**datos_comunes, 'ce_soles': total_ce_soles, 'ce_dolares': total_ce_dolares, 'fecha_incumplimiento': extremo['fecha_incumplimiento'], 'texto_del_hecho': texto_hecho_bi}
    res_bi = calcular_beneficio_ilicito(datos_bi)
    if res_bi.get('error'): return res_bi
    beneficio_ilicito_uit = res_bi.get('beneficio_ilicito_uit', 0)
    # --- Recuperar Factor F de la interfaz ---
    factor_f = datos_hecho.get('factor_f_calculado', 1.0)

    res_multa = calcular_multa({
        **datos_comunes, 
        'beneficio_ilicito': beneficio_ilicito_uit,
        'factor_f': factor_f # <--- PASAR EL FACTOR
    })
    multa_uit = res_multa.get('multa_final_uit', 0) # <-- OBTENER VALOR

    # --- INICIO: LÓGICA DE REDUCCIÓN Y TOPE (Simple) ---
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
    # --- FIN: LÓGICA DE REDUCCIÓN Y TOPE ---

    # Crear tablas de BI y Multa para el cuerpo del informe
    filas_bi_crudas = res_bi.get('table_rows', [])
    filas_bi_para_tabla = []
    for fila in filas_bi_crudas:
        nueva_fila = fila.copy()
        ref_letra = nueva_fila.get('ref')
        
        # Esta es la lógica clave que faltaba:
        texto_base = str(nueva_fila.get('descripcion_texto', ''))
        # sheets.py pone la 'T' del exponente aquí:
        super_existente = str(nueva_fila.get('descripcion_superindice', '')) 
        
        # Añadir la letra de la nota al pie (ej. (a))
        if ref_letra: super_existente += f"({ref_letra})" 
        
        nueva_fila['descripcion_texto'] = texto_base
        nueva_fila['descripcion_superindice'] = super_existente
        filas_bi_para_tabla.append(nueva_fila)
    footnotes_list_bi = [f"({letra}) {obtener_fuente_formateada(ref_key, res_bi.get('footnote_data', {}), id_infraccion=id_infraccion)}" for letra, ref_key in res_bi.get('footnote_mapping', {}).items()]
    footnotes_bi = {'list': footnotes_list_bi, 'style': 'FuenteTabla'}
    tabla_bi_subdoc = create_main_table_subdoc(doc_tpl_para_tablas, ["Descripción", "Monto"], filas_bi_para_tabla, ['descripcion_texto', 'monto'], footnotes_data=footnotes_bi, column_widths=(5, 1.5))
    tabla_multa_subdoc = create_main_table_subdoc(doc_tpl_para_tablas, ["Componentes", "Monto"], res_multa.get('multa_data_raw', []), ['Componentes', 'Monto'], column_widths=(5, 1.5), texto_posterior="Elaboración: Subdirección de Sanción y Gestión de Incentivos (SSAG) - DFAI.", estilo_texto_posterior='FuenteTabla')

# 1. Extraer datos del extremo único
    frecuencia_seleccionada = extremo.get('frecuencia_monitoreo', 'Trimestral')
    periodo = formatear_periodo_monitoreo(extremo.get('fecha_base'), frecuencia_seleccionada)
    matriz = extremo.get('tipo_monitoreo_sel', '').replace('Monitoreo de ', '').replace(' Ambiental', '')
    # --- Preparación de datos para la tabla de consideraciones ---
    cantidad_puntos = extremo.get('cantidad', 1)
    nombres_puntos = extremo.get('nombres_puntos', '')
    # Se agrega .capitalize() para que inicie con "Un (1)..."
    puntos_descripcion = f"{texto_con_numero(cantidad_puntos).capitalize()} punto{'s' if cantidad_puntos > 1 else ''} de monitoreo: {nombres_puntos}"
    
    # Nueva lógica para Parámetros con conteo y mayúscula inicial
    lista_params = extremo.get('parametros_seleccionados', [])
    cant_params = len(lista_params)
    parametros = f"{texto_con_numero(cant_params).capitalize()} parámetro{'s' if cant_params > 1 else ''}: {', '.join(lista_params)}"

    # 2. Construir la lista de datos para la tabla (Esto es lo que faltaba)
    datos_tabla_consideraciones = [{
        'Periodo del monitoreo': periodo,
        'Matriz': matriz,
        'Puntos de monitoreo': puntos_descripcion,
        'Parámetros': parametros
    }]

    # 3. Generar el subdocumento de la tabla
    # Nota: Usamos doc_tpl_para_tablas (el objeto creado desde el buffer en la línea 415)
    tabla_consideraciones_consolidada_subdoc = ""
    try:
        tabla_consideraciones_consolidada_subdoc = create_considerations_table_subdoc(
            doc_tpl_para_tablas,
            ["Periodo del monitoreo", "Matriz", "Puntos de monitoreo", "Parámetros"],
            datos_tabla_consideraciones,
            ['Periodo del monitoreo', 'Matriz', 'Puntos de monitoreo', 'Parámetros'], 
            texto_posterior="Elaboración: Subdirección de Sanción y Gestión de Incentivos (SSAG) - DFAI.", 
            estilo_texto_posterior='FuenteTabla'
        )
    except Exception as e_tabla:
        st.error(f"Error generando tabla de consideraciones: {e_tabla}")
        tabla_consideraciones_consolidada_subdoc = "Error en la generación de la tabla."

    # --- INICIO: ADICIÓN DE LABEL PARA PÁRRAFO BI ---
    tipo_monitoreo_sel = extremo.get('tipo_monitoreo_sel', '')
    tipo_servicio = extremo.get('tipo_servicio', '')
    
    es_monitoreo_completo = (tipo_servicio == "Monitoreo completo (Muestreo y Análisis)")
    
    # Este es el label que usarás en el párrafo de análisis
    label_analisis_bi = "CE2" if es_monitoreo_completo else "CE"

    # --- INICIO: Nueva Condición (Petición Adicional) ---
    # Esta bandera será 'False' solo si AMBAS condiciones se cumplen
    es_ruido_y_solo_analisis = ("Monitoreo de Ruido" in tipo_monitoreo_sel and 
                                 tipo_servicio == "Solo análisis de parámetros")
    # La bandera para la plantilla debe ser 'True' si el párrafo debe MOSTRARSE
    mostrar_parrafo_especial = not es_ruido_y_solo_analisis

# Línea 570 aprox. en INF002.py
    # --- CORRECCIÓN: Bandera para párrafo de análisis ---
    es_monitoreo_ruido = ("Monitoreo de Ruido" in tipo_monitoreo_sel)
    # El párrafo solo se muestra si NO es ruido
    mostrar_parrafo_especial = not es_monitoreo_ruido
    # --- FIN: Nueva Condición ---
    # --- FIN: ADICIÓN DE LABEL ---

# --- 2.1. Lógica de Moneda (REQ 3) ---
    moneda_calculo = res_bi.get('moneda_cos', 'USD')
    es_dolares = (moneda_calculo == 'USD')
    texto_moneda_bi = "moneda extranjera (Dólares)" if es_dolares else "moneda nacional (Soles)"
    ph_bi_abreviatura_moneda = "US$" if es_dolares else "S/"

    # --- 2.2. Superíndices BI Reordenados (REQ 4) ---
    filas_bi_crudas = res_bi.get('table_rows', [])
    fn_map_orig = res_bi.get('footnote_mapping', {})
    fn_data = res_bi.get('footnote_data', {})
    letras_usadas = sorted(list({r for f in filas_bi_crudas if f.get('ref') for r in f.get('ref').replace(" ", "").split(",") if r}))
    letras_base = "abcdefghijklmnopqrstuvwxyz"
    map_traduccion = {v: letras_base[i] for i, v in enumerate(letras_usadas)}
    nuevo_fn_map = {map_traduccion[v]: fn_map_orig[v] for v in letras_usadas if v in fn_map_orig}

    filas_bi_para_tabla = []
    for fila in filas_bi_crudas:
        ref_orig = fila.get('ref', '')
        super_final = str(fila.get('descripcion_superindice', ''))
        if ref_orig:
            nuevas = [map_traduccion[r] for r in ref_orig.replace(" ", "").split(",") if r in map_traduccion]
            if nuevas: super_final += f"({', '.join(nuevas)})"
        filas_bi_para_tabla.append({'descripcion_texto': fila.get('descripcion_texto', ''), 'descripcion_superindice': super_final, 'monto': fila.get('monto', '')})

    fn_list = [f"({l}) {obtener_fuente_formateada(k, fn_data, id_infraccion=id_infraccion)}" for l, k in sorted(nuevo_fn_map.items())]
    footnotes_bi = {'list': fn_list, 'elaboration': 'Elaboración: Subdirección de Sanción y Gestión de Incentivos (SSAG) - DFAI.', 'style': 'FuenteTabla'}
    tabla_bi_subdoc = create_main_table_subdoc(doc_tpl_para_tablas, ["Descripción", "Monto"], filas_bi_para_tabla, ['descripcion_texto', 'monto'], footnotes_data=footnotes_bi, column_widths=(5, 1.5))

    # --- 2.3. Lógica de Factores de Graduación e IPC (REQ 2, 5, 7) ---
    aplica_grad = datos_hecho.get('aplica_graduacion') == 'Sí'
    # Inicialización de variables de graduación
    tabla_grad_subdoc, ph_factor_f_completo, ph_factores_inactivos = "", "1.00 (100%)", ""
    ph_cantidad_f, ph_lista_f, detalle_grad_rt = "cero (0)", "", ""
    placeholders_anexo_grad = {}

    # --- LÓGICA DE GRADUACIÓN (COPIA EXACTA DE INF004) ---
    aplica_grad = datos_hecho.get('aplica_graduacion') == 'Sí'
    
    # Inicialización de variables para evitar errores
    tabla_grad_subdoc = ""
    ph_factor_f_completo = "1.00 (100%)"
    ph_factores_inactivos = ""
    ph_cantidad_f = "cero (0)"
    ph_lista_f = ""
    detalle_grad_rt = ""
    suma_f_acumulado = 0.0
    placeholders_anexo_grad = {} 
    
    # Extraemos datos y número de hecho
    grad_data = datos_hecho.get('graduacion', {})
    numero_hecho_int = datos_comunes.get('numero_hecho_actual', 1)
    idx_hecho_actual = numero_hecho_int - 1
    
    factores_activos_lista = []
    factores_inactivos_labels = [] 
    detalle_grad_rt = RichText() 
    rows_cuadro = []
    suma_f_acumulado = 0.0 
    letras = "abcdefghijklmnopqrstuvwxyz"
    count_f = 0
    
    # Títulos técnicos para la tabla
    titulos_f = {
        'f1': 'Gravedad del daño al interés público y/o bien jurídico protegido',
        'f2': 'El perjuicio económico causado',
        'f3': 'Aspectos ambientales o fuentes de contaminación',
        'f4': 'Reincidencia en la comisión de la infracción',
        'f5': 'Corrección de la conducta infractora',
        'f6': 'Adopción de las medidas necesarias para revertir las consecuencias de la conducta infractora',
        'f7': 'Intencionalidad en la conducta del infractor'
    }

    # Títulos para el resumen dinámico
    titulos_resumen_map = {
        'f1': 'gravedad del daño al ambiente', 'f2': 'perjuicio económico causado',
        'f3': 'aspectos ambientales o fuentes de contaminación', 'f4': 'reincidencia',
        'f5': 'corrección de la conducta infractora', 
        'f6': 'adopción de las medidas necesarias para revertir las consecuencias de la conducta infractora',
        'f7': 'intencionalidad'
    }

    if aplica_grad:
        for cod_f in ['f1', 'f2', 'f3', 'f4', 'f5', 'f6', 'f7']:
            valor_f = grad_data.get(f"subtotal_{cod_f}", 0.0)
            suma_f_acumulado += valor_f
            
            # A. Construir datos para la tabla
            rows_cuadro.append({
                'factor': f"{cod_f}. {titulos_f[cod_f]}",
                'calificacion': f"{valor_f:.0%}"
            })

            if valor_f != 0:
                letra = letras[count_f]
                factores_activos_lista.append(f"({letra}) {cod_f}: {titulos_f[cod_f].lower()}")
                count_f += 1
                
                if detalle_grad_rt.xml: detalle_grad_rt.add("\n\n")
                detalle_grad_rt.add(f"Factor {cod_f.upper()}: {titulos_f[cod_f].upper()}", bold=True, underline=True)
                
                prefix_key = f"grad_{idx_hecho_actual}_{cod_f}_"
                for key, valor_seleccionado in grad_data.items():
                    if key.startswith(prefix_key) and not key.endswith("_valor"):
                        subtitulo = key.replace(prefix_key, "")
                        detalle_grad_rt.add(f"\n{subtitulo}: ", bold=True)
                        detalle_grad_rt.add(f"{valor_seleccionado}")
            else:
                factores_inactivos_labels.append(f"{cod_f} ({titulos_resumen_map[cod_f]})")

        # Filas de Totales
        rows_cuadro.append({'factor': '(f1+f2+f3+f4+f5+f6+f7)', 'calificacion': f"{suma_f_acumulado:.0%}"})
        factor_f_final_val = 1.0 + suma_f_acumulado
        rows_cuadro.append({'factor': 'Factores: F = (1+f1+f2+f3+f4+f5+f6+f7)', 'calificacion': f"{factor_f_final_val:.0%}"})

        # Resumen de Inactivos
        if len(factores_inactivos_labels) == 1:
            ph_factores_inactivos = f"el factor {factores_inactivos_labels[0]} tiene"
        elif len(factores_inactivos_labels) > 1:
            lista_str = ", ".join(factores_inactivos_labels[:-1]) + " y " + factores_inactivos_labels[-1]
            ph_factores_inactivos = f"los factores {lista_str} tienen"

        # Crear tabla subdoc
        tabla_grad_subdoc = create_graduation_table_subdoc(
            doc_tpl_para_tablas, headers=["Factores", "Calificación"], data=rows_cuadro, keys=['factor', 'calificacion'],
            texto_posterior="Elaboración: Subdirección de Sanción y Gestión de Incentivos (SSAG) – DFAI.",
            column_widths=(5.7, 0.5)
        )

        # Generar placeholders individuales (ph_f1_valor, etc.)
        for f_key in ['f1', 'f2', 'f3', 'f4', 'f5', 'f6', 'f7']:
            subtotal_f = grad_data.get(f"subtotal_{f_key}", 0.0)
            placeholders_anexo_grad[f"ph_{f_key}_valor"] = f"{subtotal_f:.0%}"
            prefix_f = f"grad_{idx_hecho_actual}_{f_key}_"
            criterios_claves = sorted([k for k in grad_data.keys() if k.startswith(prefix_f) and k.endswith("_valor")])
            for i, key_crit in enumerate(criterios_claves, 1):
                placeholders_anexo_grad[f"ph_{f_key}_{i}_valor"] = f"{grad_data.get(key_crit, 0.0):.0%}"

        placeholders_anexo_grad["ph_suma_f_total"] = f"{suma_f_acumulado:.0%}"
        placeholders_anexo_grad["ph_hecho_numero"] = str(numero_hecho_int)
        ph_factor_f_completo = f"{factor_f_final_val:,.2f} ({factor_f_final_val:.0%})"
        ph_cantidad_f = texto_con_numero(count_f, genero='m')
        ph_lista_f = ", ".join(factores_activos_lista[:-1]) + " y " + factores_activos_lista[-1] if len(factores_activos_lista) > 1 else (factores_activos_lista[0] if factores_activos_lista else "")

    ph_num_anexo_grad = "3" if aplica_grad else "2"
    # --- FIN LÓGICA DE GRADUACIÓN ---

    # 3. Contexto y renderizado para el CUERPO DEL INFORME
    contexto_final_word = {
        **datos_comunes['context_data'],
        'hecho': {
            'numero_imputado': datos_comunes['numero_hecho_actual'],
            'descripcion': RichText(datos_hecho.get('texto_hecho', '')),
            'tabla_consideraciones_consolidada': tabla_consideraciones_consolidada_subdoc, # <-- AÑADIDO
            'tabla_bi': tabla_bi_subdoc,
            'tabla_multa': tabla_multa_subdoc,
            # --- INICIO: PLACEHOLDERS MOVIDOS AQUÍ ---
            'es_monitoreo_completo': es_monitoreo_completo,
            'label_analisis_bi': label_analisis_bi,
            'mostrar_parrafo_no_ruido_analisis': mostrar_parrafo_especial, # <-- AÑADIR ESTA LÍNEA
            # --- FIN: PLACEHOLDERS MOVIDOS AQUÍ ---
        },
        'numeral_hecho': f"IV.{datos_comunes['numero_hecho_actual']}",
        'multa_original_uit': f"{multa_uit:,.3f} UIT",
        'mh_uit': f"{multa_final_del_hecho_uit:,.3f} UIT",
        'bi_uit': f"{beneficio_ilicito_uit:,.3f} UIT",
        'fuente_cos': res_bi.get('fuente_cos', ''),

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
        # REQ 3 & 5
        'bi_moneda_es_dolares': es_dolares,
        'ph_anio_salario': res_ce.get('anio_salario', ''),
        'ph_bi_moneda_texto': texto_moneda_bi,
        'ph_bi_moneda_simbolo': ph_bi_abreviatura_moneda,
        'ph_ipc_promedio_salario': res_ce.get('texto_ipc_costeo_salario', ''),
        
        # REQ 7 (Numeración de Anexo)
        'ph_anexo_num_grad': ph_num_anexo_grad,
        
        'aplica_graduacion': aplica_grad,
        'tabla_graduacion_sancion': tabla_grad_subdoc,
        'ph_factor_f_final_completo': ph_factor_f_completo,
        'ph_factores_inactivos_resumen': ph_factores_inactivos,
        'ph_detalle_graduacion_extenso': detalle_grad_rt,
        'ph_cantidad_graduacion': ph_cantidad_f,      # Antes: ph_cantidad_f_grad
        'ph_lista_graduacion_inline': ph_lista_f,
        'ph_anexo_num_grad': ph_num_anexo_grad,
        **placeholders_anexo_grad,  # IMPORTANTE: Esto habilita ph_f1_valor, ph_f1_1_valor, etc.
        
        'sustento_seguros': res_ce.get('sustento_seguros', ''),
    }
    
    doc_tpl_cuerpo = DocxTemplate(io.BytesIO(buffer_cuerpo.getvalue()))
    doc_tpl_cuerpo.render(contexto_final_word)
    buffer_final_hecho = io.BytesIO()
    doc_tpl_cuerpo.save(buffer_final_hecho)

    # 4. Generar el ANEXO DE CE
    anexos_ce_generados = []
    if id_tpl_anexo_ce:
        buffer_anexo = descargar_archivo_drive(id_tpl_anexo_ce)
        if buffer_anexo:
            anexo_tpl = DocxTemplate(buffer_anexo)

            # Crear todas las tablas de CE específicamente para el anexo
            tabla_consideraciones_anexo_subdoc = create_considerations_table_subdoc(anexo_tpl, ["Periodo del monitoreo", "Matriz", "Puntos de monitoreo", "Parámetros"], datos_tabla_consideraciones, ['Periodo del monitoreo', 'Matriz', 'Puntos de monitoreo', 'Parámetros'])
            subdoc_anexo_ce1 = create_detailed_ce_table_subdoc(anexo_tpl, ce1_items_ordenados, res_ce['ce1_soles'], res_ce['ce1_dolares']) if res_ce['ce1_items'] else None

            subdoc_anexo_ce2_envio, subdoc_anexo_ce2_lab, subdoc_anexo_resumen_ce2, subdoc_anexo_resumen_total = None, None, None, None

            # --- CE2 ENVÍO (Diseño Personalizado) ---
            if res_ce['ce2_envio_items']:
                # Usamos la nueva función
                subdoc_anexo_ce2_envio = create_ce2_envio_table_subdoc(
                    anexo_tpl, 
                    res_ce['ce2_envio_items'], 
                    res_ce['ce2_envio_soles'], 
                    res_ce['ce2_envio_dolares']
                )
            
            # --- CE2 LABORATORIO (Diseño Personalizado) ---
            if res_ce['ce2_lab_items']:
                # 1. Preparar los datos extra (reportes, precio total)
                datos_lab_formateados = []
                for item in res_ce['ce2_lab_items']:
                    nuevo_item = item.copy()
                    
                    # Lógica: Cantidad = Puntos (ya viene así)
                    # Reportes = 1 por defecto (o puedes personalizarlo si tienes ese dato)
                    nuevo_item['reportes'] = 1 
                    
                    # Precio Total S/ = Precio Unitario * Cantidad * Reportes
                    # Nota: 'precio_unitario' en item ya tiene IGV y redondeo
                    precio_total = redondeo_excel(item['precio_unitario'] * item['cantidad'] * nuevo_item['reportes'], 3)
                    nuevo_item['precio_total'] = precio_total
                    
                    datos_lab_formateados.append(nuevo_item)

                # 2. Obtener nombre de la matriz para el encabezado
                nombre_matriz = extremo.get('tipo_monitoreo_sel', '').replace('Monitoreo de ', '').replace(' Ambiental', '')

                # 3. Crear tabla
                subdoc_anexo_ce2_lab = create_ce2_lab_table_subdoc(
                    anexo_tpl,
                    datos_lab_formateados,
                    res_ce['ce2_lab_soles'],
                    res_ce['ce2_lab_dolares'],
                    nombre_matriz
                )

            # ... (código siguiente: total_ce2_soles = ...)
            total_ce2_soles = res_ce['ce2_envio_soles'] + res_ce['ce2_lab_soles']
            total_ce2_dolares = res_ce['ce2_envio_dolares'] + res_ce['ce2_lab_dolares']
            if total_ce2_soles > 0:
                resumen_data_ce2 = []
                if res_ce['ce2_envio_soles'] > 0: resumen_data_ce2.append({'descripcion': 'Costo de envío de muestra - CE2.1', 'monto_soles': f"S/ {res_ce['ce2_envio_soles']:,.3f}", 'monto_dolares': f"US$ {res_ce['ce2_envio_dolares']:,.3f}"})
                if res_ce['ce2_lab_soles'] > 0: resumen_data_ce2.append({'descripcion': 'Costo de análisis de muestra - CE2.2', 'monto_soles': f"S/ {res_ce['ce2_lab_soles']:,.3f}", 'monto_dolares': f"US$ {res_ce['ce2_lab_dolares']:,.3f}"})
                resumen_data_ce2.append({'descripcion': 'Total', 'monto_soles': f"S/ {total_ce2_soles:,.3f}", 'monto_dolares': f"US$ {total_ce2_dolares:,.3f}"})
                subdoc_anexo_resumen_ce2 = create_table_subdoc(anexo_tpl, ["Descripción", "Monto (*) (S/)", "Monto (*) (US$)"], resumen_data_ce2, ['descripcion', 'monto_soles', 'monto_dolares'])

            resumen_data_total = [{'descripcion': 'Costo de muestreo - CE1', 'monto_soles': f"S/ {res_ce['ce1_soles']:,.3f}", 'monto_dolares': f"US$ {res_ce['ce1_dolares']:,.3f}"}, {'descripcion': 'Costo de análisis en un laboratorio acreditado - CE2', 'monto_soles': f"S/ {total_ce2_soles:,.3f}", 'monto_dolares': f"US$ {total_ce2_dolares:,.3f}"}, {'descripcion': 'Costo evitado total - CE', 'monto_soles': f"S/ {total_ce_soles:,.3f}", 'monto_dolares': f"US$ {total_ce_dolares:,.3f}"}]
            subdoc_anexo_resumen_total = create_table_subdoc(anexo_tpl, ["Descripción", "Monto (*) (S/)", "Monto (*) (US$)"], resumen_data_total, ['descripcion', 'monto_soles', 'monto_dolares'])

            contexto_anexo = {
                **contexto_final_word,
                'extremo': {
                    'numeral': 1,
                    'tipo': extremo.get('tipo_monitoreo_sel'),
                    'es_monitoreo_completo': extremo.get('tipo_servicio') == "Monitoreo completo (Muestreo y Análisis)",
                    # --- INICIO DE LA ADICIÓN (Nuevas Banderas) ---
                    'es_ruido': "Monitoreo de Ruido" in extremo.get('tipo_monitoreo_sel', ''),
                    'mostrar_ce2_envio_y_resumen': (
                        extremo.get('tipo_servicio') == "Monitoreo completo (Muestreo y Análisis)" and 
                        "Monitoreo de Ruido" not in extremo.get('tipo_monitoreo_sel', '')
                    ),
                    # --- FIN DE LA ADICIÓN ---
                    "texto_ipc_incumplimiento": res_ce.get('texto_ipc_incumplimiento', ''),
                    "texto_tc_incumplimiento": res_ce.get('texto_tc_incumplimiento', ''),
                    # --- INICIO ADICIÓN ---
                    "fuente_salario": res_ce.get('fuente_salario', ''),
                    "pdf_salario": res_ce.get('pdf_salario', ''),
                    "texto_ipc_costeo_salario": res_ce.get('texto_ipc_costeo_salario', ''),
                    "sustento_personal": res_ce.get('sustento_personal', ''),
                    # --- AÑADIR ESTOS ---
                    "sustento_seguros": res_ce.get('sustento_seguros', ''),
                    "sustento_epp": res_ce.get('sustento_epp', ''),
                    "sustento_movilidad": res_ce.get('sustento_movilidad', ''),
                    # --------------------
                },
                'tabla_consideraciones': tabla_consideraciones_anexo_subdoc,
                'tabla_anexo_ce1': subdoc_anexo_ce1,
                'tabla_anexo_ce2_envio': subdoc_anexo_ce2_envio,
                'tabla_anexo_ce2_lab': subdoc_anexo_ce2_lab,
                'tabla_anexo_resumen_ce2': subdoc_anexo_resumen_ce2,
                'tabla_anexo_resumen_total': subdoc_anexo_resumen_total
            }
            anexo_tpl.render(contexto_anexo)
            buffer_final_anexo = io.BytesIO()
            anexo_tpl.save(buffer_final_anexo)
            anexos_ce_generados.append(buffer_final_anexo)
    
    # 5. Retorno final
    resultados_app = {'extremos': [{'tipo': extremo.get('tipo_monitoreo_sel'), 'ce1_data': res_ce['ce1_items'], 'ce2_envio_data': res_ce['ce2_envio_items'], 'ce2_lab_data': res_ce['ce2_lab_items'], 'ce1_soles': res_ce['ce1_soles'], 'ce1_dolares': res_ce['ce1_dolares'], 'ce2_envio_soles': res_ce['ce2_envio_soles'], 'ce2_envio_dolares': res_ce['ce2_envio_dolares'], 'ce2_lab_soles': res_ce['ce2_lab_soles'], 'ce2_lab_dolares': res_ce['ce2_lab_dolares'], 'total_soles_extremo': res_ce['total_general_soles'], 'total_dolares_extremo': res_ce['total_general_dolares']}], 
                      'totales': {
                          'ce_total_soles': total_ce_soles, 'ce_total_dolares': total_ce_dolares, 
                          'beneficio_ilicito_uit': beneficio_ilicito_uit, 
                          'multa_final_uit': multa_uit, # Multa base
                          'bi_data_raw': res_bi.get('table_rows', []), 
                          'multa_data_raw': res_multa.get('multa_data_raw', []),
                          # --- INICIO: DATOS DE REDUCCIÓN PARA APP ---
                          'aplica_reduccion': aplica_reduccion_str,
                          'porcentaje_reduccion': porcentaje_str,
                          'multa_con_reduccion_uit': multa_con_reduccion_uit,
                          'multa_reducida_uit': multa_reducida_uit,
                          'multa_final_aplicada': multa_final_del_hecho_uit
                          # --- FIN: DATOS DE REDUCCIÓN PARA APP ---
                      }}
    
    return { 
        'contexto_final_word': contexto_final_word, 
        'doc_pre_compuesto': buffer_final_hecho,
        'resultados_para_app': resultados_app, 
        'usa_capacitacion': False, 
        'es_extemporaneo': False, 
        'ids_anexos': list(todos_los_anexos), 
        'anexos_ce_generados': anexos_ce_generados,
        # --- INICIO: DATOS DE REDUCCIÓN PARA APP ---
        'aplica_reduccion': aplica_reduccion_str,
        'porcentaje_reduccion': porcentaje_str,
        'multa_reducida_uit': multa_reducida_uit
        # --- FIN: DATOS DE REDUCCIÓN PARA APP ---
    }


def _procesar_hecho_multiple(datos_comunes, datos_hecho):
    """
    Processes a case with multiple extremes:
    1. Generates a complete CE annex for each extreme.
    2. Generates a single report body with consolidated BI, final Penalty, and consolidated Considerations table.
    (Version with corrected definition scope for datos_generales_notas)
    """
    # 1. Load templates
    df_tipificacion, id_infraccion = datos_comunes['df_tipificacion'], datos_comunes['id_infraccion']
    fila_infraccion = df_tipificacion[df_tipificacion['ID_Infraccion'] == id_infraccion].iloc[0]
    id_tpl_principal = fila_infraccion.get('ID_Plantilla_BI_Extremo') # Template for the body (BI/Penalty)
    id_tpl_anexo_ce = fila_infraccion.get('ID_Plantilla_CE_Extremo') # Template for CE annexes

    if not all([id_tpl_principal, id_tpl_anexo_ce]):
        return {'error': 'Missing template IDs for extremes in Tipificación.'}

    buffer_tpl_principal = descargar_archivo_drive(id_tpl_principal)
    buffer_tpl_anexo_ce = descargar_archivo_drive(id_tpl_anexo_ce)
    if not all([buffer_tpl_principal, buffer_tpl_anexo_ce]):
        return {'error': 'Failed to download templates for extremes.'}

    # 2. Initialize accumulators and lists
    total_bi_uit = 0
    lista_resultados_bi = []
    anexos_ce_generados = [] 
    todos_los_anexos_ids = set()
    resultados_app = {'extremos': [], 'totales': {}}
    lista_datos_consideraciones = [] 

    # 3. Iterate over each extreme to calculate BI, generate its CE ANNEX, and collect footnote data
    for i, extremo in enumerate(datos_hecho['extremos']):
        # Calculate CE for the extreme
        res_ce = _calcular_costo_evitado_monitoreo(datos_comunes, extremo)
        if res_ce.get('error'): continue

        # Calculate BI for the extreme and accumulate results
        texto_hecho_bi_extremo = f"{datos_hecho.get('texto_hecho', 'Incumplimiento')} - Extremo {i + 1}"
        datos_bi_extremo = { **datos_comunes, 'ce_soles': res_ce['total_general_soles'], 'ce_dolares': res_ce['total_general_dolares'], 'fecha_incumplimiento': extremo['fecha_incumplimiento'], 'texto_del_hecho': texto_hecho_bi_extremo }
        res_bi_extremo = calcular_beneficio_ilicito(datos_bi_extremo)
        if res_bi_extremo.get('error'): continue

        total_bi_uit += res_bi_extremo.get('beneficio_ilicito_uit', 0)
        lista_resultados_bi.append(res_bi_extremo) 
        todos_los_anexos_ids.update(res_ce['ids_anexos'])
        resultados_app['extremos'].append({ 
            'tipo': extremo.get('tipo_monitoreo_sel'),
            'ce1_data': res_ce['ce1_items'], 'ce2_envio_data': res_ce['ce2_envio_items'],
            'ce2_lab_data': res_ce['ce2_lab_items'], 'ce1_soles': res_ce['ce1_soles'],
            'ce1_dolares': res_ce['ce1_dolares'], 'ce2_envio_soles': res_ce['ce2_envio_soles'],
            'ce2_envio_dolares': res_ce['ce2_envio_dolares'], 'ce2_lab_soles': res_ce['ce2_lab_soles'],
            'ce2_lab_dolares': res_ce['ce2_lab_dolares'], 'total_soles_extremo': res_ce['total_general_soles'],
            'total_dolares_extremo': res_ce['total_general_dolares'], 'bi_data': res_bi_extremo.get('table_rows', []),
            'bi_uit': res_bi_extremo.get('beneficio_ilicito_uit', 0)
        })

        # --- TASK A: Collect data for the consolidated considerations table ---
        frecuencia_seleccionada = extremo.get('frecuencia_monitoreo', 'Trimestral')
        periodo = formatear_periodo_monitoreo(extremo.get('fecha_base'), frecuencia_seleccionada)
        matriz = extremo.get('tipo_monitoreo_sel', '').replace('Monitoreo de ', '').replace(' Ambiental', '')
        # --- TASK A: Collect data for the consolidated considerations table ---
        cantidad_puntos = extremo.get('cantidad', 1)
        nombres_puntos = extremo.get('nombres_puntos', '')
        # Capitalización de puntos
        puntos_descripcion = f"{texto_con_numero(cantidad_puntos).capitalize()} punto{'s' if cantidad_puntos > 1 else ''} de monitoreo: {nombres_puntos}"
        
        # Capitalización y formato para parámetros
        lista_params = extremo.get('parametros_seleccionados', [])
        cant_params = len(lista_params)
        parametros = f"{texto_con_numero(cant_params).capitalize()} parámetro{'s' if cant_params > 1 else ''}: {', '.join(lista_params)}"
        datos_tabla_consideraciones_extremo = {
            'Periodo del monitoreo': periodo,
            'Matriz': matriz,
            'Puntos de monitoreo': puntos_descripcion,
            'Parámetros': parametros
        }
        lista_datos_consideraciones.append(datos_tabla_consideraciones_extremo)

        # --- TASK B: Generate the CE ANNEX for this extreme ---
        doc_tpl_anexo = DocxTemplate(io.BytesIO(buffer_tpl_anexo_ce.getvalue()))

        # Create all CE tables for this extreme's annex
        tabla_consideraciones_anexo_subdoc = create_considerations_table_subdoc(
            doc_tpl_anexo,
            ["Periodo del monitoreo", "Matriz", "Puntos de monitoreo", "Parámetros"],
            [datos_tabla_consideraciones_extremo], # Pass only this extreme's data
            ['Periodo del monitoreo', 'Matriz', 'Puntos de monitoreo', 'Parámetros']
        )
        subdoc_anexo_ce1 = create_detailed_ce_table_subdoc(doc_tpl_anexo, res_ce['ce1_items'], res_ce['ce1_soles'], res_ce['ce1_dolares']) if res_ce['ce1_items'] else None

        subdoc_anexo_ce2_envio, subdoc_anexo_ce2_lab, subdoc_anexo_resumen_ce2, subdoc_anexo_resumen_total = None, None, None, None
        if res_ce['ce2_envio_items']:
            tabla_data_envio = [{**item, 'monto_soles': f"S/ {item['monto_soles']:,.3f}", 'monto_dolares': f"US$ {item['monto_dolares']:,.3f}"} for item in res_ce['ce2_envio_items']]
            tabla_data_envio.append({'descripcion': 'Subtotal Envío', 'monto_soles': f"S/ {res_ce['ce2_envio_soles']:,.3f}", 'monto_dolares': f"US$ {res_ce['ce2_envio_dolares']:,.3f}"})
            subdoc_anexo_ce2_envio = create_table_subdoc(doc_tpl_anexo, ["Descripción", "Unidad", "Cantidad", "Precio Unitario (S/)", "Factor de Ajuste", "Monto (S/)", "Monto (US$)"], tabla_data_envio, ['descripcion', 'unidad', 'cantidad', 'precio_unitario', 'factor_ajuste', 'monto_soles', 'monto_dolares'])
        if res_ce['ce2_lab_items']:
            tabla_data_lab = [{**item, 'monto_soles': f"S/ {item['monto_soles']:,.3f}", 'monto_dolares': f"US$ {item['monto_dolares']:,.3f}"} for item in res_ce['ce2_lab_items']]
            tabla_data_lab.append({'descripcion': 'Subtotal Análisis', 'monto_soles': f"S/ {res_ce['ce2_lab_soles']:,.3f}", 'monto_dolares': f"US$ {res_ce['ce2_lab_dolares']:,.3f}"})
            subdoc_anexo_ce2_lab = create_table_subdoc(doc_tpl_anexo, ["Descripción", "Unidad", "Cantidad", "Precio Unitario (S/)", "Factor de Ajuste", "Monto (S/)", "Monto (US$)"], tabla_data_lab, ['descripcion', 'unidad', 'cantidad', 'precio_unitario', 'factor_ajuste', 'monto_soles', 'monto_dolares'])

        total_ce2_soles = res_ce['ce2_envio_soles'] + res_ce['ce2_lab_soles']
        total_ce2_dolares = res_ce['ce2_envio_dolares'] + res_ce['ce2_lab_dolares']
        if total_ce2_soles > 0:
            resumen_data_ce2 = []
            if res_ce['ce2_envio_soles'] > 0: resumen_data_ce2.append({'descripcion': 'Costo por Envío de Muestras', 'monto_soles': f"S/ {res_ce['ce2_envio_soles']:,.3f}", 'monto_dolares': f"US$ {res_ce['ce2_envio_dolares']:,.3f}"})
            if res_ce['ce2_lab_soles'] > 0: resumen_data_ce2.append({'descripcion': 'Costo por Análisis de Laboratorio', 'monto_soles': f"S/ {res_ce['ce2_lab_soles']:,.3f}", 'monto_dolares': f"US$ {res_ce['ce2_lab_dolares']:,.3f}"})
            resumen_data_ce2.append({'descripcion': 'Total Costo Evitado (CE2)', 'monto_soles': f"S/ {total_ce2_soles:,.3f}", 'monto_dolares': f"US$ {total_ce2_dolares:,.3f}"})
            subdoc_anexo_resumen_ce2 = create_table_subdoc(doc_tpl_anexo, ["Descripción", "Monto (S/)", "Monto (US$)"], resumen_data_ce2, ['descripcion', 'monto_soles', 'monto_dolares'])
        resumen_data_total = [{'descripcion': 'Total Costo Evitado (CE1)', 'monto_soles': f"S/ {res_ce['ce1_soles']:,.3f}", 'monto_dolares': f"US$ {res_ce['ce1_dolares']:,.3f}"}, {'descripcion': 'Total Costo Evitado (CE2)', 'monto_soles': f"S/ {total_ce2_soles:,.3f}", 'monto_dolares': f"US$ {total_ce2_dolares:,.3f}"}, {'descripcion': 'Costo Evitado Total (CE)', 'monto_soles': f"S/ {res_ce['total_general_soles']:,.3f}", 'monto_dolares': f"US$ {res_ce['total_general_dolares']:,.3f}"}]
        subdoc_anexo_resumen_total = create_table_subdoc(doc_tpl_anexo, ["Descripción", "Monto (S/)", "Monto (US$)"], resumen_data_total, ['descripcion', 'monto_soles', 'monto_dolares'])

        contexto_anexo = {
            **datos_comunes['context_data'],
            'hecho': { 'numero_imputado': datos_comunes['numero_hecho_actual'] },
            'extremo': {
                'numeral': i + 1, # Use i here for annex numbering
                'tipo': extremo.get('tipo_monitoreo_sel'),
                'es_monitoreo_completo': extremo.get('tipo_servicio') == "Monitoreo completo (Muestreo y Análisis)",
                # --- INICIO DE LA ADICIÓN (Nuevas Banderas) ---
                'es_ruido': "Monitoreo de Ruido" in extremo.get('tipo_monitoreo_sel', ''),
                'mostrar_ce2_envio_y_resumen': (
                    extremo.get('tipo_servicio') == "Monitoreo completo (Muestreo y Análisis)" and 
                    "Monitoreo de Ruido" not in extremo.get('tipo_monitoreo_sel', '')
                ),
                # --- FIN DE LA ADICIÓN ---
                "texto_ipc_incumplimiento": res_ce.get('texto_ipc_incumplimiento', ''),
                "texto_tc_incumplimiento": res_ce.get('texto_tc_incumplimiento', ''),
                "fuente_salario": res_ce.get('fuente_salario', ''),
                "pdf_salario": res_ce.get('pdf_salario', ''),
                "texto_ipc_costeo_salario": res_ce.get('texto_ipc_costeo_salario', ''),
            },
            'tabla_consideraciones': tabla_consideraciones_anexo_subdoc,
            'tabla_anexo_ce1': subdoc_anexo_ce1,
            'tabla_anexo_ce2_envio': subdoc_anexo_ce2_envio,
            'tabla_anexo_ce2_lab': subdoc_anexo_ce2_lab,
            'tabla_anexo_resumen_ce2': subdoc_anexo_resumen_ce2,
            'tabla_anexo_resumen_total': subdoc_anexo_resumen_total
        }

        # Render and save the annex to the list
        doc_tpl_anexo.render(contexto_anexo)
        buffer_anexo_final = io.BytesIO()
        doc_tpl_anexo.save(buffer_anexo_final)
        anexos_ce_generados.append(buffer_anexo_final)

    # --- Footnote logic OUTSIDE the loop ---
    # Paso 1: Recopilar textos únicos y mapeo de (clave_original, indice_extremo) a texto
    notas_a_mapear = {} 
    map_clave_a_texto = {} 

    # Define datos_generales_notas HERE, using the now populated lista_resultados_bi
    datos_generales_notas = lista_resultados_bi[0].get('footnote_data', {}) if lista_resultados_bi else {}
    id_infraccion_hecho = datos_comunes.get('id_infraccion')

    # Iterate over the ALREADY POPULATED lista_resultados_bi to gather footnote texts
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

    # Step 2: Sort unique texts based on desired key order and assign letters
    desired_key_order = ['ce_anexo', 'cok', 'periodo_bi', 'bcrp', 'ipc_fecha', 'sunat']
    mapeo_texto_a_letra_final = {}
    letra_actual_code = ord('a')
    textos_ya_mapeados = set()

    for clave in desired_key_order:
        # Find texts generated by this key across all extremes
        textos_de_esta_clave = set()
        for (k, idx), txt in map_clave_a_texto.items():
            if k == clave:
                textos_de_esta_clave.add(txt)
        
        # Sort the unique texts for this key
        textos_ordenados = sorted(list(textos_de_esta_clave))

        # Assign letters sequentially
        for texto in textos_ordenados:
            if texto not in textos_ya_mapeados:
                letra_final = chr(letra_actual_code)
                mapeo_texto_a_letra_final[texto] = letra_final
                textos_ya_mapeados.add(texto)
                letra_actual_code += 1

    # Step 3: Generate the final list of footnotes for printing, respecting the desired order
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


    # 4. (OUTSIDE THE LOOP) Generate the REPORT BODY with consolidated BI/Penalty
    if not lista_resultados_bi: return {'error': 'Could not calculate BI for any extreme.'}

    # --- Recuperar Factor F de la interfaz ---
    factor_f = datos_hecho.get('factor_f_calculado', 1.0)

    res_multa_final = calcular_multa({
        **datos_comunes, 
        'beneficio_ilicito': total_bi_uit,
        'factor_f': factor_f # <--- PASAR EL FACTOR
    })
    multa_final_uit = res_multa_final.get('multa_final_uit', 0)
    
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

    doc_tpl_principal = DocxTemplate(io.BytesIO(buffer_tpl_principal.getvalue()))

    # --- TASK D: Create the CONSOLIDATED considerations table ---
    tabla_consideraciones_consolidada_subdoc = create_considerations_table_subdoc(
        doc_tpl_principal,
        ["Periodo del monitoreo", "Matriz", "Puntos de monitoreo", "Parámetros"],
        lista_datos_consideraciones, 
        ['Periodo del monitoreo', 'Matriz', 'Puntos de monitoreo', 'Parámetros']
    )

    # Create the CONSOLIDATED BI table (passing the correct maps)
    tabla_bi_consolidada_subdoc = create_consolidated_bi_table_subdoc(
        doc_tpl_principal,
        lista_resultados_bi,
        total_bi_uit,
        footnotes_data=footnotes_data_bi,
        map_texto_a_letra=mapeo_texto_a_letra_final,
        map_clave_a_texto=map_clave_a_texto
    )

    # Create the FINAL PENALTY table
    tabla_multa_final_subdoc = create_main_table_subdoc(doc_tpl_principal, ["Componentes", "Monto"], res_multa_final.get('multa_data_raw', []), ['Componentes', 'Monto'])

    # Assemble the FINAL context
    contexto_final = {
        **datos_comunes['context_data'],
        'hecho': {
            'numero_imputado': datos_comunes['numero_hecho_actual'],
            'descripcion': RichText(datos_hecho.get('texto_hecho', '')),
            'tabla_consideraciones_consolidada': tabla_consideraciones_consolidada_subdoc,
            'tabla_bi_consolidada': tabla_bi_consolidada_subdoc,
            'tabla_multa_final': tabla_multa_final_subdoc
        },
        'numeral_hecho': f"IV.{datos_comunes['numero_hecho_actual']}",
        'multa_original_uit': f"{multa_final_uit:,.3f} UIT",
        'mh_uit': f"{multa_final_del_hecho_uit:,.3f} UIT",
        'bi_uit': f"{total_bi_uit:,.3f} UIT",

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

    # Render the report body ONCE
    doc_tpl_principal.render(contexto_final)
    buffer_final_hecho = io.BytesIO()
    doc_tpl_principal.save(buffer_final_hecho)
    buffer_final_hecho.seek(0)

    resultados_app['totales'] = {
        'beneficio_ilicito_uit': total_bi_uit,
        'multa_final_uit': multa_final_uit, # Multa base
        'bi_data_raw': lista_resultados_bi, # <-- CORRECCIÓN: Pasar la lista completa de BIs
        'multa_data_raw': res_multa_final.get('multa_data_raw', []),
        # --- INICIO: DATOS DE REDUCCIÓN PARA APP ---
        'aplica_reduccion': aplica_reduccion_str,
        'porcentaje_reduccion': porcentaje_str,
        'multa_con_reduccion_uit': multa_con_reduccion_uit,
        'multa_reducida_uit': multa_reducida_uit,
        'multa_final_aplicada': multa_final_del_hecho_uit
        # --- FIN: DATOS DE REDUCCIÓN PARA APP ---
    }

    return {
        'doc_pre_compuesto': buffer_final_hecho,
        'resultados_para_app': resultados_app,
        'usa_capacitacion': False,
        'es_extemporaneo': False,
        'anexos_ce_generados': anexos_ce_generados,
        'ids_anexos': list(todos_los_anexos_ids),
        # --- INICIO: DATOS DE REDUCCIÓN PARA APP ---
        'aplica_reduccion': aplica_reduccion_str,
        'porcentaje_reduccion': porcentaje_str,
        'multa_reducida_uit': multa_reducida_uit
        # --- FIN: DATOS DE REDUCCIÓN PARA APP ---
    }

# --- PROCESADOR PRINCIPAL ---
def procesar_infraccion(datos_comunes, datos_hecho):
    """
    Determines whether the case is simple (1 extremo) or multiple (>1 extremo)
    and calls the corresponding processing function.
    """
    num_extremos = len(datos_hecho.get('extremos', []))

    if num_extremos == 1:
        # If there is only one instance, call the simple processing function.
        return _procesar_hecho_simple(datos_comunes, datos_hecho)
    elif num_extremos > 1:
        # If there are multiple instances, call the advanced function for multiple extremos.
        return _procesar_hecho_multiple(datos_comunes, datos_hecho)
    else:
        # If no instances have been added, return an error.
        return {'error': 'No se ha registrado ningún extremo para este hecho.'}