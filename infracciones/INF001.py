# --- Archivo: infracciones/INF001.py ---

import streamlit as st
import pandas as pd
import io
from babel.dates import format_date
from jinja2 import Environment
from num2words import num2words
from docx import Document
from docxcompose.composer import Composer
from docxtpl import DocxTemplate, RichText
from datetime import date, timedelta
import holidays

# --- IMPORTACIONES DE MÓDULOS PROPIOS ---
from textos_manager import obtener_fuente_formateada
from funciones import (create_main_table_subdoc, create_table_subdoc,
                     texto_con_numero, create_footnotes_subdoc, redondeo_excel, format_decimal_dinamico) # <-- AÑADIDO
from sheets import (calcular_beneficio_ilicito, calcular_multa,
                    descargar_archivo_drive,
                    calcular_beneficio_ilicito_extemporaneo)

# ---------------------------------------------------------------------
# FUNCIÓN AUXILIAR: CÁLCULO CE COMPLETO (CE1)
# ---------------------------------------------------------------------

def _calcular_costo_evitado_inf001(datos_comunes, datos_hecho_general, extremo_data):
    """
    Calcula el CE completo para INF001.
    - CE1: Sistematización (Horas proporcionales a secciones faltantes).
    """
    result = {
        'ce1_data_raw': [], 'ce1_soles': 0.0, 'ce1_dolares': 0.0,
        'ce_soles_para_bi': 0.0, 'ce_dolares_para_bi': 0.0,
        'ids_anexos': set(),
        'fuentes': {'ce1': {}},
        'error': None
    }
    SECCIONES_TOTALES_IAA = 12 

    try:
        # 1. Datos Generales
        tipo_presentacion = extremo_data.get('tipo_presentacion')
        fecha_incumplimiento = extremo_data.get('fecha_incumplimiento')
        num_secciones_faltantes = 12 if tipo_presentacion == "No presentó" else extremo_data.get('num_secciones_faltantes', 0)
        if not fecha_incumplimiento: raise ValueError("Falta fecha incumplimiento extremo.")
        
        # 2. Calcular CE1 (Sistematización)
        fecha_calculo_ce1 = fecha_incumplimiento

        def _calcular_ce1_interno(datos_comunes_ce1, fecha_final, secciones_faltantes):
            res_int = {'items_calculados': [], 'error': None, 'fuentes': {}}
            try:
                # Carga de datos
                df_items_inf = datos_comunes_ce1.get('df_items_infracciones')
                df_costos = datos_comunes_ce1.get('df_costos_items')
                df_coti = datos_comunes_ce1.get('df_coti_general')
                df_sal = datos_comunes_ce1.get('df_salarios_general')
                df_ind = datos_comunes_ce1.get('df_indices')
                id_rubro_ce1 = datos_comunes_ce1.get('id_rubro_seleccionado')
                id_inf_ce1 = 'INF001'
                
                if any(df is None for df in [df_items_inf, df_costos, df_coti, df_sal, df_ind]): raise ValueError("Faltan DataFrames CE1.")
                
                # IPC/TC fecha incumplimiento
                fecha_final_dt_ce1 = pd.to_datetime(fecha_final, errors='coerce')
                ipc_row_ce1 = df_ind[df_ind['Indice_Mes'].dt.to_period('M') == fecha_final_dt_ce1.to_period('M')]
                if ipc_row_ce1.empty: raise ValueError(f"Sin IPC/TC CE1 para {fecha_final_dt_ce1.strftime('%B %Y')}")
                ipc_inc_ce1, tc_inc_ce1 = ipc_row_ce1.iloc[0]['IPC_Mensual'], ipc_row_ce1.iloc[0]['TC_Mensual']
                
                fuentes_ce1 = {'placeholders_dinamicos': {}}
                items_ce1 = []
                receta_ce1 = df_items_inf[df_items_inf['ID_Infraccion'] == id_inf_ce1]
                
                # Lógica de Horas Proporcionales
                total_horas_prof = 0
                total_horas_laptop = 0
                for _, item_r in receta_ce1.iterrows():
                    if item_r.get('Tipo_Costo') == 'Remision':
                         horas_r = float(item_r.get('Cantidad_Horas', 0))
                         # Convertimos a minúscula para asegurar la coincidencia
                         nombre_item_r = str(item_r.get('Nombre_Item', '')).lower()
                         
                         if 'profesional' in nombre_item_r: 
                             total_horas_prof = horas_r
                         elif 'laptop' in nombre_item_r or 'computadora' in nombre_item_r: 
                             total_horas_laptop = horas_r

                horas_x_secc_prof = total_horas_prof / SECCIONES_TOTALES_IAA if SECCIONES_TOTALES_IAA > 0 else 0
                horas_x_secc_laptop = total_horas_laptop / SECCIONES_TOTALES_IAA if SECCIONES_TOTALES_IAA > 0 else 0
                
                # Horas finales calculadas
                horas_ce1_prof = secciones_faltantes * horas_x_secc_prof
                horas_ce1_laptop = secciones_faltantes * horas_x_secc_laptop

                # Bucle de Cálculo
                for _, item_receta in receta_ce1.iterrows():
                    if item_receta.get('Tipo_Costo') != 'Remision': continue
                    
                    # 1. Extraemos la cadena con comas y creamos una lista limpia
                    id_items_str = str(item_receta.get('ID_Item', ''))
                    lista_ids_buscar = [x.strip() for x in id_items_str.split(',') if x.strip()]
                    
                    desc_item = item_receta.get('Nombre_Item', 'N/A')
                    
                    # 2. Filtramos la base de datos usando .isin() con la lista de IDs
                    # IMPORTANTE: Ahora busca en la columna 'ID_Item' de Costos_Items
                    costos_posibles = df_costos[df_costos['ID_Item'].isin(lista_ids_buscar)].copy()
                    
                    if costos_posibles.empty: continue
                    
                    tipo_item = item_receta.get('Tipo_Item')
                    df_candidatos = pd.DataFrame()
                    if tipo_item == 'Variable':
                         id_rubro_str = str(id_rubro_ce1) if id_rubro_ce1 else ''
                         df_candidatos = costos_posibles[costos_posibles['ID_Rubro'].astype(str).str.contains(fr'\b{id_rubro_str}\b', regex=True, na=False)].copy()
                         if df_candidatos.empty: df_candidatos = costos_posibles[costos_posibles['ID_Rubro'].isin(['', 'nan', None])].copy()
                    else: df_candidatos = costos_posibles.copy()
                    
                    if df_candidatos.empty: continue
                    
                    # Filtrar fechas
                    fechas_fuente = []
                    for _, cand in df_candidatos.iterrows():
                        id_gen = cand['ID_General']; fecha_f = pd.NaT
                        if pd.notna(id_gen):
                            if 'SAL' in id_gen: f=df_sal[df_sal['ID_Salario']==id_gen]; fecha_f=pd.to_datetime(f"{int(f.iloc[0]['Costeo_Salario'])}-12-31", errors='coerce') if not f.empty else pd.NaT
                            elif 'COT' in id_gen: f=df_coti[df_coti['ID_Cotizacion']==id_gen]; fecha_f=pd.to_datetime(f.iloc[0]['Fecha_Costeo'], errors='coerce') if not f.empty else pd.NaT
                        fechas_fuente.append(fecha_f)
                    df_candidatos['Fecha_Fuente'] = fechas_fuente
                    df_candidatos.dropna(subset=['Fecha_Fuente'], inplace=True)
                    if df_candidatos.empty: continue
                    
                    # Seleccionar mejor fecha
                    df_candidatos['Diferencia_Dias'] = (df_candidatos['Fecha_Fuente'] - fecha_final_dt_ce1).dt.days.abs()
                    costo_final = df_candidatos.loc[df_candidatos['Diferencia_Dias'].idxmin()]
                    
                    # Datos del Costo
                    id_gen = costo_final['ID_General']; fecha_f = costo_final['Fecha_Fuente']
                    ipc_cost, tc_cost = 0.0, 0.0
                    
                    if pd.notna(id_gen) and 'SAL' in id_gen:
                        idx_anio = df_ind[df_ind['Indice_Mes'].dt.year == fecha_f.year]
                        ipc_cost, tc_cost = (float(idx_anio['IPC_Mensual'].mean()), float(idx_anio['TC_Mensual'].mean())) if not idx_anio.empty else (0,0)
                        
                        f_row = df_sal[df_sal['ID_Salario']==id_gen]
                        if not f_row.empty:
                             fuentes_ce1['fuente_salario'] = f_row.iloc[0].get('Fuente_Salario','')
                             fuentes_ce1['pdf_salario'] = f_row.iloc[0].get('PDF_Salario','')
                             # --- NUEVO: Placeholder IPC Promedio Salario ---
                             fuentes_ce1['texto_ipc_costeo_salario'] = f"Promedio {fecha_f.year}, IPC = {ipc_cost:,.6f}"     
                    elif pd.notna(id_gen) and 'COT' in id_gen:
                         idx_row = df_ind[df_ind['Indice_Mes'].dt.to_period('M') == fecha_f.to_period('M')]
                         ipc_cost, tc_cost = (float(idx_row.iloc[0]['IPC_Mensual']), float(idx_row.iloc[0]['TC_Mensual'])) if not idx_row.empty else (0,0)
                    
                    if ipc_cost == 0: continue

                    # --- CORRECCIÓN 2: Solo capturar si es el ítem Profesional ---
                    if 'Profesional' in desc_item:
                        sustento_txt = costo_final.get('Sustento_Item')
                        if pd.notna(sustento_txt) and str(sustento_txt).strip():
                            fuentes_ce1['sustento_item_profesional'] = str(sustento_txt).strip()
                    # Cálculo de Montos
                    costo_orig = float(costo_final.get('Costo_Unitario_Item', 0.0))
                    moneda = costo_final.get('Moneda_Item')
                    
                    precio_s = costo_orig if moneda == 'S/' else costo_orig * tc_cost
                    factor = redondeo_excel(ipc_inc_ce1 / ipc_cost, 3) if ipc_cost > 0 else 0
                    cant = float(item_receta.get('Cantidad_Recursos', 1.0))
                    
                    # Asignar Horas calculadas
                    horas_a_usar = 0.0
                    unidad_texto = "Und"
                    # Convertimos a minúscula para asegurar la coincidencia
                    desc_item_lower = str(desc_item).lower()
                    
                    if 'profesional' in desc_item_lower: 
                        horas_a_usar = horas_ce1_prof
                        unidad_texto = f"{redondeo_excel(horas_a_usar, 2)} horas"
                    elif 'laptop' in desc_item_lower or 'computadora' in desc_item_lower: 
                        horas_a_usar = horas_ce1_laptop
                        unidad_texto = f"{redondeo_excel(horas_a_usar, 2)} horas" # Opcional: Mostrar horas también en la unidad
                    else: 
                        horas_a_usar = float(item_receta.get('Cantidad_Horas', 0.0))
                    
                    # Fórmula con Redondeo Correcto
                    monto_s_raw = cant * horas_a_usar * precio_s * factor
                    monto_s = redondeo_excel(monto_s_raw, 3)
                    monto_d = redondeo_excel(monto_s / tc_inc_ce1 if tc_inc_ce1 > 0 else 0, 3)
                    
                    items_ce1.append({
                        "descripcion": desc_item, 
                        "cantidad": cant, 
                        "unidad": unidad_texto, 
                        "horas": horas_a_usar, 
                        "precio_soles": precio_s, 
                        "factor_ajuste": factor, 
                        "monto_soles": monto_s, 
                        "monto_dolares": monto_d, 
                        "id_anexo": costo_final.get('ID_Anexo_Drive')
                    })
                
                # Fuentes adicionales
                res_int['items_calculados'] = items_ce1
                res_int['fuentes'] = fuentes_ce1
                
                # --- CORRECCIÓN: Añadir fi_mes ---
                res_int['fuentes']['fi_mes'] = format_date(fecha_final_dt_ce1, "MMMM 'de' yyyy", locale='es')
                # ---------------------------------
                
                res_int['fuentes']['fi_ipc'] = float(ipc_inc_ce1)
                res_int['fuentes']['fi_tc'] = float(tc_inc_ce1)
                
            except Exception as e_int: res_int['error'] = f"Error interno CE1: {e_int}"
            return res_int

        # Ejecutar CE1
        res_ce1 = _calcular_ce1_interno(datos_comunes, fecha_calculo_ce1, num_secciones_faltantes)
        if res_ce1.get('error'): result['error'] = f"CE1: {res_ce1['error']}"; return result
        
        result['ce1_data_raw'] = res_ce1.get('items_calculados', [])
        result['ce1_soles'] = sum(i.get('monto_soles',0) for i in result['ce1_data_raw'])
        result['ce1_dolares'] = sum(i.get('monto_dolares',0) for i in result['ce1_data_raw'])
        result['ids_anexos'].update(i.get('id_anexo') for i in result['ce1_data_raw'] if i.get('id_anexo'))
        result['fuentes']['ce1'] = res_ce1.get('fuentes', {})

        # 4. Consolidar para BI (Solo CE1)
        result['ce_soles_para_bi'] = result['ce1_soles']
        result['ce_dolares_para_bi'] = result['ce1_dolares']

        return result

    except Exception as e:
        import traceback; traceback.print_exc()
        result['error'] = f"Error crítico en _calcular_costo_evitado_inf001: {e}"
        return result
    
    
# ---------------------------------------------------------------------
# FUNCIÓN 2: RENDERIZAR INPUTS (Sin cambios mayores, solo firma)
# ---------------------------------------------------------------------
def renderizar_inputs_especificos(i, df_dias_no_laborables=None):
    st.markdown("##### Detalles de la infracción")
    datos_hecho = st.session_state.imputaciones_data[i]

    st.markdown("###### **" \
    "Extremos del incumplimiento**")
    if 'extremos' not in datos_hecho: datos_hecho['extremos'] = [{}]
    if st.button("➕ Añadir Extremo", key=f"add_extremo_iaa_{i}"): datos_hecho['extremos'].append({}); st.rerun()

    for j, extremo in enumerate(datos_hecho['extremos']):
        with st.container(border=True):
            st.markdown(f"**1. Detalles del Informe Ambiental Anual (IAA)**")
            
            # --- PARTE A: FECHAS AUTOMÁTICAS ---
            col_anio, col_metrica = st.columns([2, 2])
            with col_anio:
                anio_iaa = st.number_input(
                    "Año correspondiente al IAA:", 
                    min_value=2000, max_value=date.today().year, step=1, 
                    key=f"anio_iaa_{i}_{j}", 
                    value=extremo.get('anio_iaa', date.today().year - 1)
                )
                extremo['anio_iaa'] = anio_iaa
                fecha_max_calc = date(anio_iaa + 1, 3, 30)
                fecha_inc_calc = date(anio_iaa + 1, 3, 31)
                extremo['fecha_maxima_presentacion'] = fecha_max_calc
                extremo['fecha_incumplimiento'] = fecha_inc_calc

            with col_metrica:
                st.info(f"**Fecha límite de presentación:** {fecha_max_calc.strftime('%d/%m/%Y')}\n\n**Fecha de incumplimiento:** {fecha_inc_calc.strftime('%d/%m/%Y')}")

            tipo_presentacion = st.radio(
                "Tipo de incumplimiento:", ["No presentó", "Presentó incompleto"], 
                key=f"tipo_presentacion_iaa_{i}_{j}", 
                index=0 if extremo.get('tipo_presentacion') == "No presentó" else 1, horizontal=True
            )
            extremo['tipo_presentacion'] = tipo_presentacion
            
            if tipo_presentacion == "Presentó incompleto":
                extremo['num_secciones_faltantes'] = st.number_input(
                    "Secciones faltantes/incompletas (de 12):", min_value=1, max_value=12, step=1, 
                    key=f"num_secciones_{i}_{j}", value=extremo.get('num_secciones_faltantes', 1)
                )
            else:
                extremo['num_secciones_faltantes'] = 12

            if st.button(f"🗑️ Eliminar extremo", key=f"del_extremo_{i}_{j}"):
                datos_hecho['extremos'].pop(j)
                st.rerun()
    return datos_hecho

# --- 3. VALIDACIÓN ---
def validar_inputs(datos_hecho):
    if not datos_hecho.get('extremos'): return False
    for extremo in datos_hecho.get('extremos', []):
        if not all([ extremo.get('anio_iaa'), extremo.get('fecha_incumplimiento'), extremo.get('tipo_presentacion') ]): return False
    return True

# --- 4. PROCESADOR PRINCIPAL ---
def procesar_infraccion(datos_comunes, datos_hecho):
    num_extremos = len(datos_hecho.get('extremos', []))
    if num_extremos == 1: return _procesar_hecho_simple(datos_comunes, datos_hecho)
    elif num_extremos > 1: return _procesar_hecho_multiple(datos_comunes, datos_hecho)
    else: return {'error': 'No se ha registrado ningún Año IAA.'}

# ---------------------------------------------------------------------
# FUNCIÓN 5: PROCESAR HECHO SIMPLE (1 EXTREMO)
# ---------------------------------------------------------------------
def _procesar_hecho_simple(datos_comunes, datos_hecho):
    """
    Procesa un hecho INF001 con un único año IAA.
    Tablas alineadas al formato INF008.
    """
    try:
        # --- INICIO CORRECCIÓN ---
        jinja_env = Environment(trim_blocks=True, lstrip_blocks=True)
        # --- FIN CORRECCIÓN ---
        # 1. Cargar plantillas
        df_tipificacion, id_infraccion = datos_comunes['df_tipificacion'], 'INF001'
        fila_inf = df_tipificacion[df_tipificacion['ID_Infraccion'] == id_infraccion].iloc[0]
        id_tpl_bi, id_tpl_ce = fila_inf.get('ID_Plantilla_BI'), fila_inf.get('ID_Plantilla_CE')
        
        if not id_tpl_bi or not id_tpl_ce: 
            return {'error': f'Faltan IDs plantilla simple (BI o CE) para {id_infraccion}.'}
            
        buf_bi, buf_ce = descargar_archivo_drive(id_tpl_bi), descargar_archivo_drive(id_tpl_ce)
        
        if not buf_bi or not buf_ce: 
            return {'error': f'Fallo descarga plantilla simple para {id_infraccion}.'}
            
        doc_tpl_bi = DocxTemplate(buf_bi)
        tpl_anexo = DocxTemplate(buf_ce)

        # 2. Calcular CE
        extremo = datos_hecho['extremos'][0]
        fecha_inc = extremo.get('fecha_incumplimiento')
        tipo_pres = extremo.get('tipo_presentacion')
        num_secciones_faltantes = 12 if tipo_pres == "No presentó" else extremo.get('num_secciones_faltantes', 0)
        
        res_ce = _calcular_costo_evitado_inf001(datos_comunes, datos_hecho, extremo)
        if res_ce.get('error'): return {'error': f"Error CE: {res_ce['error']}"}
                
        aplicar_capacitacion = False
        label_ce_principal = "CE"

        # 3. Calcular BI y Multa
        texto_hecho_bi = datos_hecho.get('texto_hecho', '')
        
        datos_bi_base = {
            **datos_comunes, 
            'ce_soles': res_ce['ce_soles_para_bi'], 
            'ce_dolares': res_ce['ce_dolares_para_bi'], 
            'fecha_incumplimiento': fecha_inc, 
            'texto_del_hecho': texto_hecho_bi
        }
        
        res_bi = calcular_beneficio_ilicito(datos_bi_base)
        es_ext = False 
        if not res_bi or res_bi.get('error'): return res_bi or {'error': 'Error BI.'}
        bi_uit = res_bi.get('beneficio_ilicito_uit', 0)

        # --- Factor F ---
        factor_f = datos_hecho.get('factor_f_calculado', 1.0)

        res_multa = calcular_multa({
            **datos_comunes, 
            'beneficio_ilicito': bi_uit,
            'factor_f': factor_f
        })
        multa_uit = res_multa.get('multa_final_uit', 0)

        # --- Reducción y Tope ---
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

        # 4. Generar Tablas Cuerpo
        
        # --- Formatear datos CE1 (con estructura INF008) ---
        ce1_fmt = []
        for idx, item in enumerate(res_ce['ce1_data_raw'], 1):
            ce1_fmt.append({
                'descripcion': f"{item.get('descripcion', '')} {idx}/",
                'cantidad': format_decimal_dinamico(item.get('cantidad', 0)),
                'horas': format_decimal_dinamico(item.get('horas', 0)),
                'unidad': item.get('unidad', ''),
                'precio_soles': f"S/ {item.get('precio_soles', 0):,.3f}",
                'factor_ajuste': f"{item.get('factor_ajuste', 0):,.3f}",
                'monto_soles': f"S/ {item.get('monto_soles', 0):,.3f}",
                'monto_dolares': f"US$ {item.get('monto_dolares', 0):,.3f}"
            })
        ce1_fmt.append({
            'descripcion': f"Total", 
            'monto_soles': f"S/ {res_ce['ce1_soles']:,.3f}", 
            'monto_dolares': f"US$ {res_ce['ce1_dolares']:,.3f}"
        })
        

        # --- SOLUCIÓN: Compactar Notas BI ---
        filas_bi_crudas, fn_map_orig, fn_data = res_bi.get('table_rows', []), res_bi.get('footnote_mapping', {}), res_bi.get('footnote_data', {})

        # Identificar letras realmente usadas
        letras_usadas = sorted(list({r for f in filas_bi_crudas if f.get('ref') for r in f.get('ref').replace(" ", "").split(",") if r}))
        
        letras_base = "abcdefghijklmnopqrstuvwxyz"
        map_traduccion = {v: letras_base[i] for i, v in enumerate(letras_usadas)}
        nuevo_fn_map = {map_traduccion[v]: fn_map_orig[v] for v in letras_usadas if v in fn_map_orig}

        filas_bi_para_tabla = []
        for fila in filas_bi_crudas:
            nueva_fila = fila.copy()
            ref_orig = nueva_fila.get('ref', '')
            super_final = str(nueva_fila.get('descripcion_superindice', ''))
            if ref_orig:
                nuevas = [map_traduccion[r] for r in ref_orig.replace(" ", "").split(",") if r in map_traduccion]
                if nuevas: super_final += f"({', '.join(nuevas)})"
            
            nueva_fila['descripcion_texto'] = str(nueva_fila.get('descripcion_texto', ''))
            nueva_fila['descripcion_superindice'] = super_final
            filas_bi_para_tabla.append(nueva_fila)

        fn_list = [f"({l}) {obtener_fuente_formateada(k, fn_data, id_infraccion, es_ext)}" for l, k in sorted(nuevo_fn_map.items())]
        fn_data_dict = {'list': fn_list, 'elaboration': 'Elaboración: Subdirección de Sanción y Gestión de Incentivos (SSAG) - DFAI.', 'style': 'FuenteTabla'}
        tabla_bi = create_main_table_subdoc(doc_tpl_bi, ["Descripción", "Monto"], filas_bi_para_tabla, ['descripcion_texto', 'monto'], footnotes_data=fn_data_dict, column_widths=(5, 1))
        
        tabla_multa = create_main_table_subdoc(doc_tpl_bi, ["Componentes", "Monto"], res_multa.get('multa_data_raw', []), ['Componentes', 'Monto'], texto_posterior="Elaboración: Subdirección de Sanción y Gestión de Incentivos (SSAG) - DFAI.", estilo_texto_posterior='FuenteTabla', column_widths=(5, 1))

        # --- Placeholders de Horas y Secciones ---
        horas_secciones = 0
        for item in res_ce.get('ce1_data_raw', []):
            if 'Profesional' in item.get('descripcion', ''):
                horas_secciones = item.get('horas', 0)
                break
        texto_horas_secciones = texto_con_numero(horas_secciones, genero='f')
        sufijo_hora = "hora" if horas_secciones == 1 else "horas"
        ph_horas_secciones = f"{texto_horas_secciones} {sufijo_hora}"
        
        texto_num_secciones = texto_con_numero(num_secciones_faltantes, genero='f')
        sufijo_seccion = "sección" if num_secciones_faltantes == 1 else "secciones"
        ph_secciones_texto = f"{texto_num_secciones} {sufijo_seccion}"

        # 5. Contexto Final
        fuentes_ce = res_ce.get('fuentes', {})
        contexto_word = {
            **datos_comunes['context_data'], 
            'es_no_presento': tipo_pres == "No presentó", # <--- Condicional requerido
            'ph_ipc_promedio_salario_ce1': fuentes_ce.get('ce1', {}).get('texto_ipc_costeo_salario', ''), # <--- IPC Salario
            # ... resto de campos
            'acronyms': datos_comunes['acronym_manager'],
            'hecho': {'numero_imputado': datos_comunes['numero_hecho_actual'], 'descripcion': RichText(datos_hecho.get('texto_hecho', ''))},
            'numeral_hecho': f"IV.{datos_comunes['numero_hecho_actual'] + 1}",
            'tipo_presentacion_iaa': tipo_pres, 
            'num_secciones_faltantes_iaa': num_secciones_faltantes, 
            'num_secciones_faltantes_texto': ph_secciones_texto,
            'horas_secciones_faltantes_texto': ph_horas_secciones,
            'anio_iaa': extremo.get('anio_iaa', ''),
            # ---------------------------------------------------
            'label_ce_principal': label_ce_principal,
            'tabla_bi': tabla_bi, 'tabla_multa': tabla_multa,
            'multa_original_uit': f"{multa_uit:,.3f} UIT",
            'mh_uit': f"{multa_final_del_hecho_uit:,.3f} UIT",
            'bi_uit': f"{bi_uit:,.3f} UIT",
            'aplica_reduccion': aplica_reduccion_str == 'Sí', 'porcentaje_reduccion': porcentaje_str,
            'texto_reduccion': datos_hecho_completos.get('texto_reduccion', ''),
            'memo_num': datos_hecho_completos.get('memo_num', ''), 'memo_fecha': format_date(datos_hecho_completos.get('memo_fecha'), "d 'de' MMMM 'de' yyyy", locale='es') if datos_hecho_completos.get('memo_fecha') else '',
            'escrito_num': datos_hecho_completos.get('escrito_num', ''), 'escrito_fecha': format_date(datos_hecho_completos.get('escrito_fecha'), "d 'de' MMMM 'de' yyyy", locale='es') if datos_hecho_completos.get('escrito_fecha') else '',
            'multa_con_reduccion_uit': f"{multa_con_reduccion_uit:,.3f} UIT", 'se_aplica_tope': se_aplica_tope, 'tope_multa_uit': f"{tope_multa_uit:,.3f} UIT",
            **(fuentes_ce.get('ce1', {}).get('placeholders_dinamicos', {})),
            'fi_mes': fuentes_ce.get('ce1', {}).get('fi_mes', ''),
            'fi_ipc': fuentes_ce.get('ce1', {}).get('fi_ipc', 0),
            'fi_tc': fuentes_ce.get('ce1', {}).get('fi_tc', 0),
            'fuente_cos': res_bi.get('fuente_cos', ''),
            'fecha_incumplimiento_larga': (format_date(fecha_inc, "d 'de' MMMM 'de' yyyy", locale='es') if fecha_inc else "N/A").replace("septiembre", "setiembre").replace("Septiembre", "Setiembre"),
        }
        
        doc_tpl_bi.render(contexto_word, autoescape=True, jinja_env=jinja_env)
        buf_final_hecho = io.BytesIO()
        doc_tpl_bi.save(buf_final_hecho)

        # 6. Generar Anexo CE
        anexos_ce = []
        tabla_ce1_anx = create_table_subdoc(tpl_anexo, ["Descripción", "Cantidad", "Horas", "Precio asociado (S/)", "Factor de ajuste 3/", "Monto (*) (S/)", "Monto (*) (US$) 4/"], ce1_fmt, ['descripcion', 'cantidad', 'horas', 'precio_soles', 'factor_ajuste', 'monto_soles', 'monto_dolares'])
        
        contexto_anx = {
            **contexto_word,
            'extremo': {
                 'tipo': f"Declaración Anual {extremo.get('anio', 'N/A')}",
                 'periodicidad': f"Anual {extremo.get('anio', 'N/A')}",
                 'tipo_incumplimiento': tipo_pres,
                 'fecha_incumplimiento': format_date(fecha_inc, "d/MM/yyyy"),
                 'fecha_extemporanea': "N/A",
            },
            'tabla_ce1_anexo': tabla_ce1_anx,
            'fuente_salario_ce1': fuentes_ce.get('ce1', {}).get('fuente_salario', ''),
            'pdf_salario_ce1': fuentes_ce.get('ce1', {}).get('pdf_salario', ''),
            'sustento_prof_ce1': fuentes_ce.get('ce1', {}).get('sustento_item_profesional', ''),
            'ph_ipc_promedio_salario_ce1': fuentes_ce.get('ce1', {}).get('texto_ipc_costeo_salario', ''), # <--- Añadir aquí también
            # ...,
            'fuente_coti_ce1': fuentes_ce.get('ce1', {}).get('fuente_coti', ''),
        }
        tpl_anexo.render(contexto_anx, autoescape=True, jinja_env=jinja_env)
        buf_anexo_final = io.BytesIO()
        tpl_anexo.save(buf_anexo_final)
        anexos_ce.append(buf_anexo_final)

        # 7. Devolver Resultados (Estructura Corregida para App)
        resultados_app = {
             'extremos': [{
                  'tipo': f"IAA {extremo.get('anio_iaa')} ({tipo_pres})",
                  # Datos para tablas CE1 y CE2
                  'ce1_data': res_ce['ce1_data_raw'], 
                  
                  # Datos para resumen por extremo (si la app lo usa)
                  'ce1_soles': res_ce['ce1_soles'], 
                  'ce1_dolares': res_ce['ce1_dolares'],
                  
                  # Datos para tabla BI
                  'bi_data': res_bi.get('table_rows', []), 
                  'bi_uit': bi_uit,

             }],
             'totales': {
                  # Totales acumulados (iguales al extremo en caso simple)
                  'ce1_total_soles': res_ce['ce1_soles'], 
                  'ce1_total_dolares': res_ce['ce1_dolares'],
                  
                  'beneficio_ilicito_uit': bi_uit,
                  'multa_final_uit': multa_uit, 
                  'bi_data_raw': res_bi.get('table_rows', []), 
                  'multa_data_raw': res_multa.get('multa_data_raw', []),
                  
                  # Flags y Reducción
                  'aplica_reduccion': aplica_reduccion_str,
                  'porcentaje_reduccion': porcentaje_str,
                  'multa_con_reduccion_uit': multa_con_reduccion_uit,
                  'multa_reducida_uit': multa_reducida_uit,
                  'multa_final_aplicada': multa_final_del_hecho_uit
             }
        }
        return {
            'doc_pre_compuesto': buf_final_hecho,
            'resultados_para_app': resultados_app,
            'es_extemporaneo': False,
            'usa_capacitacion': aplicar_capacitacion,
            'anexos_ce_generados': anexos_ce,
            'ids_anexos': list(filter(None, res_ce.get('ids_anexos', set()))),
            'texto_explicacion_prorrateo': ''
        }

    except Exception as e:
        import traceback; traceback.print_exc()
        st.error(f"Error _procesar_simple INF001: {e}")
        return {'error': f"Error _procesar_simple INF001: {e}"}
    

# ---------------------------------------------------------------------
# FUNCIÓN 6: PROCESAR HECHO MÚLTIPLE (Implementar lógica similar)
# ---------------------------------------------------------------------
def _procesar_hecho_multiple(datos_comunes, datos_hecho):
    """
    Procesa INF001 con múltiples años IAA, estructura INF008.
    """
    try:
        jinja_env = Environment(trim_blocks=True, lstrip_blocks=True)
        df_tipificacion, id_infraccion = datos_comunes['df_tipificacion'], 'INF001'
        fila_inf = df_tipificacion[df_tipificacion['ID_Infraccion'] == id_infraccion].iloc[0]
        
        id_tpl_principal = fila_inf.get('ID_Plantilla_BI_Extremo') 
        id_tpl_anx = fila_inf.get('ID_Plantilla_CE_Extremo')
        if not id_tpl_principal or not id_tpl_anx: return {'error': f'Faltan IDs de plantilla.'}
        
        buf_plantilla, buf_anexo = descargar_archivo_drive(id_tpl_principal), descargar_archivo_drive(id_tpl_anx)
        if not buf_plantilla or not buf_anexo: return {'error': f'Fallo descarga plantillas.'}
        tpl_principal = DocxTemplate(buf_plantilla)

        # Inicializar acumuladores
        total_bi_uit = 0.0
        lista_bi_resultados_completos = [] 
        lista_bi_app = []
        anexos_ids = set()
        num_hecho = datos_comunes['numero_hecho_actual']
        anexos_ce = []
        lista_extremos_plantilla_word = []
        lista_ce_data_raw = []
        
        aplicar_capacitacion_general = False

        resultados_app = {'extremos': [], 'totales': {'ce1_total_soles': 0}}

        # Iterar extremos
        for j, extremo in enumerate(datos_hecho['extremos']):
            tipo_pres = extremo.get('tipo_presentacion')
            num_secciones_faltantes = 12 if tipo_pres == "No presentó" else extremo.get('num_secciones_faltantes', 0)
            if num_secciones_faltantes <= 0: continue

            res_ce = _calcular_costo_evitado_inf001(datos_comunes, datos_hecho, extremo)
            if res_ce.get('error'): st.error(f"Error CE Extremo {j+1}: {res_ce['error']}"); continue

            if res_ce.get('ce1_data_raw'): lista_ce_data_raw.extend(res_ce['ce1_data_raw'])

            fecha_inc = extremo.get('fecha_incumplimiento')
            anio_iaa = extremo.get('anio_iaa', '')
            texto_bi = f"{datos_hecho.get('texto_hecho', '')} - Extremo {j+1}"
            
            datos_bi_base = {**datos_comunes, 'ce_soles': res_ce['ce_soles_para_bi'], 'ce_dolares': res_ce['ce_dolares_para_bi'], 'fecha_incumplimiento': fecha_inc, 'texto_del_hecho': texto_bi}
            res_bi_parcial = calcular_beneficio_ilicito(datos_bi_base)
            if not res_bi_parcial or res_bi_parcial.get('error'): st.warning(f"Error BI Extremo {j+1}: {res_bi_parcial.get('error', 'Error')}"); continue

            bi_uit = res_bi_parcial.get('beneficio_ilicito_uit', 0.0); total_bi_uit += bi_uit
            anexos_ids.update(res_ce.get('ids_anexos', set()))
            
            resultados_app['totales']['ce1_total_soles'] += res_ce.get('ce1_soles', 0.0)

            resultados_app['extremos'].append({ 'tipo': f"IAA {anio_iaa} ({tipo_pres})", 'ce1_data': res_ce['ce1_data_raw'], 'ce1_soles': res_ce['ce1_soles'], 'ce_soles_para_bi': res_ce['ce_soles_para_bi'], 'bi_data': res_bi_parcial.get('table_rows', []), 'bi_uit': bi_uit})
            lista_bi_resultados_completos.append(res_bi_parcial)
            lista_bi_app.extend(res_bi_parcial.get('table_rows', []))

            # Anexo CE
            # CORRECCIÓN: Usamos 'buf_anexo' que es la variable real descargada al inicio
            tpl_anx_loop = DocxTemplate(io.BytesIO(buf_anexo.getvalue()))
            
            # --- SOLUCIÓN: Formato idéntico al hecho simple (Números y decimales) ---
            ce1_fmt_anx = []
            for idx, item in enumerate(res_ce['ce1_data_raw'], 1):
                ce1_fmt_anx.append({
                    'descripcion': f"{item.get('descripcion', '')} {idx}/",
                    'cantidad': format_decimal_dinamico(item.get('cantidad', 0)),
                    'horas': format_decimal_dinamico(item.get('horas', 0)),
                    'precio_soles': f"S/ {item.get('precio_soles', 0):,.3f}",
                    'factor_ajuste': f"{item.get('factor_ajuste', 0):,.3f}",
                    'monto_soles': f"S/ {item.get('monto_soles', 0):,.3f}",
                    'monto_dolares': f"US$ {item.get('monto_dolares', 0):,.3f}"
                })
            
            ce1_fmt_anx.append({
                'descripcion': "Total", 
                'cantidad': "", 'horas': "", 'precio_soles': "", 'factor_ajuste': "",
                'monto_soles': f"S/ {res_ce['ce1_soles']:,.3f}", 
                'monto_dolares': f"US$ {res_ce['ce1_dolares']:,.3f}"
            })
            
            tabla_ce1_anx = create_table_subdoc(
                tpl_anx_loop, 
                ["Descripción", "Cantidad", "Horas", "Precio asociado (S/)", "Factor de ajuste 3/", "Monto (*) (S/)", "Monto (*) (US$) 4/"], 
                ce1_fmt_anx, 
                ['descripcion', 'cantidad', 'horas', 'precio_soles', 'factor_ajuste', 'monto_soles', 'monto_dolares']
            )

            fuentes_ce = res_ce.get('fuentes', {})
            contexto_anx = {
                **datos_comunes['context_data'], **(fuentes_ce.get('ce1', {}).get('placeholders_dinamicos', {})),
                'hecho': {'numero_imputado': num_hecho},
                'extremo': {'numeral': j+1, 'tipo': f"IAA {anio_iaa} - {tipo_pres}", 'periodicidad': f"{anio_iaa}", 'fecha_incumplimiento': format_date(fecha_inc, "d/MM/yyyy"), 'fecha_extemporanea': "N/A"},
                'tabla_ce1_anexo': tabla_ce1_anx,
                'fi_mes': fuentes_ce.get('ce1', {}).get('fi_mes', ''), 'fi_ipc': fuentes_ce.get('ce1', {}).get('fi_ipc', 0), 'fi_tc': fuentes_ce.get('ce1', {}).get('fi_tc', 0),
                'fuente_salario_ce1': fuentes_ce.get('ce1', {}).get('fuente_salario', ''), 'pdf_salario_ce1': fuentes_ce.get('ce1', {}).get('pdf_salario', ''), 'sustento_prof_ce1': fuentes_ce.get('ce1', {}).get('sustento_item_profesional', ''), 'fuente_coti_ce1': fuentes_ce.get('ce1', {}).get('fuente_coti', ''),
                
                # --- NUEVOS PLACEHOLDERS PARA LA PLANTILLA CE ---
                'texto_ipc_costeo_salario': fuentes_ce.get('ce1', {}).get('texto_ipc_costeo_salario', ''),
                'ph_ipc_promedio_salario_ce1': fuentes_ce.get('ce1', {}).get('texto_ipc_costeo_salario', ''),
            }
            
            # CORRECCIÓN: Guardamos en un nuevo buffer 'buf_anexo_final' para evitar sobreescribir la plantilla base
            tpl_anx_loop.render(contexto_anx, autoescape=True, jinja_env=jinja_env)
            buf_anexo_final = io.BytesIO()
            tpl_anx_loop.save(buf_anexo_final)
            anexos_ce.append(buf_anexo_final)

            # --- SOLUCIÓN: Compactar Notas BI (Múltiple) ---
            filas_bi_crudas_ext, fn_map_orig_ext, fn_data_ext = res_bi_parcial.get('table_rows', []), res_bi_parcial.get('footnote_mapping', {}), res_bi_parcial.get('footnote_data', {})
            
            letras_usadas_ext = sorted(list({r for f in filas_bi_crudas_ext if f.get('ref') for r in f.get('ref').replace(" ", "").split(",") if r}))
            
            letras_base = "abcdefghijklmnopqrstuvwxyz"
            map_traduccion_ext = {v: letras_base[i] for i, v in enumerate(letras_usadas_ext)}
            nuevo_fn_map_ext = {map_traduccion_ext[v]: fn_map_orig_ext[v] for v in letras_usadas_ext if v in fn_map_orig_ext}

            filas_bi_con_superindice = []
            for fila in filas_bi_crudas_ext:
                nueva_fila = fila.copy()
                ref_orig = nueva_fila.get('ref', '')
                super_final = str(nueva_fila.get('descripcion_superindice', ''))
                if ref_orig:
                    nuevas = [map_traduccion_ext[r] for r in ref_orig.replace(" ", "").split(",") if r in map_traduccion_ext]
                    if nuevas: super_final += f"({', '.join(nuevas)})"
                
                nueva_fila['descripcion_texto'] = str(nueva_fila.get('descripcion_texto', nueva_fila.get('descripcion', '')))
                nueva_fila['descripcion_superindice'] = super_final
                filas_bi_con_superindice.append(nueva_fila)

            fn_list_ext = [f"({l}) {obtener_fuente_formateada(k, fn_data_ext, id_infraccion, False)}" for l, k in sorted(nuevo_fn_map_ext.items())]
            fn_data_dict_ext = {'list': fn_list_ext, 'elaboration': 'Elaboración: Subdirección de Sanción y Gestión de Incentivos (SSAG) - DFAI.', 'style': 'FuenteTabla'}
            tabla_bi_cuerpo = create_main_table_subdoc(tpl_principal, ["Descripción", "Monto"], filas_bi_con_superindice, ['descripcion_texto', 'monto'], footnotes_data=fn_data_dict_ext, column_widths=(5, 1))

            # --- SOLUCIÓN: Textos Dinámicos de Razonabilidad por Extremo ---
            horas_secciones_ext = 0
            for item in res_ce.get('ce1_data_raw', []):
                if 'Profesional' in item.get('descripcion', ''):
                    horas_secciones_ext = item.get('horas', 0)
                    break
                    
            texto_horas_secciones_ext = texto_con_numero(horas_secciones_ext, genero='f')
            sufijo_hora_ext = "hora" if horas_secciones_ext == 1 else "horas"
            ph_horas_secciones_ext = f"{texto_horas_secciones_ext} {sufijo_hora_ext}"
            
            texto_num_secciones_ext = texto_con_numero(num_secciones_faltantes, genero='f')
            sufijo_seccion_ext = "sección" if num_secciones_faltantes == 1 else "secciones"
            ph_secciones_texto_ext = f"{texto_num_secciones_ext} {sufijo_seccion_ext}"

            lista_extremos_plantilla_word.append({
                'loop_index': j + 1, 'numeral': f"{num_hecho}.{j + 1}", 'descripcion': f"Cálculo para IAA {anio_iaa} ({tipo_pres})",
                'label_ce_principal': "CE1",
                'tabla_bi': tabla_bi_cuerpo, 'bi_uit_extremo': f"{bi_uit:,.3f} UIT",
                'num_secciones_faltantes_iaa': num_secciones_faltantes, 'num_secciones_faltantes_texto': ph_secciones_texto_ext,
                'horas_secciones_faltantes_texto': ph_horas_secciones_ext,
                'fecha_incumplimiento_larga': (format_date(fecha_inc, "d 'de' MMMM 'de' yyyy", locale='es') if fecha_inc else "N/A").replace("septiembre", "setiembre").replace("Septiembre", "Setiembre"),
                
                # --- NUEVOS PLACEHOLDERS PARA EL BUCLE EN EL CUERPO ---
                'anio_iaa': anio_iaa,
                'fecha_max_presentacion': format_date(extremo.get('fecha_maxima_presentacion'), "d 'de' MMMM 'de' yyyy", locale='es') if extremo.get('fecha_maxima_presentacion') else "N/A",
                'tipo_presentacion_iaa': tipo_pres,
                'es_no_presento': tipo_pres == "No presentó",
            })

        # 5. Post-Cálculo
        if not lista_bi_resultados_completos: return {'error': 'No se pudo calcular BI.'}
        
        factor_f = datos_hecho.get('factor_f_calculado', 1.0)
        res_multa_final = calcular_multa({**datos_comunes, 'beneficio_ilicito': total_bi_uit, 'factor_f': factor_f})
        multa_final_uit = res_multa_final.get('multa_final_uit', 0.0)

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
        
        # --- INICIO: Generar Texto Desglose BI (Estilo INF007) ---
        lista_desglose = []
        for i, ext_res in enumerate(resultados_app.get('extremos', [])):
            bi_valor = ext_res.get('bi_uit', 0.0)
            tipo_desc = f"extremo n.° {i+1}"
            lista_desglose.append(f"{bi_valor:,.3f} UIT del {tipo_desc}")
        
        texto_desglose_bi = ""
        num_extremos_bi = len(lista_desglose)
        if num_extremos_bi == 1:
            texto_desglose_bi = lista_desglose[0]
        elif num_extremos_bi == 2:
            texto_desglose_bi = " y ".join(lista_desglose)
        elif num_extremos_bi > 2:
            texto_desglose_bi = ", ".join(lista_desglose[:-1]) + ", y " + lista_desglose[-1]
        # --- FIN: Generar Texto Desglose BI ---

        contexto_final = {
            **datos_comunes['context_data'], 'acronyms': datos_comunes['acronym_manager'],
            'hecho': {'numero_imputado': num_hecho, 'descripcion': RichText(datos_hecho.get('texto_hecho', '')), 'lista_extremos': lista_extremos_plantilla_word},
            'numeral_hecho': f"IV.{num_hecho + 1}", 'bi_uit_total': f"{total_bi_uit:,.3f} UIT", 'multa_original_uit': f"{multa_final_uit:,.3f} UIT", 'mh_uit': f"{multa_final_del_hecho_uit:,.3f} UIT",
            'tabla_multa_final': tabla_multa_final_subdoc, 'se_usa_capacitacion': aplicar_capacitacion_general,
            
            # --- NUEVO PLACEHOLDER DE DESGLOSE ---
            'texto_desglose_bi': texto_desglose_bi,
            
            'aplica_reduccion': aplica_reduccion_str == 'Sí', 'porcentaje_reduccion': porcentaje_str, 'texto_reduccion': datos_hecho_completos.get('texto_reduccion', ''), 'memo_num': datos_hecho_completos.get('memo_num', ''), 'memo_fecha': format_date(datos_hecho_completos.get('memo_fecha'), "d 'de' MMMM 'de' yyyy", locale='es') if datos_hecho_completos.get('memo_fecha') else '', 'escrito_num': datos_hecho_completos.get('escrito_num', ''), 'escrito_fecha': format_date(datos_hecho_completos.get('escrito_fecha'), "d 'de' MMMM 'de' yyyy", locale='es') if datos_hecho_completos.get('escrito_fecha') else '',
            'multa_con_reduccion_uit': f"{multa_con_reduccion_uit:,.3f} UIT", 'se_aplica_tope': se_aplica_tope, 'tope_multa_uit': f"{tope_multa_uit:,.3f} UIT",
        }

        tpl_principal.render(contexto_final, autoescape=True, jinja_env=jinja_env); buf_final = io.BytesIO(); tpl_principal.save(buf_final)

        # 7. Preparar datos para App
        resultados_app['totales'] = {
            'beneficio_ilicito_uit': total_bi_uit, 'multa_final_uit': multa_final_uit, 
            'bi_data_raw': lista_bi_app, 'multa_data_raw': res_multa_final.get('multa_data_raw', []),
            'ce1_data_raw': lista_ce_data_raw,
            'aplica_reduccion': aplica_reduccion_str, 'porcentaje_reduccion': porcentaje_str, 'multa_con_reduccion_uit': multa_con_reduccion_uit, 'multa_reducida_uit': multa_reducida_uit, 'multa_final_aplicada': multa_final_del_hecho_uit
        }

        return { 
            'doc_pre_compuesto': buf_final, 
            'resultados_para_app': resultados_app, 
            'es_extemporaneo': False, 
            'usa_capacitacion': aplicar_capacitacion_general, 
            'anexos_ce_generados': anexos_ce, 
            'ids_anexos': list(filter(None, anexos_ids)), 
            'aplica_reduccion': aplica_reduccion_str, 
            'porcentaje_reduccion': porcentaje_str, 
            'multa_reducida_uit': multa_reducida_uit,
            
            # --- ADICIONES PARA ANEXO DE GRADUACIÓN ---
            'contexto_grad': contexto_final,
        }
    except Exception as e:
        import traceback; traceback.print_exc(); return {'error': f"Error _procesar_multiple INF001: {e}"}