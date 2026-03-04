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

# --- IMPORTACIONES DE MÓDULOS PROPIOS ---
from textos_manager import obtener_fuente_formateada
from funciones import (create_main_table_subdoc, create_table_subdoc,
                     texto_con_numero, create_footnotes_subdoc, format_decimal_dinamico)
from sheets import (calcular_beneficio_ilicito, calcular_multa,
                    descargar_archivo_drive,
                    calcular_beneficio_ilicito_extemporaneo)
from funciones import create_main_table_subdoc, create_table_subdoc, texto_con_numero, create_footnotes_subdoc, format_decimal_dinamico, redondeo_excel

# ---------------------------------------------------------------------
# FUNCIÓN AUXILIAR DE FECHAS: DECLARACIÓN ANUAL (INF008)
# ---------------------------------------------------------------------

def _calcular_fechas_declaracion(anio, df_dias_no_laborables=None):
    """
    Calcula la fecha máxima (15 días hábiles de Abril del año siguiente)
    y la fecha de incumplimiento (día calendario siguiente).
    """
    if not anio:
        return None, None

    # --- INICIO CORRECCIÓN HÍBRIDA ---
    feriados_pe = holidays.PE()
    dias_no_laborables_set = set()
    if df_dias_no_laborables is not None and 'Fecha_No_Laborable' in df_dias_no_laborables.columns:
        # Convertir la columna a datetime (asumiendo formato DD/MM/YYYY)
        fechas_nl = pd.to_datetime(df_dias_no_laborables['Fecha_No_Laborable'], format='%d/%m/%Y', errors='coerce').dt.date
        dias_no_laborables_set = set(fechas_nl.dropna())
    # --- FIN CORRECCIÓN HÍBRIDA ---
    
    # 1. Determinar el primer día de Abril del AÑO SIGUIENTE
    mes_siguiente = 4
    anio_calculo = anio + 1 # Ej: Declaración 2022 se presenta en Abril 2023

    dia_actual = date(anio_calculo, mes_siguiente, 1)
    dias_habiles_contados = 0

    # 2. Contar 15 días hábiles
    while dias_habiles_contados < 15:
        # --- MODIFICADO ---
        es_habil = dia_actual.weekday() < 5 and dia_actual not in feriados_pe and dia_actual not in dias_no_laborables_set
        # ---
        if es_habil:
            dias_habiles_contados += 1

        if dias_habiles_contados < 15:
            dia_actual += timedelta(days=1)
            
    fecha_maxima_presentacion = dia_actual
    
    # 3. Calcular fecha de incumplimiento (siguiente día calendario)
    fecha_incumplimiento = fecha_maxima_presentacion + timedelta(days=1)
        
    return fecha_maxima_presentacion, fecha_incumplimiento

# ---------------------------------------------------------------------
# FUNCIÓN AUXILIAR: CÁLCULO CE COMPLETO (CE1 + CE2)
# ---------------------------------------------------------------------

def _calcular_costo_evitado_extremo_inf008(datos_comunes, datos_hecho_general, extremo_data):
    """
    Calcula el CE completo (CE1 + CE2 condicional) para un único extremo de INF008.
    """
    result = {
        'ce1_data_raw': [], 'ce1_soles': 0.0, 'ce1_dolares': 0.0,
        'ce_soles_para_bi': 0.0, 'ce_dolares_para_bi': 0.0,
        'ids_anexos': set(),
        'fuentes': {'ce1': {}},
        'error': None
    }
    try:
        # --- 1. Datos del Extremo y Generales ---
        tipo_incumplimiento = extremo_data.get('tipo_extremo')
        # Ahora el flag es automático: Si no presentó, SIEMPRE incluye capacitación
        incluir_capacitacion_flag = (tipo_incumplimiento == "No presentó")
        fecha_incumplimiento_extremo = extremo_data.get('fecha_incumplimiento')

        if not fecha_incumplimiento_extremo:
            raise ValueError("Falta la fecha de incumplimiento del extremo.")

        # --- 2. Unificar Fecha y Fuentes de Incumplimiento ---
        df_ind = datos_comunes.get('df_indices')
        if df_ind is None: raise ValueError("Faltan df_indices en datos_comunes.")
        
        fecha_final_dt = pd.to_datetime(fecha_incumplimiento_extremo, errors='coerce')
        if pd.isna(fecha_final_dt): raise ValueError(f"Fecha de incumplimiento inválida: {fecha_incumplimiento_extremo}")

        ipc_row = df_ind[df_ind['Indice_Mes'].dt.to_period('M') == fecha_final_dt.to_period('M')]
        if ipc_row.empty:
            raise ValueError(f"No se encontró IPC/TC para la fecha de incumplimiento {fecha_final_dt.strftime('%B %Y')}")
            
        ipc_inc, tc_inc = ipc_row.iloc[0]['IPC_Mensual'], ipc_row.iloc[0]['TC_Mensual']
        if pd.isna(ipc_inc) or pd.isna(tc_inc) or tc_inc == 0:
            raise ValueError("Valores IPC/TC inválidos para la fecha de incumplimiento.")

        result['fuentes']['fi_mes'] = format_date(fecha_final_dt, "MMMM 'de' yyyy", locale='es')
        result['fuentes']['fi_ipc'] = float(ipc_inc)
        result['fuentes']['fi_tc'] = float(tc_inc)
        
        # --- 3. Calcular CE1 (Remisión SIGERSOL - Lógica Interna Horas Fijas) ---
        fecha_calculo_ce1 = fecha_final_dt

        # --- INICIO Lógica interna para calcular CE1 ---
        def _calcular_ce1_interno(datos_comunes_ce1, fecha_final):
            res_int = {'items_calculados': [], 'error': None, 'fuentes': {}}
            try:
                df_items_inf = datos_comunes_ce1.get('df_items_infracciones')
                df_costos = datos_comunes_ce1.get('df_costos_items')
                df_coti = datos_comunes_ce1.get('df_coti_general')
                df_sal = datos_comunes_ce1.get('df_salarios_general')
                df_ind_ce1 = datos_comunes_ce1.get('df_indices')
                id_rubro_ce1 = datos_comunes_ce1.get('id_rubro_seleccionado')
                id_inf_ce1 = 'INF008' # <--- CAMBIO
                if any(df is None for df in [df_items_inf, df_costos, df_coti, df_sal, df_ind_ce1]): raise ValueError("Faltan DataFrames CE1.")

                ipc_inc_ce1, tc_inc_ce1 = ipc_inc, tc_inc
                fecha_final_dt_ce1 = fecha_final

                fuentes_ce1 = {'placeholders_dinamicos': {}}; items_ce1 = []; sal_capturado = False
                receta_ce1 = df_items_inf[df_items_inf['ID_Infraccion'] == id_inf_ce1]
                if receta_ce1.empty: raise ValueError(f"No hay receta CE1 para {id_inf_ce1}")

                for _, item_receta in receta_ce1.iterrows():
                    if item_receta.get('Tipo_Costo') != 'Remision': continue
                    
                    # --- INICIO: Búsqueda a prueba de fallos y comas ---
                    id_raw = item_receta.get('ID_Item')
                    if pd.isna(id_raw) or id_raw == '':
                        id_raw = item_receta.get('ID_Item_Infraccion', '')
                    id_items_str = str(id_raw)
                    if id_items_str.lower() == 'nan': id_items_str = ''
                    
                    lista_ids_buscar = [x.strip() for x in id_items_str.split(',') if x.strip()]
                    desc_item = item_receta.get('Nombre_Item', 'N/A')
                    
                    col_id_costos = 'ID_Item' if 'ID_Item' in df_costos.columns else 'ID_Item_Infraccion'
                    costos_posibles = df_costos[df_costos[col_id_costos].isin(lista_ids_buscar)].copy()
                    
                    if costos_posibles.empty: continue
                    # --- FIN ---
                    tipo_item = item_receta.get('Tipo_Item'); df_candidatos = pd.DataFrame()
                    if tipo_item == 'Variable': 
                        id_rubro_str = str(id_rubro_ce1) if id_rubro_ce1 else ''; df_candidatos = costos_posibles[costos_posibles['ID_Rubro'].astype(str).str.contains(fr'\b{id_rubro_str}\b', regex=True, na=False)].copy() if id_rubro_str else pd.DataFrame();
                        if df_candidatos.empty: df_candidatos = costos_posibles[costos_posibles['ID_Rubro'].isin(['', 'nan', None])].copy()
                    elif tipo_item == 'Fijo': df_candidatos = costos_posibles.copy()
                    if df_candidatos.empty: continue
                    fechas_fuente = []
                    for _, cand in df_candidatos.iterrows():
                        id_gen = cand['ID_General']; fecha_f = pd.NaT
                        if pd.notna(id_gen):
                            if 'SAL' in id_gen: f = df_sal[df_sal['ID_Salario'] == id_gen]; fecha_f = pd.to_datetime(f"{int(f.iloc[0]['Costeo_Salario'])}-12-31", errors='coerce') if not f.empty else pd.NaT
                            elif 'COT' in id_gen: f = df_coti[df_coti['ID_Cotizacion'] == id_gen]; fecha_f = pd.to_datetime(f.iloc[0]['Fecha_Costeo'], errors='coerce') if not f.empty else pd.NaT
                        fechas_fuente.append(fecha_f)
                    df_candidatos['Fecha_Fuente'] = fechas_fuente; df_candidatos.dropna(subset=['Fecha_Fuente'], inplace=True)
                    if df_candidatos.empty: continue
                    fecha_naive = fecha_final_dt_ce1.tz_localize(None) if fecha_final_dt_ce1.tzinfo else fecha_final_dt_ce1
                    df_candidatos['Fecha_Fuente_Naive'] = df_candidatos['Fecha_Fuente'].apply(lambda x: x.tz_localize(None) if pd.notna(x) and x.tzinfo else x)
                    df_candidatos['Diferencia_Dias'] = (df_candidatos['Fecha_Fuente_Naive'] - fecha_naive).dt.days.abs()
                    costo_final = df_candidatos.loc[df_candidatos['Diferencia_Dias'].idxmin()]
                    id_gen = costo_final['ID_General']; fecha_f = costo_final['Fecha_Fuente']
                    ipc_cost, tc_cost = 0.0, 0.0
                    
                    # --- PASO 1: Cálculo de IPC y TC diferenciado ---
                    if pd.notna(id_gen) and 'SAL' in id_gen: 
                        idx_anio = df_ind_ce1[df_ind_ce1['Indice_Mes'].dt.year == fecha_f.year]
                        ipc_cost, tc_cost = (float(idx_anio['IPC_Mensual'].mean()), float(idx_anio['TC_Mensual'].mean())) if not idx_anio.empty else (0.0, 0.0)
                    elif pd.notna(id_gen) and 'COT' in id_gen: 
                        idx_row = df_ind_ce1[df_ind_ce1['Indice_Mes'].dt.to_period('M') == fecha_f.to_period('M')]
                        ipc_cost, tc_cost = (float(idx_row.iloc[0]['IPC_Mensual']), float(idx_row.iloc[0]['TC_Mensual'])) if not idx_row.empty else (0.0, 0.0)
                    
                    # Validación crucial: Si no hay IPC, se salta el ítem y no captura textos erróneos
                    if ipc_cost == 0 or pd.isna(ipc_cost): continue
                    
                    # --- PASO 2: Extracción de Fuentes diferenciada (Solo si el IPC fue válido) ---
                    if pd.notna(id_gen):
                        if 'SAL' in id_gen:
                            f_row = df_sal[df_sal['ID_Salario'] == id_gen]
                            if not f_row.empty:
                                fuentes_ce1['texto_ipc_costeo_salario'] = f"Promedio {fecha_f.year}, IPC = {ipc_cost:,.6f}"
                                if not sal_capturado:
                                    fuentes_ce1['fuente_salario'] = f_row.iloc[0].get('Fuente_Salario','')
                                    fuentes_ce1['pdf_salario'] = f_row.iloc[0].get('PDF_Salario','')
                                    sal_capturado = True
                        elif 'COT' in id_gen:
                            f_row = df_coti[df_coti['ID_Cotizacion'] == id_gen]
                            sust = f_row.iloc[0].get('Fuente_Cotizacion') if not f_row.empty else None
                            if sust: fuentes_ce1.setdefault('fuente_coti', []).append(sust)
                            
                    # --- PASO 3: Sustento profesional (Búsqueda a prueba de mayúsculas/minúsculas) ---
                    if 'profesional' in str(desc_item).lower(): 
                        sustento_txt = costo_final.get('Sustento_Item')
                        if pd.notna(sustento_txt) and str(sustento_txt).strip():
                            fuentes_ce1['sustento_item_profesional'] = str(sustento_txt).strip()
                    # --- PASO 4: Placeholders Dinámicos ---
                    try: 
                        key_ph = f"fuente_{desc_item.split()[0].lower().replace(':','')}"
                        fecha_ph = format_date(fecha_f, 'MMMM yyyy', locale='es').lower()
                        texto_ph = f"{desc_item}:\n{fecha_ph}, IPC={ipc_cost:,.3f}"
                        if 'placeholders_dinamicos' not in fuentes_ce1: 
                            fuentes_ce1['placeholders_dinamicos'] = {}
                        fuentes_ce1['placeholders_dinamicos'][key_ph] = texto_ph
                    except Exception: 
                        pass
                    
                    # --- PASO 5: Cálculos Finales del Ítem ---
                    costo_orig = float(costo_final.get('Costo_Unitario_Item', 0.0))
                    moneda = costo_final.get('Moneda_Item')
                    
                    # Validación de Moneda y TC
                    if moneda != 'S/' and (tc_cost == 0 or pd.isna(tc_cost)): 
                        continue
                        
                    precio_s = costo_orig if moneda == 'S/' else costo_orig * tc_cost
                    factor = redondeo_excel(ipc_inc_ce1 / ipc_cost, 3) if ipc_cost > 0 else 0
                    cant = float(item_receta.get('Cantidad_Recursos', 1.0))
                    horas = float(item_receta.get('Cantidad_Horas', 1.0))
                    
                    monto_s = redondeo_excel(cant * horas * precio_s * factor, 3)
                    monto_d = redondeo_excel(monto_s / tc_inc_ce1 if tc_inc_ce1 > 0 else 0, 3)
                    
                    items_ce1.append({
                        "descripcion": desc_item, 
                        "cantidad": cant, 
                        "horas": horas, 
                        "precio_soles": precio_s, 
                        "precio_dolares": round(precio_s / tc_inc_ce1 if tc_inc_ce1 > 0 else 0, 3), 
                        "factor_ajuste": factor, 
                        "monto_soles": monto_s, 
                        "monto_dolares": monto_d, 
                        "id_anexo": costo_final.get('ID_Anexo_Drive')
                    })
                if 'fuente_coti' in fuentes_ce1: fuentes_ce1['fuente_coti'] = "\n".join(filter(None, set(fuentes_ce1['fuente_coti'])))
                res_int['items_calculados'] = items_ce1
                res_int['fuentes'] = fuentes_ce1
            except Exception as e_int: res_int['error'] = f"Error interno CE1: {e_int}"
            return res_int
        # --- FIN Lógica interna CE1 ---

        res_ce1 = _calcular_ce1_interno(datos_comunes, fecha_calculo_ce1)
        if res_ce1.get('error'):
            result['error'] = f"CE1: {res_ce1['error']}"
            return result
        result['ce1_data_raw'] = res_ce1.get('items_calculados', [])
        result['ce1_soles'] = sum(item.get('monto_soles', 0) for item in result['ce1_data_raw'])
        result['ce1_dolares'] = sum(item.get('monto_dolares', 0) for item in result['ce1_data_raw'])
        result['ids_anexos'].update(item.get('id_anexo') for item in result['ce1_data_raw'] if item.get('id_anexo'))
        result['fuentes']['ce1'] = res_ce1.get('fuentes', {})

        # --- 4. Consolidar para BI ---
        result['ce_soles_para_bi'] = result['ce1_soles']
        result['ce_dolares_para_bi'] = result['ce1_dolares']

        if not result['error']: result['error'] = None # Éxito
        return result
    
    except Exception as e:
        import traceback; traceback.print_exc()
        result['error'] = f"Error crítico en _calcular_costo_evitado_extremo_inf008: {e}"
        return result
    
# ---------------------------------------------------------------------
# FUNCIÓN 2: RENDERIZAR INPUTS (Req 4, 5)
# ---------------------------------------------------------------------

def renderizar_inputs_especificos(i, df_dias_no_laborables=None):
    st.markdown("##### Detalles del Incumplimiento: Declaración Anual RRSS (INF008)")
    datos_hecho = st.session_state.imputaciones_data[i] 

    st.markdown("###### **Registro de Incumplimientos (Extremos)**")
    if 'extremos' not in datos_hecho: datos_hecho['extremos'] = [{}]
    if st.button("➕ Añadir Extremo", key=f"add_extremo_{i}"): datos_hecho['extremos'].append({}); st.rerun()

    for j, extremo in enumerate(datos_hecho['extremos']):
        with st.container(border=True):
            st.markdown(f"**Extremo n.° {j + 1}**")

            # --- FILA 1: Año y Fechas (Diseño Limpio) ---
            col_anio, col_vacia = st.columns(2)
            with col_anio:
                anio_actual = date.today().year
                extremo['anio'] = st.number_input("Año de la Declaración", min_value=2000, max_value=anio_actual - 1, step=1, key=f"anio_{i}_{j}", value=extremo.get('anio', anio_actual - 1))
            
            if extremo.get('anio'):
                fecha_max, fecha_inc = _calcular_fechas_declaracion(extremo['anio'], df_dias_no_laborables)
                extremo['fecha_maxima_presentacion'] = fecha_max
                extremo['fecha_incumplimiento'] = fecha_inc
                
                col_fecha_max, col_fecha_inc = st.columns([2, 2])
                with col_fecha_max:
                    st.text_input("Fecha máxima de presentación", value=fecha_max.strftime('%d/%m/%Y'), disabled=True, key=f"fmax_mock_{i}_{j}")
                with col_fecha_inc:
                    st.info(f"**Fecha de incumplimiento:**\n\n{fecha_inc.strftime('%d/%m/%Y')}")

            # --- FILA 2: Tipo de Incumplimiento ---
            tipo_extremo = st.radio(
                "Tipo de incumplimiento", 
                ["No presentó", "Presentó fuera de plazo"], 
                key=f"tipo_extremo_{i}_{j}", 
                index=0 if extremo.get('tipo_extremo') == "No presentó" else 1 if extremo.get('tipo_extremo') == "Presentó fuera de plazo" else None, 
                horizontal=True
            )
            extremo['tipo_extremo'] = tipo_extremo
            
            if tipo_extremo == "Presentó fuera de plazo":
                fecha_inc_actual = extremo.get('fecha_incumplimiento')
                min_fecha_ext = fecha_inc_actual if fecha_inc_actual else date.today()
                extremo['fecha_extemporanea'] = st.date_input("Fecha cumplimiento extemporáneo", min_value=min_fecha_ext, key=f"fecha_ext_{i}_{j}", value=extremo.get('fecha_extemporanea'), format="DD/MM/YYYY")

            if st.button(f"🗑️ Eliminar", key=f"del_extremo_{i}_{j}"): datos_hecho['extremos'].pop(j); st.rerun()
    return datos_hecho

# ---------------------------------------------------------------------
# FUNCIÓN 3: VALIDACIÓN DE INPUTS (Req 4)
# ---------------------------------------------------------------------
def validar_inputs(datos_hecho):
    if not datos_hecho.get('extremos'): return False
    for extremo in datos_hecho.get('extremos', []):
        if not all([extremo.get('anio'), extremo.get('fecha_incumplimiento'), extremo.get('tipo_extremo')]): 
            return False
        if extremo.get('tipo_extremo') == "Presentó fuera de plazo" and not extremo.get('fecha_extemporanea'): 
            return False
    return True

# ---------------------------------------------------------------------
# FUNCIÓN 4: DESPACHADOR PRINCIPAL (Req 2)
# ---------------------------------------------------------------------
def procesar_infraccion(datos_comunes, datos_hecho):
    """
    Decide si procesar como hecho simple (1 extremo) o múltiple (>1 extremo).
    """
    num_extremos = len(datos_hecho.get('extremos', []))
    if num_extremos == 0: return {'error': 'No se ha registrado ningún extremo.'}
    elif num_extremos == 1: return _procesar_hecho_simple(datos_comunes, datos_hecho)
    else: return _procesar_hecho_multiple(datos_comunes, datos_hecho)

# ---------------------------------------------------------------------
# FUNCIÓN 5: PROCESAR HECHO SIMPLE (Req 2, 3)
# ---------------------------------------------------------------------
def _procesar_hecho_simple(datos_comunes, datos_hecho):
    """
    Procesa un hecho INF008 con un único extremo.
    """
    try:
        jinja_env = Environment(trim_blocks=True, lstrip_blocks=True)
        
        # 1. Cargar plantillas BI y CE simples
        df_tipificacion, id_infraccion = datos_comunes['df_tipificacion'], 'INF008'
        filas_inf = df_tipificacion[df_tipificacion['ID_Infraccion'] == id_infraccion]
        if filas_inf.empty: return {'error': f"No se encontró ID '{id_infraccion}' en Tipificación."}
        fila_inf = filas_inf.iloc[0]
        id_tpl_bi, id_tpl_ce = fila_inf.get('ID_Plantilla_BI'), fila_inf.get('ID_Plantilla_CE')
        if not id_tpl_bi or not id_tpl_ce: return {'error': f'Faltan IDs plantilla simple (BI o CE) para {id_infraccion}.'}
        buf_bi, buf_ce = descargar_archivo_drive(id_tpl_bi), descargar_archivo_drive(id_tpl_ce)
        if not buf_bi or not buf_ce: return {'error': f'Fallo descarga plantilla simple para {id_infraccion}.'}
        doc_tpl_bi = DocxTemplate(buf_bi); tpl_anexo = DocxTemplate(buf_ce)

        # 2. Calcular CE (unificado)
        extremo = datos_hecho['extremos'][0]
        res_ce = _calcular_costo_evitado_extremo_inf008(datos_comunes, datos_hecho, extremo)
        if res_ce.get('error'): return {'error': f"Error CE: {res_ce['error']}"}
        label_ce_principal = "CE"

        # 3. Calcular BI y Multa
        tipo_inc, fecha_inc, fecha_ext = extremo.get('tipo_extremo'), extremo.get('fecha_incumplimiento'), extremo.get('fecha_extemporanea')
        es_ext = (tipo_inc == "Presentó fuera de plazo") # Definición anticipada
        
        texto_bi = f"{datos_hecho.get('texto_hecho', 'Hecho no especificado')}"
        datos_bi_base = {**datos_comunes, 'ce_soles': res_ce['ce_soles_para_bi'], 'ce_dolares': res_ce['ce_dolares_para_bi'], 'fecha_incumplimiento': fecha_inc, 'texto_del_hecho': texto_bi}
        res_bi = calcular_beneficio_ilicito_extemporaneo({**datos_bi_base, 'fecha_cumplimiento_extemporaneo': fecha_ext, **calcular_beneficio_ilicito(datos_bi_base)}) if es_ext else calcular_beneficio_ilicito(datos_bi_base)
        
        # --- Lógica de Moneda ---
        moneda_calculo = res_bi.get('moneda_cos', 'USD') 
        es_dolares = (moneda_calculo == 'USD')
        texto_moneda_bi = "moneda extranjera (Dólares)" if es_dolares else "moneda nacional (Soles)"
        ph_bi_abreviatura_moneda = "US$" if es_dolares else "S/"

        if not res_bi or res_bi.get('error'): return res_bi or {'error': 'Error BI.'}
        bi_uit = res_bi.get('beneficio_ilicito_uit', 0)
        res_multa = calcular_multa({**datos_comunes, 'beneficio_ilicito': bi_uit})
        multa_uit = res_multa.get('multa_final_uit', 0)

        # --- Formato BI (con superíndices) ---
        #  --- Reordenamiento de Superíndices ---
        filas_bi_crudas, fn_map_orig, fn_data = res_bi.get('table_rows', []), res_bi.get('footnote_mapping', {}), res_bi.get('footnote_data', {})
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
        tabla_bi = create_main_table_subdoc(doc_tpl_bi, ["Descripción", "Monto"], filas_bi_para_tabla, keys=['descripcion_texto', 'monto'], footnotes_data=fn_data_dict, column_widths=(5.5, 0.5))

        # --- Formato Multa (Con filtro de vacíos) ---
        multa_raw_simple = res_multa.get('multa_data_raw', [])
        multa_limpia_simple = [fila for fila in multa_raw_simple if str(fila.get('Componentes', '')).strip()]
        tabla_multa = create_main_table_subdoc(doc_tpl_bi, ["Componentes", "Monto"], multa_limpia_simple, ['Componentes', 'Monto'], texto_posterior="Elaboración: Subdirección de Sanción y Gestión de Incentivos (SSAG) - DFAI.", estilo_texto_posterior='FuenteTabla', column_widths=(5.5, 0.5))
        # 5. Contexto y Renderizado Cuerpo
        fuentes_ce = res_ce.get('fuentes', {})

        # --- INICIO: (REQ 1) LÓGICA DE REDUCCIÓN Y TOPE ---
        datos_hecho_completos = datos_comunes.get('datos_hecho_completos', {})
        
        # 1. Aplicar Reducción 50%/30% (Reconocimiento)
        aplica_reduccion_str = datos_hecho_completos.get('aplica_reduccion', 'No')
        porcentaje_str = datos_hecho_completos.get('porcentaje_reduccion', '0%')
        multa_con_reduccion_uit = multa_uit # Valor por defecto
        
        if aplica_reduccion_str == 'Sí':
            reduccion_factor = 0.5 if porcentaje_str == '50%' else 0.7
            multa_con_reduccion_uit = redondeo_excel(multa_uit * reduccion_factor, 3)

        # 2. Obtener Tope de Multa
        infraccion_info = datos_comunes['df_tipificacion'][datos_comunes['df_tipificacion']['ID_Infraccion'] == id_infraccion]
        tope_multa_uit = float('inf') # Infinito por defecto (sin tope)
        if not infraccion_info.empty and pd.notna(infraccion_info.iloc[0].get('Tope_Multa_Infraccion')):
            tope_multa_uit = float(infraccion_info.iloc[0]['Tope_Multa_Infraccion'])

        # 3. Aplicar Tope
        multa_final_del_hecho_uit = min(multa_con_reduccion_uit, tope_multa_uit)
        se_aplica_tope = multa_con_reduccion_uit > tope_multa_uit
        # --- FIN: (REQ 1) ---

        # 5. Contexto Final (CORREGIDO)
        fuentes_ce = res_ce.get('fuentes', {})
        contexto_word = {
            **datos_comunes['context_data'],
            **fuentes_ce.get('ce1', {}).get('placeholders_dinamicos', {}),
            'acronyms': datos_comunes['acronym_manager'],
            'hecho': {
                'numero_imputado': datos_comunes['numero_hecho_actual'], 
                'descripcion': RichText(datos_hecho.get('texto_hecho', '')),
                'tabla_bi': tabla_bi,
                'tabla_multa': tabla_multa,
            },
            'numeral_hecho': f"IV.{datos_comunes['numero_hecho_actual'] + 1}",
            'anio_declaracion': f"{extremo.get('anio', 'N/A')}",
            
            'fuente_salario_ce1': fuentes_ce.get('ce1', {}).get('fuente_salario', ''),
            'pdf_salario_ce1': fuentes_ce.get('ce1', {}).get('pdf_salario', ''),
            'sustento_prof_ce1': fuentes_ce.get('ce1', {}).get('sustento_item_profesional', ''),
            'ref_ipc_salario': fuentes_ce.get('ce1', {}).get('placeholders_dinamicos', {}).get('ref_ipc_salario', ''),
            'ph_ipc_promedio_salario_ce1': fuentes_ce.get('ce1', {}).get('texto_ipc_costeo_salario', ''),
            'fecha_incumplimiento_larga': (format_date(fecha_inc, "d 'de' MMMM 'de' yyyy", locale='es') if fecha_inc else "N/A").replace("septiembre", "setiembre").replace("Septiembre", "Setiembre"),
            'fecha_extemporanea_larga': (format_date(fecha_ext, "d 'de' MMMM 'de' yyyy", locale='es') if fecha_ext else "N/A").replace("septiembre", "setiembre").replace("Septiembre", "Setiembre"),
            'label_ce_principal': label_ce_principal,
            'multa_original_uit': f"{multa_uit:,.3f} UIT",
            'mh_uit': f"{multa_final_del_hecho_uit:,.3f} UIT",
            'bi_uit': f"{bi_uit:,.3f} UIT",
            'fuente_cos': res_bi.get('fuente_cos', ''),
            'texto_explicacion_prorrateo': '',

            # --- PLACEHOLDERS DE REDUCCIÓN Y TOPE ---
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
            
            'ph_anexo_ce_num': "3" if datos_hecho.get('aplica_graduacion') == 'Sí' else "2",
            'bi_moneda_es_dolares': es_dolares,
            'ph_bi_moneda_texto': texto_moneda_bi,
            'ph_bi_moneda_simbolo': ph_bi_abreviatura_moneda,
        }

        doc_tpl_bi.render(contexto_word, autoescape=True, jinja_env=jinja_env)
        buf_final_hecho = io.BytesIO()
        doc_tpl_bi.save(buf_final_hecho)

        # 6. Generar Anexo CE (Simple)
        anexos_ce = []
        ce1_fmt = []
        for idx, item in enumerate(res_ce['ce1_data_raw'], 1):
            ce1_fmt.append({
                'descripcion': f"{item.get('descripcion', '')} {idx}/",
                'cantidad': format_decimal_dinamico(item.get('cantidad', 0)),
                'horas': format_decimal_dinamico(item.get('horas', 0)),
                'precio_soles': f"S/ {item.get('precio_soles', 0):,.3f}",
                'factor_ajuste': f"{item.get('factor_ajuste', 0):,.3f}",
                'monto_soles': f"S/ {item.get('monto_soles', 0):,.3f}",
                'monto_dolares': f"US$ {item.get('monto_dolares', 0):,.3f}"
            })
        ce1_fmt.append({
            'descripcion': "Total", 'cantidad': "", 'horas': "", 'precio_soles': "", 'factor_ajuste': "",
            'monto_soles': f"S/ {res_ce['ce1_soles']:,.3f}", 'monto_dolares': f"US$ {res_ce['ce1_dolares']:,.3f}"
        })
        
        tabla_ce1_anx = create_table_subdoc(
            tpl_anexo, 
            ["Descripción", "Cantidad", "Horas", "Precio asociado (S/)", "Factor de ajuste 3/", "Monto (*) (S/)", "Monto (*) (US$) 4/"],
            ce1_fmt, 
            ['descripcion', 'cantidad', 'horas', 'precio_soles', 'factor_ajuste', 'monto_soles', 'monto_dolares']
        )
        
        contexto_anx = {
            **contexto_word,
            'extremo': {
                 'tipo': f"Declaración Anual {extremo.get('anio', 'N/A')}",
                 'periodicidad': f"Anual {extremo.get('anio', 'N/A')}",
                 'tipo_incumplimiento': tipo_inc,
                 'fecha_incumplimiento': format_date(fecha_inc, "d/MM/yyyy"),
                 'fecha_extemporanea': format_date(fecha_ext, "d/MM/yyyy") if fecha_ext else "N/A",
            },
            'tabla_ce1_anexo': tabla_ce1_anx,
            'fuente_salario_ce1': fuentes_ce.get('ce1', {}).get('fuente_salario', ''),
            'pdf_salario_ce1': fuentes_ce.get('ce1', {}).get('pdf_salario', ''),
            'sustento_prof_ce1': fuentes_ce.get('ce1', {}).get('sustento_item_profesional', ''),
            'fuente_coti_ce1': fuentes_ce.get('ce1', {}).get('fuente_coti', ''),
            'fi_mes': fuentes_ce.get('fi_mes', ''),
            'fi_ipc': fuentes_ce.get('fi_ipc', 0),
            'fi_tc': fuentes_ce.get('fi_tc', 0),
            'ph_ipc_promedio_salario_ce1': fuentes_ce.get('ce1', {}).get('texto_ipc_costeo_salario', ''),
        }
        tpl_anexo.render(contexto_anx, autoescape=True, jinja_env=jinja_env)
        buf_anexo_final = io.BytesIO()
        tpl_anexo.save(buf_anexo_final)
        anexos_ce.append(buf_anexo_final)

        # 7. Devolver Resultados
        resultados_app = {
             'extremos': [{
                  'tipo': f"Declaración Anual {extremo.get('anio')} ({tipo_inc})",
                  'ce1_data': res_ce['ce1_data_raw'],
                  'ce1_soles': res_ce['ce1_soles'], 'ce1_dolares': res_ce['ce1_dolares'],
                  'ce_soles_para_bi': res_ce['ce_soles_para_bi'], 'ce_dolares_para_bi': res_ce['ce_dolares_para_bi'],
                  'bi_data': res_bi.get('table_rows', []), 'bi_uit': bi_uit,
             }],
             'totales': {
                  'ce1_total_soles': res_ce['ce1_soles'], 'ce1_total_dolares': res_ce['ce1_dolares'],
                  'ce_total_soles_para_bi': res_ce['ce_soles_para_bi'], 'ce_total_dolares_para_bi': res_ce['ce_dolares_para_bi'],
                  'beneficio_ilicito_uit': bi_uit, 'multa_final_uit': multa_uit, 
                  'bi_data_raw': res_bi.get('table_rows', []), 'multa_data_raw': res_multa.get('multa_data_raw', []),
                  'aplica_reduccion': aplica_reduccion_str, 'porcentaje_reduccion': porcentaje_str,
                  'multa_con_reduccion_uit': multa_con_reduccion_uit, 'multa_final_aplicada': multa_final_del_hecho_uit 
             }
        }
        return {
            'doc_pre_compuesto': buf_final_hecho, 'resultados_para_app': resultados_app,
            'es_extemporaneo': es_ext, 'usa_capacitacion': False,
            'anexos_ce_generados': anexos_ce, 'ids_anexos': list(filter(None, res_ce.get('ids_anexos', set()))),
            'texto_explicacion_prorrateo': ''
        }
    
    except Exception as e:
        import traceback; traceback.print_exc()
        st.error(f"Error _procesar_simple INF008: {e}")
        return {'error': f"Error _procesar_simple INF008: {e}"}

# ---------------------------------------------------------------------
# FUNCIÓN 6: PROCESAR HECHO MÚLTIPLE (Req 2, 3)
# ---------------------------------------------------------------------
def _procesar_hecho_multiple(datos_comunes, datos_hecho):
    """
    Procesa INF008 con múltiples extremos.
    """
    try:
        jinja_env = Environment(trim_blocks=True, lstrip_blocks=True)
        
        # 1. Cargar Plantillas
        df_tipificacion, id_infraccion = datos_comunes['df_tipificacion'], 'INF008'
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
        tpl_principal = DocxTemplate(buffer_plantilla)

        # 2. Inicializar acumuladores
        total_bi_uit = 0.0; lista_bi_resultados_completos = []; anexos_ids = set()
        num_hecho = datos_comunes['numero_hecho_actual']; anexos_ce = []; lista_extremos_plantilla_word = []
        aplicar_capacitacion_general = any(ext.get('tipo_extremo') == 'No presentó' for ext in datos_hecho['extremos'])
        resultados_app = {'extremos': [], 'totales': {'ce1_total_soles': 0, 'ce1_total_dolares': 0, 'ce_total_soles_para_bi': 0, 'ce_total_dolares_para_bi': 0}}

        # 4. Iterar sobre cada extremo
        for j, extremo in enumerate(datos_hecho['extremos']):
            # a. Calcular CE
            res_ce = _calcular_costo_evitado_extremo_inf008(datos_comunes, datos_hecho, extremo)
            if res_ce.get('error'): st.error(f"Error CE Extremo {j+1}: {res_ce['error']}"); continue

            # b. Calcular BI
            tipo_inc, fecha_inc, fecha_ext = extremo.get('tipo_extremo'), extremo.get('fecha_incumplimiento'), extremo.get('fecha_extemporanea')
            texto_bi = f"{datos_hecho.get('texto_hecho', 'Hecho no especificado')} - Extremo {j + 1}"
            datos_bi_base = {**datos_comunes, 'ce_soles': res_ce['ce_soles_para_bi'], 'ce_dolares': res_ce['ce_dolares_para_bi'], 'fecha_incumplimiento': fecha_inc, 'texto_del_hecho': texto_bi}
            res_bi_parcial = calcular_beneficio_ilicito_extemporaneo({**datos_bi_base, 'fecha_cumplimiento_extemporaneo': fecha_ext, **calcular_beneficio_ilicito(datos_bi_base)}) if tipo_inc == "Presentó fuera de plazo" else calcular_beneficio_ilicito(datos_bi_base)
            if not res_bi_parcial or res_bi_parcial.get('error'): st.warning(f"Error BI Extremo {j+1}: {res_bi_parcial.get('error', 'Error')}"); continue

            # c. Acumular totales
            bi_uit = res_bi_parcial.get('beneficio_ilicito_uit', 0.0); total_bi_uit += bi_uit
            anexos_ids.update(res_ce.get('ids_anexos', set()))
            resultados_app['totales']['ce1_total_soles'] += res_ce.get('ce1_soles', 0.0)
            resultados_app['totales']['ce1_total_dolares'] += res_ce.get('ce1_dolares', 0.0)
            resultados_app['totales']['ce_total_soles_para_bi'] += res_ce['ce_soles_para_bi']
            resultados_app['totales']['ce_total_dolares_para_bi'] += res_ce['ce_dolares_para_bi']
            resultados_app['extremos'].append({ 
                'tipo': f"Declaración Anual {extremo.get('anio')} ({tipo_inc})",
                'ce1_data': res_ce['ce1_data_raw'],
                'ce1_soles': res_ce['ce1_soles'], 'ce1_dolares': res_ce['ce1_dolares'],
                'ce_soles_para_bi': res_ce['ce_soles_para_bi'], 'ce_dolares_para_bi': res_ce['ce_dolares_para_bi'],
                'bi_data': res_bi_parcial.get('table_rows', []), 'bi_uit': bi_uit, 
            })
            lista_bi_resultados_completos.append(res_bi_parcial)

            # d. Generar Anexo CE del extremo
            tpl_anx_loop = DocxTemplate(io.BytesIO(buffer_anexo.getvalue()))
            
            # --- SOLUCIÓN: Formato CE1 Anexo Corregido ---
            ce1_fmt_anx = []
            for idx, item in enumerate(res_ce['ce1_data_raw'], 1):
                ce1_fmt_anx.append({
                    'descripcion': f"{item.get('descripcion', '')} {idx}/", # <-- Número después del texto
                    'cantidad': format_decimal_dinamico(item.get('cantidad', 0)),
                    'horas': format_decimal_dinamico(item.get('horas', 0)),
                    'precio_soles': f"S/ {item.get('precio_soles', 0):,.3f}", 
                    'factor_ajuste': f"{item.get('factor_ajuste', 0):,.3f}",
                    'monto_soles': f"S/ {item.get('monto_soles', 0):,.3f}", 
                    'monto_dolares': f"US$ {item.get('monto_dolares', 0):,.3f}"
                })
            
            # Celdas vacías para la fila de total
            ce1_fmt_anx.append({
                'descripcion': 'Total', 'cantidad': "", 'horas': "", 'precio_soles': "", 'factor_ajuste': "", 
                'monto_soles': f"S/ {res_ce['ce1_soles']:,.3f}", 'monto_dolares': f"US$ {res_ce['ce1_dolares']:,.3f}"
            })
            
            # Encabezados con marcadores 3/ y 4/
            tabla_ce1_anx = create_table_subdoc(
                tpl_anx_loop, 
                ["Descripción", "Cantidad", "Horas", "Precio asociado (S/)", "Factor de ajuste 3/", "Monto (*) (S/)", "Monto (*) (US$) 4/"],
                ce1_fmt_anx, 
                ['descripcion', 'cantidad', 'horas', 'precio_soles', 'factor_ajuste', 'monto_soles', 'monto_dolares']
            )

            # --- Contexto Anexo ---
            fuentes_ce = res_ce.get('fuentes', {})
            contexto_anx = {
                **datos_comunes['context_data'],
                **(fuentes_ce.get('ce1', {}).get('placeholders_dinamicos', {})),
                'acronyms': datos_comunes['acronym_manager'],
                'hecho': {'numero_imputado': num_hecho},
                'extremo': {
                    'numeral': j+1,
                    'tipo': f"Declaración Anual {extremo.get('anio', 'N/A')}",
                    'periodicidad': f"Anual {extremo.get('anio', 'N/A')}",
                    'tipo_incumplimiento': tipo_inc,
                    'fecha_incumplimiento': format_date(fecha_inc, "d/MM/yyyy"),
                    'fecha_extemporanea': format_date(fecha_ext, "d/MM/yyyy") if fecha_ext else "N/A",
                },
                'tabla_ce1_anexo': tabla_ce1_anx,
                'fuente_salario_ce1': fuentes_ce.get('ce1', {}).get('fuente_salario', ''),
                'pdf_salario_ce1': fuentes_ce.get('ce1', {}).get('pdf_salario', ''),
                'sustento_prof_ce1': fuentes_ce.get('ce1', {}).get('sustento_item_profesional', ''),
                'fuente_coti_ce1': fuentes_ce.get('ce1', {}).get('fuente_coti', ''),
                'fi_mes': fuentes_ce.get('fi_mes', ''),
                'fi_ipc': fuentes_ce.get('fi_ipc', 0),
                'fi_tc': fuentes_ce.get('fi_tc', 0),
                'ph_ipc_promedio_salario_ce1': fuentes_ce.get('ce1', {}).get('texto_ipc_costeo_salario', ''),
                'ph_anexo_ce_num': "3" if datos_hecho.get('aplica_graduacion') == 'Sí' else "2",
            }
            tpl_anx_loop.render(contexto_anx, autoescape=True, jinja_env=jinja_env); 
            buf_anx = io.BytesIO(); tpl_anx_loop.save(buf_anx); anexos_ce.append(buf_anx)

            # --- Tabla BI (con superíndices y reordenamiento) ---
            filas_bi_crudas_ext, fn_map_orig_ext, fn_data_ext = res_bi_parcial.get('table_rows', []), res_bi_parcial.get('footnote_mapping', {}), res_bi_parcial.get('footnote_data', {})
            es_ext_iter = (tipo_inc == "Presentó fuera de plazo")

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
                nueva_fila['descripcion_texto'] = str(nueva_fila.get('descripcion_texto', ''))
                nueva_fila['descripcion_superindice'] = super_final
                filas_bi_con_superindice.append(nueva_fila)

            fn_list_ext = [f"({l}) {obtener_fuente_formateada(k, fn_data_ext, id_infraccion, es_ext_iter)}" for l, k in sorted(nuevo_fn_map_ext.items())]
            fn_data_dict_ext = {'list': fn_list_ext, 'elaboration': 'Elaboración: Subdirección de Sanción y Gestión de Incentivos (SSAG) - DFAI.', 'style': 'FuenteTabla'}

            tabla_bi_cuerpo = create_main_table_subdoc(tpl_principal, ["Descripción", "Monto"], filas_bi_con_superindice, keys=['descripcion_texto', 'monto'], footnotes_data=fn_data_dict_ext, column_widths=(5.5, 0.5))

            # f. Añadir datos del extremo a la lista para el bucle
            lista_extremos_plantilla_word.append({
                'loop_index': j + 1, 'numeral': f"{num_hecho}.{j + 1}",
                'descripcion': f"Cálculo para el Extremo {j+1}: Declaración Anual {extremo.get('anio', 'N/A')} ({tipo_inc})",
                'fecha_incumplimiento_larga': (format_date(fecha_inc, "d 'de' MMMM 'de' yyyy", locale='es') if fecha_inc else "N/A").replace("septiembre", "setiembre").replace("Septiembre", "Setiembre"),
                'fecha_extemporanea_larga': (format_date(fecha_ext, "d 'de' MMMM 'de' yyyy", locale='es') if fecha_ext else "N/A").replace("septiembre", "setiembre").replace("Septiembre", "Setiembre"),
                'label_ce_principal': "CE",
                'tabla_bi': tabla_bi_cuerpo,
                'bi_uit_extremo': f"{bi_uit:,.3f} UIT",
                'texto_razonabilidad': RichText(""),
                'es_extemporaneo': es_ext_iter,
                'ph_ipc_promedio_salario_ce1': fuentes_ce.get('ce1', {}).get('texto_ipc_costeo_salario', ''),
                # --- SOLUCIÓN: Añadimos las variables faltantes ---
                'anio_declaracion': f"{extremo.get('anio', 'N/A')}",
                'fecha_max_presentacion': format_date(extremo.get('fecha_maxima_presentacion'), "d 'de' MMMM 'de' yyyy", locale='es') if extremo.get('fecha_maxima_presentacion') else "N/A",
            })

        # 5. Post-Cálculo: Multa Final
        if not lista_bi_resultados_completos: return {'error': 'No se pudo calcular BI para ningún extremo.'}
        res_multa_final = calcular_multa({**datos_comunes, 'beneficio_ilicito': total_bi_uit})
        multa_final_uit = res_multa_final.get('multa_final_uit', 0.0)

        multa_final_uit = res_multa_final.get('multa_final_uit', 0.0)

        # --- INICIO: (REQ 1) LÓGICA DE REDUCCIÓN Y TOPE ---
        datos_hecho_completos = datos_comunes.get('datos_hecho_completos', {})
        
        aplica_reduccion_str = datos_hecho_completos.get('aplica_reduccion', 'No')
        porcentaje_str = datos_hecho_completos.get('porcentaje_reduccion', '0%')
        multa_con_reduccion_uit = multa_final_uit # Valor por defecto
        
        if aplica_reduccion_str == 'Sí':
            reduccion_factor = 0.5 if porcentaje_str == '50%' else 0.7
            multa_con_reduccion_uit = redondeo_excel(multa_final_uit * reduccion_factor, 3)

        infraccion_info = datos_comunes['df_tipificacion'][datos_comunes['df_tipificacion']['ID_Infraccion'] == id_infraccion]
        tope_multa_uit = float('inf')
        if not infraccion_info.empty and pd.notna(infraccion_info.iloc[0].get('Tope_Multa_Infraccion')):
            tope_multa_uit = float(infraccion_info.iloc[0]['Tope_Multa_Infraccion'])

        multa_final_del_hecho_uit = min(multa_con_reduccion_uit, tope_multa_uit)
        se_aplica_tope = multa_con_reduccion_uit > tope_multa_uit
        
        # --- ESTA LÍNEA DEFINE LA VARIABLE QUE TE FALTA ---
        multa_reducida_uit = multa_con_reduccion_uit if aplica_reduccion_str == 'Sí' else multa_final_uit
        # --- FIN: (REQ 1) ---
        
        # --- Filtrar tabla multa y crear subdoc ---
        multa_raw_multiple = res_multa_final.get('multa_data_raw', [])
        multa_limpia_multiple = [fila for fila in multa_raw_multiple if str(fila.get('Componentes', '')).strip()]
        tabla_multa_final_subdoc = create_main_table_subdoc( tpl_principal, ["Componentes", "Monto"], multa_limpia_multiple, ['Componentes', 'Monto'], texto_posterior="Elaboración: Subdirección de Sanción y Gestión de Incentivos (SSAG) - DFAI.", estilo_texto_posterior='FuenteTabla', column_widths=(5.5, 0.5) )

        # --- Generar Texto Desglose BI ---
        lista_desglose = []
        for i, ext_res in enumerate(resultados_app.get('extremos', [])):
            bi_valor = ext_res.get('bi_uit', 0.0)
            tipo_desc = f"extremo n.° {i+1}"
            lista_desglose.append(f"{bi_valor:,.3f} UIT del {tipo_desc}")
        texto_desglose_bi = lista_desglose[0] if len(lista_desglose) == 1 else (" y ".join(lista_desglose) if len(lista_desglose) == 2 else ", ".join(lista_desglose[:-1]) + ", y " + lista_desglose[-1])

        # 6. Contexto Final y Renderizado
        contexto_final = {
            **datos_comunes['context_data'], 'acronyms': datos_comunes['acronym_manager'],
            'hecho': {'numero_imputado': num_hecho, 'descripcion': RichText(datos_hecho.get('texto_hecho', '')), 'lista_extremos': lista_extremos_plantilla_word},
            'numeral_hecho': f"IV.{num_hecho + 1}",
            'bi_uit_total': f"{total_bi_uit:,.3f} UIT",
            'texto_desglose_bi': texto_desglose_bi, # Desglose añadido
            'multa_original_uit': f"{multa_final_uit:,.3f} UIT", 
            'mh_uit': f"{multa_final_del_hecho_uit:,.3f} UIT", 
            'tabla_multa_final': tabla_multa_final_subdoc,
            'texto_explicacion_prorrateo': '',
            # --- SOLUCIÓN: Añadir el placeholder del número de anexo ---
            'ph_anexo_ce_num': "3" if datos_hecho.get('aplica_graduacion') == 'Sí' else "2",
            # -----------------------------------------------------------
            'aplica_reduccion': aplica_reduccion_str == 'Sí', 'porcentaje_reduccion': porcentaje_str, 'texto_reduccion': datos_hecho_completos.get('texto_reduccion', ''), 'memo_num': datos_hecho_completos.get('memo_num', ''), 'memo_fecha': format_date(datos_hecho_completos.get('memo_fecha'), "d 'de' MMMM 'de' yyyy", locale='es') if datos_hecho_completos.get('memo_fecha') else '', 'escrito_num': datos_hecho_completos.get('escrito_num', ''), 'escrito_fecha': format_date(datos_hecho_completos.get('escrito_fecha'), "d 'de' MMMM 'de' yyyy", locale='es') if datos_hecho_completos.get('escrito_fecha') else '', 'multa_con_reduccion_uit': f"{multa_con_reduccion_uit:,.3f} UIT", 'se_aplica_tope': se_aplica_tope, 'tope_multa_uit': f"{tope_multa_uit:,.3f} UIT",
        }

        tpl_principal.render(contexto_final, autoescape=True, jinja_env=jinja_env)
        buf_final = io.BytesIO(); tpl_principal.save(buf_final)

        # 7. Preparar datos para App
        resultados_app['totales'] = {
            **resultados_app['totales'], 
            'beneficio_ilicito_uit': total_bi_uit, 
            'multa_data_raw': res_multa_final.get('multa_data_raw', []), 
            'multa_final_uit': multa_final_uit, 
            'bi_data_raw': lista_bi_resultados_completos, 
            'aplica_reduccion': aplica_reduccion_str, 'porcentaje_reduccion': porcentaje_str, 'multa_con_reduccion_uit': multa_con_reduccion_uit, 'multa_reducida_uit': multa_reducida_uit, 'multa_final_aplicada': multa_final_del_hecho_uit 
        }

        return {
            'doc_pre_compuesto': buf_final, 'resultados_para_app': resultados_app,
            'texto_explicacion_prorrateo': '', 'usa_capacitacion': False,
            'es_extemporaneo': any(e.get('tipo_extremo') == 'Presentó fuera de plazo' for e in datos_hecho['extremos']),
            'anexos_ce_generados': anexos_ce, 'ids_anexos': list(filter(None, anexos_ids)),
            'aplica_reduccion': aplica_reduccion_str, 'porcentaje_reduccion': porcentaje_str, 'multa_reducida_uit': multa_reducida_uit
        }
    except Exception as e:
        import traceback; traceback.print_exc(); return {'error': f"Error _procesar_multiple INF008: {e}"}