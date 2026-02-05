# --- Archivo: infracciones/INF010.py ---

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
                     texto_con_numero, create_footnotes_subdoc,
                     create_personal_table_subdoc, format_decimal_dinamico, redondeo_excel)
from sheets import (calcular_beneficio_ilicito, calcular_multa,
                    descargar_archivo_drive,
                    calcular_beneficio_ilicito_extemporaneo)

try:
    from modulos.calculo_capacitacion import calcular_costo_capacitacion
except ImportError:
    st.error("No se pudo importar 'calcular_costo_capacitacion'.")
    def calcular_costo_capacitacion(*args, **kwargs):
        return {'error': 'Módulo cálculo capacitación no encontrado.'}

# ---------------------------------------------------------------------
# 1. FUNCIÓN DE FECHAS AUTOMÁTICAS (ESTILO INF008)
# ---------------------------------------------------------------------
# ---------------------------------------------------------------------
# 1. FUNCIÓN DE FECHAS AUTOMÁTICAS (FECHA FIJA: 30 DE MARZO)
# ---------------------------------------------------------------------
def _calcular_fechas_inf010(anio, df_dias_no_laborables=None):
    """
    Calcula la fecha máxima (30 de marzo del año siguiente)
    y la fecha de incumplimiento (31 de marzo).
    """
    if not anio: return None, None
    
    anio_siguiente = anio + 1
    
    # Fecha máxima: 30 de marzo
    fecha_maxima = date(anio_siguiente, 3, 30)
    
    # Fecha incumplimiento: 31 de marzo
    fecha_incumplimiento = date(anio_siguiente, 3, 31)
            
    return fecha_maxima, fecha_incumplimiento

# ---------------------------------------------------------------------
# 1. FUNCIÓN DE FECHAS AUTOMÁTICAS (ESTILO INF008)
# ---------------------------------------------------------------------
def _calcular_costo_evitado_inf010_interno(datos_comunes, datos_hecho_general, extremo_data):
    """
    Motor de Cálculo de CE para INF010.
    Corregido: Busca el costo más cercano y filtra por rubro (Lógica INF008).
    """
    result = {'ce1_data_raw': [], 'ce1_soles': 0.0, 'ce1_dolares': 0.0, 
              'ce_soles_para_bi': 0.0, 'ce_dolares_para_bi': 0.0,
              'aplicar_ce2_a_bi': False, 'fuentes': {'ce1': {}}, 'ids_anexos': set(), 'error': None}
    
    try:
        # 1. Carga y validación de datos
        df_items_inf = datos_comunes.get('df_items_infracciones')
        df_costos = datos_comunes.get('df_costos_items')
        df_coti = datos_comunes.get('df_coti_general')
        df_sal = datos_comunes.get('df_salarios_general')
        df_ind = datos_comunes.get('df_indices')
        id_rubro = datos_comunes.get('id_rubro_seleccionado')
        id_inf = 'INF010'
        
        fecha_inc = extremo_data.get('fecha_incumplimiento')
        fecha_final_dt = pd.to_datetime(fecha_inc)

        # 2. Obtener IPC/TC del incumplimiento
        ipc_row_inc = df_ind[df_ind['Indice_Mes'].dt.to_period('M') == fecha_final_dt.to_period('M')]
        if ipc_row_inc.empty:
            result['error'] = f"Sin IPC/TC para {fecha_final_dt.strftime('%B %Y')}"
            return result
        ipc_inc, tc_inc = ipc_row_inc.iloc[0]['IPC_Mensual'], ipc_row_inc.iloc[0]['TC_Mensual']

        # 3. Procesar Receta
        receta = df_items_inf[df_items_inf['ID_Infraccion'] == id_inf]
        if receta.empty:
            result['error'] = f"No hay receta para {id_inf}"
            return result

        items_final = []
        fuentes_ce1 = {'placeholders_dinamicos': {}}
        salario_capturado = False

        for _, item_receta in receta.iterrows():
            if item_receta.get('Tipo_Costo') != 'Remision': continue
            
            id_item = item_receta['ID_Item_Infraccion']
            desc_item = item_receta.get('Nombre_Item', 'N/A')
            
            # --- INICIO LÓGICA DE BÚSQUEDA CORREGIDA (Referencia INF008) ---
            costos_posibles = df_costos[df_costos['ID_Item_Infraccion'] == id_item].copy()
            if costos_posibles.empty: continue
            
            # Filtrado por Rubro (si es Variable)
            tipo_item = item_receta.get('Tipo_Item')
            df_candidatos = pd.DataFrame()
            if tipo_item == 'Variable':
                id_rubro_str = str(id_rubro) if id_rubro else ''
                df_candidatos = costos_posibles[costos_posibles['ID_Rubro'].astype(str).str.contains(fr'\b{id_rubro_str}\b', regex=True, na=False)].copy()
                if df_candidatos.empty:
                    df_candidatos = costos_posibles[costos_posibles['ID_Rubro'].isin(['', 'nan', None])].copy()
            else:
                df_candidatos = costos_posibles.copy()
                
            if df_candidatos.empty: continue

            # Determinar Fecha_Fuente para cada candidato
            fechas_fuente = []
            for _, cand in df_candidatos.iterrows():
                id_gen = cand['ID_General']; f_f = pd.NaT
                if pd.notna(id_gen):
                    if 'SAL' in id_gen:
                        f = df_sal[df_sal['ID_Salario'] == id_gen]
                        if not f.empty: f_f = pd.to_datetime(f"{int(f.iloc[0]['Costeo_Salario'])}-12-31")
                    elif 'COT' in id_gen:
                        f = df_coti[df_coti['ID_Cotizacion'] == id_gen]
                        if not f.empty: f_f = pd.to_datetime(f.iloc[0]['Fecha_Costeo'])
                fechas_fuente.append(f_f)
            
            df_candidatos['Fecha_Fuente'] = fechas_fuente
            df_candidatos.dropna(subset=['Fecha_Fuente'], inplace=True)
            if df_candidatos.empty: continue

            # Seleccionar el más cercano a la fecha de incumplimiento
            df_candidatos['Diferencia_Dias'] = (df_candidatos['Fecha_Fuente'] - fecha_final_dt).dt.days.abs()
            costo_final = df_candidatos.loc[df_candidatos['Diferencia_Dias'].idxmin()]
            # --- FIN LÓGICA DE BÚSQUEDA ---

            # 4. Obtener IPC del Costeo
            id_gen = costo_final['ID_General']; f_f = costo_final['Fecha_Fuente']
            ipc_cost, tc_cost = 0.0, 0.0
            
            if pd.notna(id_gen) and 'SAL' in id_gen:
                idx_anio = df_ind[df_ind['Indice_Mes'].dt.year == f_f.year]
                ipc_cost = float(idx_anio['IPC_Mensual'].mean()) if not idx_anio.empty else 0.0
                # Capturar fuentes de salario
                f_row = df_sal[df_sal['ID_Salario'] == id_gen]
                if not f_row.empty and not salario_capturado:
                    fuentes_ce1['fuente_salario'] = f_row.iloc[0].get('Fuente_Salario','')
                    fuentes_ce1['pdf_salario'] = f_row.iloc[0].get('PDF_Salario','')
                    fuentes_ce1['placeholders_dinamicos']['ref_ipc_salario'] = f"Promedio {f_f.year}, IPC = {ipc_cost}"
                    salario_capturado = True
            elif pd.notna(id_gen) and 'COT' in id_gen:
                idx_row = df_ind[df_ind['Indice_Mes'].dt.to_period('M') == f_f.to_period('M')]
                ipc_cost = float(idx_row.iloc[0]['IPC_Mensual']) if not idx_row.empty else 0.0

            if ipc_cost == 0: continue
            if "Profesional" in desc_item: fuentes_ce1['sustento_item_profesional'] = costo_final.get('Sustento_Item', '')

            # 5. Cálculo de Montos
            precio_s = float(costo_final.get('Costo_Unitario_Item', 0.0))
            factor = redondeo_excel(ipc_inc / ipc_cost, 3)
            horas_fijas = float(item_receta.get('Cantidad_Horas', 0.0))
            
            monto_s = redondeo_excel(1.0 * horas_fijas * precio_s * factor, 3)
            
            items_final.append({
                "descripcion": desc_item, "cantidad": 1.0, "horas": horas_fijas,
                "precio_soles": precio_s, "factor_ajuste": factor,
                "monto_soles": monto_s, "monto_dolares": redondeo_excel(monto_s / tc_inc, 3)
            })
            if costo_final.get('ID_Anexo_Drive'): result['ids_anexos'].add(costo_final.get('ID_Anexo_Drive'))

        # 6. Consolidar Resultados
        result.update({
            'ce1_data_raw': items_final, 'ce1_soles': sum(x['monto_soles'] for x in items_final),
            'ce1_dolares': sum(x['monto_dolares'] for x in items_final),
            'ce_soles_para_bi': sum(x['monto_soles'] for x in items_final),
            'ce_dolares_para_bi': sum(x['monto_dolares'] for x in items_final)
        })
        fuentes_ce1.update({
            'fi_mes': format_date(fecha_final_dt, "MMMM 'de' yyyy", locale='es'),
            'fi_ipc': float(ipc_inc), 'fi_tc': float(tc_inc)
        })
        result['fuentes']['ce1'] = fuentes_ce1
        return result
        
    except Exception as e:
        import traceback; traceback.print_exc()
        result['error'] = f"Error motor CE INF010: {e}"
        return result

# ---------------------------------------------------------------------
# 3. INTERFAZ DE USUARIO (ESTILO INF008 + PERSONAL INTEGRADO)
# ---------------------------------------------------------------------
def renderizar_inputs_especificos(i, df_dias_no_laborables=None):
    st.markdown("##### Detalles del Incumplimiento (INF010)")
    datos_hecho = st.session_state.imputaciones_data[i]

    # Eliminamos inicialización de tabla_personal (ya no se usa CE2)

    if 'extremos' not in datos_hecho: 
        datos_hecho['extremos'] = [{}]
    
    if st.button("➕ Añadir Año", key=f"add_{i}"):
        datos_hecho['extremos'].append({})
        st.rerun()

    for j, extremo in enumerate(datos_hecho['extremos']):
        with st.container(border=True):
            st.markdown(f"**Año de Incumplimiento n.° {j + 1}**")
            
            anio = st.number_input("Año del reporte", min_value=2000, max_value=date.today().year, key=f"anio_{i}_{j}", value=extremo.get('anio', date.today().year-1))
            extremo['anio'] = anio
            
            # Cálculo de fechas fijas
            fecha_max, fecha_inc = _calcular_fechas_inf010(anio)
            extremo['fecha_maxima_presentacion'] = fecha_max
            extremo['fecha_incumplimiento'] = fecha_inc
            
            st.info(f"📅 **Límite:** {fecha_max.strftime('%d/%m/%Y')} | 🚨 **Incumplimiento:** {fecha_inc.strftime('%d/%m/%Y')}")

            # Forzamos el tipo único (No existe extemporáneo para esta infracción)
            extremo['tipo_extremo'] = "No presentó"

    return datos_hecho

# ---------------------------------------------------------------------
# 4. VALIDACIÓN
# ---------------------------------------------------------------------
def validar_inputs(datos_hecho):
    if not datos_hecho.get('extremos'): return False
    for ext in datos_hecho['extremos']:
        if not all([ext.get('anio'), ext.get('tipo_extremo')]): return False
    return True

# ---------------------------------------------------------------------
# 5. PROCESADORES (LOGICA DE MONEDA Y SUPERINDICES)
# ---------------------------------------------------------------------
def _procesar_hecho_simple(datos_comunes, datos_hecho):
    """
    Procesa un hecho INF010 con un único extremo.
    Calcula el CE de remisión (sin capacitación), BI (sin extemporaneidad),
    Multa con reducciones/topes y genera los subdocumentos.
    """
    try:
        jinja_env = Environment(trim_blocks=True, lstrip_blocks=True)
        id_infraccion = 'INF010'
        
        # 1. Cargar plantillas BI y CE
        df_tipificacion = datos_comunes['df_tipificacion']
        filas_inf = df_tipificacion[df_tipificacion['ID_Infraccion'] == id_infraccion]
        if filas_inf.empty: 
            return {'error': f"No se encontró ID '{id_infraccion}' en Tipificación."}
        
        fila_inf = filas_inf.iloc[0]
        id_tpl_bi = fila_inf.get('ID_Plantilla_BI')
        id_tpl_ce = fila_inf.get('ID_Plantilla_CE')
        
        if not id_tpl_bi or not id_tpl_ce: 
            return {'error': f'Faltan IDs de plantilla (BI o CE) para {id_infraccion}.'}
            
        buf_bi = descargar_archivo_drive(id_tpl_bi)
        buf_ce = descargar_archivo_drive(id_tpl_ce)
        if not buf_bi or not buf_ce: 
            return {'error': f'Fallo descarga de plantillas para {id_infraccion}.'}
            
        doc_tpl_bi = DocxTemplate(buf_bi)
        tpl_anexo = DocxTemplate(buf_ce)

        # 2. Calcular Costo Evitado (CE) - Solo Remisión
        extremo = datos_hecho['extremos'][0]
        res_ce = _calcular_costo_evitado_inf010_interno(datos_comunes, datos_hecho, extremo)
        if res_ce.get('error'): 
            return {'error': f"Error CE: {res_ce['error']}"}
        
        # No existe capacitación para INF010 por instrucción del usuario
        aplicar_capacitacion = False
        label_ce_principal = "CE"

        # 3. Calcular BI y Multa
        # Infracción no admite cumplimiento extemporáneo por instrucción
        fecha_inc = extremo.get('fecha_incumplimiento')
        texto_hecho_bi = datos_hecho.get('texto_hecho', 'Hecho no especificado')
        
        datos_bi_base = {
            **datos_comunes, 
            'ce_soles': res_ce['ce_soles_para_bi'], 
            'ce_dolares': res_ce['ce_dolares_para_bi'], 
            'fecha_incumplimiento': fecha_inc, 
            'texto_del_hecho': texto_hecho_bi
        }
        
        res_bi = calcular_beneficio_ilicito(datos_bi_base)
        if not res_bi or res_bi.get('error'): 
            return res_bi or {'error': 'Error al calcular BI.'}
        
        bi_uit = res_bi.get('beneficio_ilicito_uit', 0)

        # --- Lógica de Moneda (Estilo INF004) ---
        moneda_calculo = res_bi.get('moneda_cos', 'USD') 
        es_dolares = (moneda_calculo == 'USD')
        texto_moneda_bi = "moneda extranjera (Dólares)" if es_dolares else "moneda nacional (Soles)"
        ph_bi_abreviatura_moneda = "US$" if es_dolares else "S/"

        # --- Cálculo de Multa y Reducciones ---
        factor_f = datos_hecho.get('factor_f_calculado', 1.0)
        res_multa = calcular_multa({**datos_comunes, 'beneficio_ilicito': bi_uit, 'factor_f': factor_f})
        multa_uit = res_multa.get('multa_final_uit', 0)

        datos_hecho_completos = datos_comunes.get('datos_hecho_completos', {})
        aplica_reduccion_str = datos_hecho_completos.get('aplica_reduccion', 'No')
        porcentaje_str = datos_hecho_completos.get('porcentaje_reduccion', '0%')
        multa_con_reduccion_uit = multa_uit
        
        if aplica_reduccion_str == 'Sí':
            reduccion_factor = 0.5 if porcentaje_str == '50%' else 0.7
            multa_con_reduccion_uit = redondeo_excel(multa_uit * reduccion_factor, 3)

        tope_multa_uit = float(fila_inf.get('Tope_Multa_Infraccion', float('inf')))
        multa_final_del_hecho_uit = min(multa_con_reduccion_uit, tope_multa_uit)
        se_aplica_tope = multa_con_reduccion_uit > tope_multa_uit

        # 4. Generar Tablas Cuerpo
        
        # --- Tabla CE1 (Numerada 1/, 2/...) ---
        ce1_fmt = []
        for idx, item in enumerate(res_ce['ce1_data_raw'], 1):
            ce1_fmt.append({
                'descripcion': f"{item.get('descripcion')} {idx}/",
                'cantidad': format_decimal_dinamico(item.get('cantidad', 0)),
                'horas': format_decimal_dinamico(item.get('horas', 0)),
                'precio_soles': f"S/ {item.get('precio_soles', 0):,.3f}",
                'factor_ajuste': f"{item.get('factor_ajuste', 0):,.3f}",
                'monto_soles': f"S/ {item.get('monto_soles', 0):,.3f}",
                'monto_dolares': f"US$ {item.get('monto_dolares', 0):,.3f}"
            })
        ce1_fmt.append({'descripcion': 'Total', 'monto_soles': f"S/ {res_ce['ce1_soles']:,.3f}", 'monto_dolares': f"US$ {res_ce['ce1_dolares']:,.3f}"})
        
        tabla_ce1_cuerpo = create_table_subdoc(doc_tpl_bi, ["Descripción", "Cantidad", "Horas", "Precio asociado (S/)", "Factor de ajuste 3/", "Monto (*) (S/)", "Monto (*) (US$) 4/"],
                                             ce1_fmt, ['descripcion', 'cantidad', 'horas', 'precio_soles', 'factor_ajuste', 'monto_soles', 'monto_dolares'])

        # --- Tabla BI (Superíndices Alfabéticos Secuenciales) ---
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

        fn_list = [f"({l}) {obtener_fuente_formateada(k, fn_data, id_infraccion, False)}" for l, k in sorted(nuevo_fn_map.items())]
        fn_data_dict = {'list': fn_list, 'elaboration': 'Elaboración: Subdirección de Sanción y Gestión de Incentivos (SSAG) - DFAI.', 'style': 'FuenteTabla'}
        tabla_bi = create_main_table_subdoc(doc_tpl_bi, ["Descripción", "Monto"], filas_bi_para_tabla, keys=['descripcion_texto', 'monto'], footnotes_data=fn_data_dict, column_widths=(5.5, 0.5))

        # --- Tabla Multa ---
        tabla_multa = create_main_table_subdoc(doc_tpl_bi, ["Componentes", "Monto"], res_multa.get('multa_data_raw', []), ['Componentes', 'Monto'], texto_posterior="Elaboración: Subdirección de Sanción y Gestión de Incentivos (SSAG) - DFAI.", estilo_texto_posterior='FuenteTabla', column_widths=(5.5, 0.5))

        # 5. Contexto Final y Renderizado
        fuentes_ce = res_ce.get('fuentes', {})
        aplica_grad = datos_hecho.get('aplica_graduacion') == 'Sí'
        
        contexto_word = {
            **datos_comunes['context_data'],
            **(fuentes_ce.get('ce1', {}).get('placeholders_dinamicos', {})),
            'hecho': {
                'numero_imputado': datos_comunes['numero_hecho_actual'], 
                'descripcion': RichText(datos_hecho.get('texto_hecho', '')),
                'tabla_ce1': tabla_ce1_cuerpo,
                'tabla_bi': tabla_bi,
                'tabla_multa': tabla_multa,
            },
            'numeral_hecho': f"IV.{datos_comunes['numero_hecho_actual'] + 1}",
            'anio_reporte': f"{extremo.get('anio', 'N/A')}",
            
            # Placeholders de fuentes y fechas
            'fuente_salario_ce1': fuentes_ce.get('ce1', {}).get('fuente_salario', ''),
            'pdf_salario_ce1': fuentes_ce.get('ce1', {}).get('pdf_salario', ''),
            'sustento_prof_ce1': fuentes_ce.get('ce1', {}).get('sustento_item_profesional', ''),
            'ref_ipc_salario': fuentes_ce.get('ce1', {}).get('placeholders_dinamicos', {}).get('ref_ipc_salario', ''),
            'fi_mes': fuentes_ce.get('ce1', {}).get('fi_mes', ''),
            'fi_ipc': f"{fuentes_ce.get('ce1', {}).get('fi_ipc', 0):,.3f}",
            'fi_tc': f"{fuentes_ce.get('ce1', {}).get('fi_tc', 0):,.3f}",
            'fecha_incumplimiento_larga': format_date(fecha_inc, "d 'de' MMMM 'de' yyyy", locale='es').replace("septiembre", "setiembre"),
            
            # BI y Multa
            'bi_uit': f"{bi_uit:,.3f} UIT",
            'mh_uit': f"{multa_final_del_hecho_uit:,.3f} UIT",
            'multa_original_uit': f"{multa_uit:,.3f} UIT",
            'multa_con_reduccion_uit': f"{multa_con_reduccion_uit:,.3f} UIT",
            'tope_multa_uit': f"{tope_multa_uit:,.3f} UIT",
            'se_aplica_tope': se_aplica_tope,
            'aplica_reduccion': aplica_reduccion_str == 'Sí',
            'porcentaje_reduccion': porcentaje_str,
            
            # Moneda y Anexos
            'ph_anexo_ce_num': "3" if aplica_grad else "2",
            'bi_moneda_es_dolares': es_dolares,
            'ph_bi_moneda_texto': texto_moneda_bi,
            'ph_bi_moneda_simbolo': ph_bi_abreviatura_moneda,
            'label_ce_principal': label_ce_principal
        }
        
        doc_tpl_bi.render(contexto_word, autoescape=True, jinja_env=jinja_env)
        buf_final_hecho = io.BytesIO(); doc_tpl_bi.save(buf_final_hecho)

        # 6. Generar Anexo CE (Idéntico a cuerpo)
        tabla_ce1_anexo = create_table_subdoc(tpl_anexo, ["Descripción", "Cantidad", "Horas", "Precio asociado (S/)", "Factor de ajuste 3/", "Monto (*) (S/)", "Monto (*) (US$) 4/"],
                                            ce1_fmt, ['descripcion', 'cantidad', 'horas', 'precio_soles', 'factor_ajuste', 'monto_soles', 'monto_dolares'])
        
        contexto_anexo = {
            **contexto_word,
            'extremo': {
                'tipo': f"Incumplimiento Anual {extremo.get('anio', 'N/A')}", 
                'fecha_incumplimiento': format_date(fecha_inc, "d/MM/yyyy")
            },
            'tabla_ce1_anexo': tabla_ce1_anexo,
            'tabla_resumen_anexo': create_table_subdoc(tpl_anexo, ["Componente", "Monto (S/)", "Monto (US$)"], 
                                                     [{'desc': 'Costo de sistematización y remisión - CE', 'sol': f"S/ {res_ce['ce1_soles']:,.3f}", 'dol': f"US$ {res_ce['ce1_dolares']:,.3f}"}], 
                                                     ['desc', 'sol', 'dol'])
        }
        tpl_anexo.render(contexto_anexo, autoescape=True, jinja_env=jinja_env)
        buf_anexo = io.BytesIO(); tpl_anexo.save(buf_anexo)

        # 7. Devolver Resultados
        resultados_app = {
            'totales': {
                'beneficio_ilicito_uit': bi_uit, 
                'multa_final_aplicada': multa_final_del_hecho_uit,
                'ce1_total_soles': res_ce['ce1_soles'],
                'ce1_total_dolares': res_ce['ce1_dolares'],
                'multa_data_raw': res_multa.get('multa_data_raw', [])
            }
        }

        return {
            'doc_pre_compuesto': buf_final_hecho,
            'resultados_para_app': resultados_app,
            'usa_capacitacion': False, 
            'es_extemporaneo': False,
            'anexos_ce_generados': [buf_anexo], 
            'ids_anexos': list(filter(None, res_ce.get('ids_anexos', [])))
        }

    except Exception as e:
        import traceback; traceback.print_exc()
        return {'error': f"Error crítico en _procesar_hecho_simple INF010: {e}"}

def procesar_infraccion(datos_comunes, datos_hecho):
    num = len(datos_hecho.get('extremos', []))
    if num == 1: return _procesar_hecho_simple(datos_comunes, datos_hecho)
    # ...