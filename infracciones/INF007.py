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
                     texto_con_numero, create_footnotes_subdoc,
                     create_personal_table_subdoc, format_decimal_dinamico)
from sheets import (calcular_beneficio_ilicito, calcular_multa,
                    descargar_archivo_drive,
                    calcular_beneficio_ilicito_extemporaneo)
try:
    from modulos.calculo_capacitacion import calcular_costo_capacitacion
except ImportError:
    st.error("No se pudo importar 'calcular_costo_capacitacion'.")
    def calcular_costo_capacitacion(*args, **kwargs):
        return {'error': 'Módulo cálculo capacitación no encontrado.'}
from funciones import create_main_table_subdoc, create_table_subdoc, texto_con_numero, create_footnotes_subdoc, format_decimal_dinamico, redondeo_excel

# ---------------------------------------------------------------------
# FUNCIÓN AUXILIAR DE FECHAS: MANIFIESTOS
# ---------------------------------------------------------------------

def _calcular_fechas_manifiesto(anio, trimestre, df_dias_no_laborables=None):
    """
    Calcula la fecha máxima de presentación (15 días hábiles del mes siguiente)
    y la fecha de incumplimiento.
    """
    if not anio or not trimestre:
        return None, None

    # --- INICIO CORRECCIÓN HÍBRIDA ---
    feriados_pe = holidays.PE()
    dias_no_laborables_set = set()
    if df_dias_no_laborables is not None and 'Fecha_No_Laborable' in df_dias_no_laborables.columns:
        # Convertir la columna a datetime (asumiendo formato DD/MM/YYYY)
        fechas_nl = pd.to_datetime(df_dias_no_laborables['Fecha_No_Laborable'], format='%d/%m/%Y', errors='coerce').dt.date
        dias_no_laborables_set = set(fechas_nl.dropna())
    # --- FIN CORRECCIÓN HÍBRIDA ---
    
    # 1. Determinar el primer día del mes siguiente al trimestre
    if "Trimestre 1" in trimestre: # Ene-Mar -> Siguiente mes es Abril (4)
        mes_siguiente = 4
        anio_calculo = anio
    elif "Trimestre 2" in trimestre: # Abr-Jun -> Siguiente mes es Julio (7)
        mes_siguiente = 7
        anio_calculo = anio
    elif "Trimestre 3" in trimestre: # Jul-Sep -> Siguiente mes es Octubre (10)
        mes_siguiente = 10
        anio_calculo = anio
    elif "Trimestre 4" in trimestre: # Oct-Dic -> Siguiente mes es Enero (1)
        mes_siguiente = 1
        anio_calculo = anio + 1 # del siguiente año
    else:
        return None, None

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
    
    # 3. Calcular fecha de incumplimiento (siguiente día hábil)
    fecha_incumplimiento = fecha_maxima_presentacion + timedelta(days=1)
        
    return fecha_maxima_presentacion, fecha_incumplimiento

# ---------------------------------------------------------------------
# FUNCIÓN AUXILIAR: CÁLCULO CE COMPLETO (CE1 + CE2)
# ---------------------------------------------------------------------

def _calcular_costo_evitado_extremo_inf007(datos_comunes, datos_hecho_general, extremo_data):
    """
    Calcula el CE completo (CE1 + CE2 condicional) para un único extremo de INF007.
    - REFACTORIZADO para unificar fuentes de fecha de incumplimiento.
    """
    result = {
        'ce1_data_raw': [], 'ce1_soles': 0.0, 'ce1_dolares': 0.0,
        'ce2_data_raw': [], 'ce2_soles_calculado': 0.0, 'ce2_dolares_calculado': 0.0,
        'ce_soles_para_bi': 0.0, 'ce_dolares_para_bi': 0.0,
        'aplicar_ce2_a_bi': False,
        'ids_anexos': set(),
        'fuentes': {'ce1': {}, 'ce2': {}}, # Contenedores para fuentes específicas
        'error': None
    }
    try:
        # --- 1. Datos del Extremo y Generales ---
        tipo_incumplimiento = extremo_data.get('tipo_extremo')
        # Ahora el flag es automático: Si no presentó, SIEMPRE incluye capacitación
        incluir_capacitacion_flag = (tipo_incumplimiento == "No presentó")
        fecha_incumplimiento_extremo = extremo_data.get('fecha_incumplimiento')
        num_personal_ce2 = datos_hecho_general.get('num_personal_capacitacion', 0)

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

        # Guardar las fuentes unificadas en el nivel superior de 'fuentes'
        result['fuentes']['fi_mes'] = format_date(fecha_final_dt, "MMMM 'de' yyyy", locale='es')
        result['fuentes']['fi_ipc'] = float(ipc_inc)
        result['fuentes']['fi_tc'] = float(tc_inc)
        
        # --- 3. Calcular CE1 (Remisión SIGERSOL - Lógica Interna Horas Fijas) ---
        fecha_calculo_ce1 = fecha_final_dt # Usar el datetime object

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
                id_inf_ce1 = 'INF007'
                if any(df is None for df in [df_items_inf, df_costos, df_coti, df_sal, df_ind_ce1]): raise ValueError("Faltan DataFrames CE1.")

                # Usar los valores IPC/TC ya calculados
                ipc_inc_ce1, tc_inc_ce1 = ipc_inc, tc_inc
                fecha_final_dt_ce1 = fecha_final

                fuentes_ce1 = {'placeholders_dinamicos': {}}; items_ce1 = []; sal_capturado = False
                receta_ce1 = df_items_inf[df_items_inf['ID_Infraccion'] == id_inf_ce1]
                if receta_ce1.empty: raise ValueError(f"No hay receta CE1 para {id_inf_ce1}")

                for _, item_receta in receta_ce1.iterrows():
                    if item_receta.get('Tipo_Costo') != 'Remision': continue
                    id_item = item_receta['ID_Item_Infraccion']; desc_item = item_receta.get('Nombre_Item', 'N/A')
                    costos_posibles = df_costos[df_costos['ID_Item_Infraccion'] == id_item].copy();
                    if costos_posibles.empty: continue
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
                    if pd.notna(id_gen) and 'SAL' in id_gen: 
                        idx_anio = df_ind_ce1[df_ind_ce1['Indice_Mes'].dt.year == fecha_f.year]; 
                        ipc_cost, tc_cost = (float(idx_anio['IPC_Mensual'].mean()), float(idx_anio['TC_Mensual'].mean())) if not idx_anio.empty else (0.0, 0.0)
                        
                        # --- NUEVO: Capturar Texto de IPC Promedio ---
                        f_row = df_sal[df_sal['ID_Salario'] == id_gen]
                        if not f_row.empty:
                            if 'placeholders_dinamicos' not in fuentes_ce1: 
                                fuentes_ce1['placeholders_dinamicos'] = {}
                            # Genera: "Promedio 2015, IPC = 108.456"
                            fuentes_ce1['placeholders_dinamicos']['ref_ipc_salario'] = f"Promedio {fecha_f.year}, IPC = {ipc_cost:,.6f}"
                    elif pd.notna(id_gen) and 'COT' in id_gen: idx_row = df_ind_ce1[df_ind_ce1['Indice_Mes'].dt.to_period('M') == fecha_f.to_period('M')]; ipc_cost, tc_cost = (float(idx_row.iloc[0]['IPC_Mensual']), float(idx_row.iloc[0]['TC_Mensual'])) if not idx_row.empty else (0.0, 0.0)
                    if ipc_cost == 0 or pd.isna(ipc_cost): continue
                    if pd.notna(id_gen):
                        if 'COT' in id_gen: 
                            f_row=df_coti[df_coti['ID_Cotizacion']==id_gen]; sust=f_row.iloc[0].get('Fuente_Cotizacion') if not f_row.empty else None; fuentes_ce1.setdefault('fuente_coti',[]).append(sust) if sust else None
                        elif 'SAL' in id_gen and not sal_capturado: 
                            f_row=df_sal[df_sal['ID_Salario']==id_gen];
                            if not f_row.empty: 
                                fuentes_ce1['fuente_salario']=f_row.iloc[0].get('Fuente_Salario',''); fuentes_ce1['pdf_salario']=f_row.iloc[0].get('PDF_Salario',''); sal_capturado=True
                    if "Profesional" in desc_item: fuentes_ce1['sustento_item_profesional'] = costo_final.get('Sustento_Item', '')
                    try: 
                        key_ph=f"fuente_{desc_item.split()[0].lower().replace(':','')}"; fecha_ph=format_date(fecha_f, 'MMMM yyyy', locale='es').lower(); texto_ph=f"{desc_item}:\n{fecha_ph}, IPC={ipc_cost:,.3f}"; 
                        if 'placeholders_dinamicos' not in fuentes_ce1: fuentes_ce1['placeholders_dinamicos'] = {}
                        fuentes_ce1['placeholders_dinamicos'][key_ph]=texto_ph
                    except: pass
                    costo_orig=float(costo_final.get('Costo_Unitario_Item', 0.0)); moneda=costo_final.get('Moneda_Item')
                    if moneda!='S/' and (tc_cost==0 or pd.isna(tc_cost)): continue
                    precio_s=costo_orig if moneda=='S/' else costo_orig*tc_cost
                    factor=redondeo_excel(ipc_inc_ce1/ipc_cost, 3) if ipc_cost>0 else 0
                    cant=float(item_receta.get('Cantidad_Recursos', 1.0)); horas=float(item_receta.get('Cantidad_Horas', 1.0))
                    
                    monto_s = redondeo_excel(cant*horas*precio_s*factor, 3)
                    monto_d = redondeo_excel(monto_s/tc_inc_ce1 if tc_inc_ce1 > 0 else 0, 3)
                    
                    items_ce1.append({"descripcion": desc_item, "cantidad": cant, "horas": horas, "precio_soles": precio_s, "precio_dolares": round(precio_s / tc_inc_ce1 if tc_inc_ce1 > 0 else 0, 3), "factor_ajuste": factor, "monto_soles": monto_s, "monto_dolares": monto_d, "id_anexo": costo_final.get('ID_Anexo_Drive')})
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

        # --- 4. Calcular CE2 (Capacitación) ---
        res_ce2 = {}
        fecha_calculo_ce2 = fecha_final_dt # Usar el datetime object

        # Solo calcular CE2 si:
        # 1. El tipo es "No presentó"
        # 2. El flag de incluir capacitación está en "Sí"
        # 3. Hay personal para capacitar
        if tipo_incumplimiento == "No presentó" and incluir_capacitacion_flag and num_personal_ce2 > 0:
            datos_comunes_ce2 = {**datos_comunes, 'fecha_incumplimiento': fecha_calculo_ce2}
            res_ce2 = calcular_costo_capacitacion(num_personal_ce2, datos_comunes_ce2)

            if res_ce2.get('error'):
                result['error'] = f"CE2: {res_ce2['error']}"
                return result
            elif res_ce2:
                # --- INICIO PRORRATEO CE2 (AÑADIR/ACTUALIZAR ESTO) ---
                anio_inc = fecha_calculo_ce2.year
                factor_prorrateo = datos_hecho_general.get('mapa_factores_prorrateo', {}).get(anio_inc, 1.0)
                items_ce2 = res_ce2.get('items_calculados', [])
                if factor_prorrateo < 1.0:
                    for item in items_ce2:
                        item['monto_soles'] = redondeo_excel(item['monto_soles'] * factor_prorrateo, 3)
                        item['monto_dolares'] = redondeo_excel(item['monto_dolares'] * factor_prorrateo, 3)
                        if 'precio_soles' in item:
                            item['precio_soles'] = redondeo_excel(item['precio_soles'] * factor_prorrateo, 3)
                        if 'precio_dolares' in item:
                            item['precio_dolares'] = redondeo_excel(item['precio_dolares'] * factor_prorrateo, 3)
                # --- FIN PRORRATEO CE2 ---
                
                result['ce2_data_raw'] = items_ce2
                result['ce2_soles_calculado'] = sum(item.get('monto_soles', 0) for item in result['ce2_data_raw'])
                result['ce2_dolares_calculado'] = sum(item.get('monto_dolares', 0) for item in result['ce2_data_raw'])
                result['ids_anexos'].update(res_ce2.get('ids_anexos', []))
                result['fuentes']['ce2'] = { 'fuente_salario': res_ce2.get('fuente_salario', ''), 'pdf_salario': res_ce2.get('pdf_salario', ''), 'fuente_coti': res_ce2.get('fuente_coti', '')}

        # --- 5. Aplicar Lógica Condicional para BI ---
        result['aplicar_ce2_a_bi'] = (tipo_incumplimiento == "No presentó" and incluir_capacitacion_flag)
        result['ce_soles_para_bi'] = result['ce1_soles']
        result['ce_dolares_para_bi'] = result['ce1_dolares']
        if result['aplicar_ce2_a_bi']:
            result['ce_soles_para_bi'] += result['ce2_soles_calculado']
            result['ce_dolares_para_bi'] += result['ce2_dolares_calculado']

        if not result['error']: result['error'] = None # Éxito
        return result

    except Exception as e:
        import traceback; traceback.print_exc()
        result['error'] = f"Error crítico en _calcular_costo_evitado_extremo_inf007: {e}"
        return result
    
# ---------------------------------------------------------------------
# FUNCIÓN 2: RENDERIZAR INPUTS (REQ 1, 4, 5)
# ---------------------------------------------------------------------

def renderizar_inputs_especificos(i, df_dias_no_laborables=None):
    st.markdown("##### Detalles del Incumplimiento: Manifiestos SIGERSOL (INF007)")
    datos_hecho = st.session_state.imputaciones_data[i] 

    # --- INICIO CAMBIO 1: Autorelleno y Ubicación de Personal ---
    if 'tabla_personal' not in datos_hecho or not isinstance(datos_hecho['tabla_personal'], list):
        datos_hecho['tabla_personal'] = [
    {
        'Perfil': 'Gerente General', 
        'Descripción': (
            "El Gerente General se encuentra encargado de:\n"
            "- Planear, dirigir y aprobar los objetivos y metas inherentes a las actividades administrativas, operativas y financieras de la empresa.\n"
            "- Administrar los recursos materiales, económicos y tecnológicos de la empresa.\n"
            "- Supervisar, monitorear y evaluar el desarrollo de los procesos y sistemas que se lleva a cabo en la empresa.\n"
            "Planificar, organizar y mantener canales de comunicación que garanticen la aplicación de las disposiciones necesarias para el cumplimiento de los objetivos de la empresa."
        ), 
        'Cantidad': 1
    },
    {
        'Perfil': 'Jefe de Seguridad, Salud Ocupacional y Medio Ambiente', 
        'Descripción': ("El Jefe de Seguridad, Salud Ocupacional y Medio Ambiente se encuentra encargado de:\n"
                        "- Implementar y gestionar el Sistema de Seguridad, Salud Ocupacional y Medio Ambiente y supervisar la correcta ejecución de las políticas, planes y actividades establecidas en el marco de la legislación vigente.\n"
                        "- Elaborar el Plan y Programa Anual de Seguridad, Salud y Medio ambiente en el Trabajo, según la normativa vigente.\n"
                        "- Elaborar el programa anual de entrenamiento y capacitación en temas de seguridad, salud y medio ambiente.\n"
                        "- Diseñar, implementar y liderar la ejecución del plan de auditorías e inspecciones tanto internas como externas en materia de Seguridad, Salud Ocupacional y Medio Ambiente."
                        ), 
        'Cantidad': 1
    }
]

    st.markdown("###### **Extremos del incumplimiento**")
    if 'extremos' not in datos_hecho: datos_hecho['extremos'] = [{}]
    if st.button("➕ Añadir Extremo", key=f"add_extremo_{i}"): datos_hecho['extremos'].append({}); st.rerun()

    for j, extremo in enumerate(datos_hecho['extremos']):
        with st.container(border=True):
            st.markdown(f"**Extremo n.° {j + 1}**")

            col_anio, col_trim, col_tipo = st.columns(3)
            with col_anio:
                anio_actual = date.today().year
                extremo['anio'] = st.number_input("Año", min_value=2000, max_value=anio_actual, step=1, key=f"anio_{i}_{j}", value=extremo.get('anio', anio_actual))
            with col_trim:
                trimestres = ["Trimestre 1 (Ene-Mar)", "Trimestre 2 (Abr-Jun)", "Trimestre 3 (Jul-Sep)", "Trimestre 4 (Oct-Dic)"]
                extremo['trimestre'] = st.selectbox("Trimestre", trimestres, key=f"trimestre_{i}_{j}", index=trimestres.index(extremo.get('trimestre')) if extremo.get('trimestre') in trimestres else None, placeholder="Seleccione...")
            with col_tipo:
                tipo_extremo = st.radio("Tipo de incumplimiento", ["No presentó", "Presentó fuera de plazo"], key=f"tipo_extremo_{i}_{j}", index=0 if extremo.get('tipo_extremo') == "No presentó" else 1 if extremo.get('tipo_extremo') == "Presentó fuera de plazo" else None, horizontal=True)
                extremo['tipo_extremo'] = tipo_extremo

            # Eliminada la opción de considerar capacitación Sí/No (ahora es automático por tipo)

            if extremo.get('anio') and extremo.get('trimestre'):
                fecha_max, fecha_inc = _calcular_fechas_manifiesto(extremo['anio'], extremo['trimestre'], df_dias_no_laborables)
                extremo['fecha_maxima_presentacion'] = fecha_max
                extremo['fecha_incumplimiento'] = fecha_inc
                col_m1, col_m2 = st.columns(2)
                with col_m1: st.metric("Fecha Límite", fecha_max.strftime('%d/%m/%Y'))
                with col_m2: st.metric("Fecha Incumplimiento", fecha_inc.strftime('%d/%m/%Y'))
            
            if tipo_extremo == "Presentó fuera de plazo":
                fecha_inc_actual = extremo.get('fecha_incumplimiento')
                min_fecha_ext = fecha_inc_actual if fecha_inc_actual else date.today()
                extremo['fecha_extemporanea'] = st.date_input("Fecha cumplimiento extemporáneo", min_value=min_fecha_ext, key=f"fecha_ext_{i}_{j}", value=extremo.get('fecha_extemporanea'), format="DD/MM/YYYY")

            # --- TABLA DE PERSONAL INTEGRADA ---
            st.divider()
            st.markdown("###### **Personal a capacitar (CE2)**")
            df_personal = pd.DataFrame(datos_hecho['tabla_personal'])
            if j == 0:
                edited_df = st.data_editor(df_personal, num_rows="dynamic", key=f"data_editor_personal_{i}_{j}", hide_index=True, use_container_width=True, column_config={"Perfil": st.column_config.TextColumn(required=True), "Descripción": st.column_config.TextColumn(width="large"), "Cantidad": st.column_config.NumberColumn(min_value=0, step=1, required=True, format="%d")})
                datos_hecho['tabla_personal'] = edited_df.to_dict('records')
            else:
                st.dataframe(df_personal, use_container_width=True, hide_index=True)

            cant_num = [pd.to_numeric(p.get('Cantidad'), errors='coerce') for p in datos_hecho['tabla_personal']]
            datos_hecho['num_personal_capacitacion'] = int(pd.Series(cant_num).fillna(0).sum())
            if j == 0: st.metric("Total de Personal", f"{datos_hecho['num_personal_capacitacion']}")

            if st.button(f"🗑️ Eliminar", key=f"del_extremo_{i}_{j}"): datos_hecho['extremos'].pop(j); st.rerun()
    return datos_hecho

# ---------------------------------------------------------------------
# FUNCIÓN 3: VALIDACIÓN DE INPUTS (Req 4)
# ---------------------------------------------------------------------
def validar_inputs(datos_hecho):
    """
    Valida inputs de INF007 (Manifiestos).
    """
    # 1. Validar que haya personal (dato global del hecho)
    if not datos_hecho.get('num_personal_capacitacion', 0) > 0: 
        return False

    # 2. Validar que exista al menos un extremo
    if not datos_hecho.get('extremos'): 
        return False

    for extremo in datos_hecho.get('extremos', []):
        # 3. Validar campos obligatorios de cada extremo
        if not all([
            extremo.get('anio'),
            extremo.get('trimestre'),
            extremo.get('fecha_incumplimiento'),
            extremo.get('tipo_extremo')
        ]): 
            return False

        # 4. Validar fecha extemporánea si corresponde
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
    Procesa un hecho INF007 con un único extremo.
    """
    try:
        jinja_env = Environment(trim_blocks=True, lstrip_blocks=True)
        
        # 1. Cargar plantillas BI y CE simples
        df_tipificacion, id_infraccion = datos_comunes['df_tipificacion'], 'INF007'
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
        res_ce = _calcular_costo_evitado_extremo_inf007(datos_comunes, datos_hecho, extremo)
        if res_ce.get('error'): return {'error': f"Error CE: {res_ce['error']}"}
        aplicar_capacitacion = res_ce['aplicar_ce2_a_bi']
        label_ce_principal = "CE1" if aplicar_capacitacion else "CE"

        # 3. Calcular BI y Multa
        # 3. Calcular BI y Multa
        tipo_inc = extremo.get('tipo_extremo')
        fecha_inc = extremo.get('fecha_incumplimiento')
        fecha_ext = extremo.get('fecha_extemporanea')
        
        # --- DEFINICIÓN ANTICIPADA PARA EVITAR EL ERROR ---
        es_ext = (tipo_inc == "Presentó fuera de plazo")
        # --------------------------------------------------

        texto_bi = f"{datos_hecho.get('texto_hecho', 'Hecho no especificado')}"
        datos_bi_base = {**datos_comunes, 'ce_soles': res_ce['ce_soles_para_bi'], 'ce_dolares': res_ce['ce_dolares_para_bi'], 'fecha_incumplimiento': fecha_inc, 'texto_del_hecho': texto_bi}
        res_bi = calcular_beneficio_ilicito_extemporaneo({**datos_bi_base, 'fecha_cumplimiento_extemporaneo': fecha_ext, **calcular_beneficio_ilicito(datos_bi_base)}) if tipo_inc == "Presentó fuera de plazo" else calcular_beneficio_ilicito(datos_bi_base)
        if not res_bi or res_bi.get('error'): return res_bi or {'error': 'Error BI.'}
        bi_uit = res_bi.get('beneficio_ilicito_uit', 0)

        # --- INICIO: Lógica de Moneda (Basado en INF004) ---
        moneda_calculo = res_bi.get('moneda_cos', 'USD') 
        es_dolares = (moneda_calculo == 'USD')
        texto_moneda_bi = "moneda extranjera (Dólares)" if es_dolares else "moneda nacional (Soles)"
        ph_bi_abreviatura_moneda = "US$" if es_dolares else "S/"
        res_multa = calcular_multa({**datos_comunes, 'beneficio_ilicito': bi_uit})
        multa_uit = res_multa.get('multa_final_uit', 0)

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

        # 4. Generar Tablas Cuerpo (Formato INF004)
        
        # --- Formato CE1 ---
        ce1_fmt = []
        for i, item in enumerate(res_ce['ce1_data_raw']):
            desc_orig = item.get('descripcion', '')
            texto_adicional = f"{i+1}/ " # Prefijo numérico simple
            
            ce1_fmt.append({
                'descripcion': f"{desc_orig} {texto_adicional}",
                'cantidad': format_decimal_dinamico(item.get('cantidad', 0)),
                'horas': format_decimal_dinamico(item.get('horas', 0)),
                'precio_soles': f"S/ {item.get('precio_soles', 0):,.3f}",
                'factor_ajuste': f"{item.get('factor_ajuste', 0):,.3f}",
                'monto_soles': f"S/ {item.get('monto_soles', 0):,.3f}",
                'monto_dolares': f"US$ {item.get('monto_dolares', 0):,.3f}"
            })
        
        ce1_fmt.append({
            'descripcion': 'Total',
            'monto_soles': f"S/ {res_ce['ce1_soles']:,.3f}",
            'monto_dolares': f"US$ {res_ce['ce1_dolares']:,.3f}"
        })

        tabla_ce1 = create_table_subdoc(
            doc_tpl_bi, 
            ["Descripción", "Cantidad", "Horas", "Precio asociado (S/)", "Factor de ajuste 3/", "Monto (*) (S/)", "Monto (*) (US$) 4/"],
            ce1_fmt, 
            ['descripcion', 'cantidad', 'horas', 'precio_soles', 'factor_ajuste', 'monto_soles', 'monto_dolares']
        )

        # --- Formato CE2 ---
        tabla_ce2 = None
        ce2_fmt = []
        if aplicar_capacitacion and res_ce['ce2_data_raw']:
            for i, item in enumerate(res_ce['ce2_data_raw']):
                ce2_fmt.append({
                    'descripcion': f"{item.get('descripcion', '')} {i+1}/",
                    'precio_dolares': f"US$ {item.get('precio_dolares', 0):,.3f}",
                    'precio_soles': f"S/ {item.get('precio_soles', 0):,.3f}",
                    'factor_ajuste': f"{item.get('factor_ajuste', 0):,.3f}",
                    'monto_soles': f"S/ {item.get('monto_soles', 0):,.3f}",
                    'monto_dolares': f"US$ {item.get('monto_dolares', 0):,.3f}"
                })
            ce2_fmt.append({'descripcion': 'Total', 'monto_soles': f"S/ {res_ce['ce2_soles_calculado']:,.3f}", 'monto_dolares': f"US$ {res_ce['ce2_dolares_calculado']:,.3f}"})
            tabla_ce2 = create_table_subdoc(doc_tpl_bi, ["Descripción", "Precio (US$)", "Precio (S/)", "Factor de ajuste 2/", "Monto (*) (S/)", "Monto (*) (US$) 3/"],
                                          ce2_fmt, ['descripcion', 'precio_dolares', 'precio_soles', 'factor_ajuste', 'monto_soles', 'monto_dolares'])
        
        # --- SOLUCIÓN: Compactar y Reordenar Notas al Pie de BI ---
        filas_bi_crudas = res_bi.get('table_rows', [])
        fn_map_orig = res_bi.get('footnote_mapping', {})
        fn_data = res_bi.get('footnote_data', {})
        
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
            
            filas_bi_para_tabla.append({
                'descripcion_texto': fila.get('descripcion_texto', ''),
                'descripcion_superindice': super_final,
                'monto': fila.get('monto', '')
            })

        fn_list = [f"({l}) {obtener_fuente_formateada(k, fn_data, id_infraccion, es_ext)}" for l, k in sorted(nuevo_fn_map.items())]
        fn_data_dict = {'list': fn_list, 'elaboration': 'Elaboración: Subdirección de Sanción y Gestión de Incentivos (SSAG) - DFAI.', 'style': 'FuenteTabla'}
        
        tabla_bi = create_main_table_subdoc(doc_tpl_bi, ["Descripción", "Monto"], filas_bi_para_tabla, keys=['descripcion_texto', 'monto'], footnotes_data=fn_data_dict, column_widths=(5, 1))
        # --- Formato Multa ---
        tabla_multa = create_main_table_subdoc(doc_tpl_bi, ["Componentes", "Monto"], res_multa.get('multa_data_raw', []), ['Componentes', 'Monto'], texto_posterior="Elaboración: Subdirección de Sanción y Gestión de Incentivos (SSAG) - DFAI.", estilo_texto_posterior='FuenteTabla', column_widths=(5, 1))

        # --- Formato Tabla Personal ---
        tabla_personal = None
        tabla_pers_data = []
        num_personal_total = datos_hecho.get('num_personal_capacitacion', 0)
        if aplicar_capacitacion:
            tabla_pers_render = datos_hecho.get('tabla_personal', [])
            tabla_pers_sin_total = []
            for fila in tabla_pers_render:
                perfil = fila.get('Perfil'); cantidad = pd.to_numeric(fila.get('Cantidad'), errors='coerce')
                if perfil and cantidad > 0: tabla_pers_sin_total.append({'Perfil': perfil, 'Descripción': fila.get('Descripción', ''), 'Cantidad': int(cantidad)})
            tabla_pers_data = tabla_pers_sin_total + [{'Perfil':'Total', 'Descripción':'', 'Cantidad': int(num_personal_total)}]
            tabla_personal = create_personal_table_subdoc(doc_tpl_bi, ["Perfil (1)", "Descripción", "Cantidad"], tabla_pers_data, ['Perfil', 'Descripción', 'Cantidad'], column_widths=(2,3,1), texto_posterior="Elaboración: Subdirección de Sanción y Gestión de Incentivos (SSAG) - DFAI.")

        # --- INICIO: Formatear Trimestre (Req. Usuario) ---
        trimestre_str_largo = extremo.get('trimestre', 'N/A')
        anio_str = extremo.get('anio', 'N/A')
        
        mapa_trimestre = {
            "Trimestre 1 (Ene-Mar)": "primer trimestre",
            "Trimestre 2 (Abr-Jun)": "segundo trimestre",
            "Trimestre 3 (Jul-Sep)": "tercer trimestre",
            "Trimestre 4 (Oct-Dic)": "cuarto trimestre"
        }
        
        trimestre_formateado = mapa_trimestre.get(trimestre_str_largo, trimestre_str_largo) # Usa el mapeo
        placeholder_final_trimestre = f"{trimestre_formateado} de {anio_str}"
        # --- FIN: Formatear Trimestre ---

        # --- INICIO: Buscar Precios Base de CE2 (Simple) ---
        ce2_base_precio_soles_simple = 0.0
        ce2_base_precio_dolares_simple = 0.0
        if aplicar_capacitacion and res_ce.get('ce2_data_raw'):
            try:
                primer_item_ce2_simple = res_ce['ce2_data_raw'][0]
                ce2_base_precio_soles_simple = primer_item_ce2_simple.get('precio_soles', 0.0)
                ce2_base_precio_dolares_simple = primer_item_ce2_simple.get('precio_dolares', 0.0)
            except IndexError:
                pass # ce2_data_raw existe pero está vacío
        # --- FIN: Buscar Precios Base de CE2 (Simple) ---

        # 5. Contexto y Renderizado Cuerpo
        fuentes_ce = res_ce.get('fuentes', {})
        contexto_word = {
            **datos_comunes['context_data'],
            'ph_anexo_ce_num': "3" if datos_hecho.get('aplica_graduacion') == 'Sí' else "2",
            **fuentes_ce.get('ce1', {}).get('placeholders_dinamicos', {}), # Placeholders dinámicos de CE1
            'acronyms': datos_comunes['acronym_manager'],
            'es_extemporaneo': es_ext,
            # --- INICIO DE LA CORRECCIÓN DE FORMATO ---
            'fecha_max_presentacion': format_date(extremo.get('fecha_maxima_presentacion'), "d 'de' MMMM 'de' yyyy", locale='es') if extremo.get('fecha_maxima_presentacion') else "N/A",
            'fecha_extemporanea': format_date(fecha_ext, "d 'de' MMMM 'de' yyyy", locale='es') if fecha_ext else "N/A",
            # --- FIN DE LA CORRECCIÓN DE FORMATO ---
            'hecho': {'numero_imputado': datos_comunes['numero_hecho_actual'], 'descripcion': RichText(datos_hecho.get('texto_hecho', ''))},
            'numeral_hecho': f"IV.{datos_comunes['numero_hecho_actual'] + 1}", # Suma 1
            
            'trimestre_manifiesto': placeholder_final_trimestre,
            'aplicar_capacitacion': aplicar_capacitacion,
            'label_ce_principal': label_ce_principal,
            # --- INICIO DE LA ADICIÓN (INF007) ---
            'fecha_incumplimiento_larga': (format_date(fecha_inc, "d 'de' MMMM 'de' yyyy", locale='es') if fecha_inc else "N/A").replace("septiembre", "setiembre").replace("Septiembre", "Setiembre"),
            'fecha_extemporanea_larga': (format_date(fecha_ext, "d 'de' MMMM 'de' yyyy", locale='es') if fecha_ext else "N/A").replace("septiembre", "setiembre").replace("Septiembre", "Setiembre"),
            # --- FIN DE LA ADICIÓN ---
            'tabla_ce1': tabla_ce1,
            'tabla_ce2': tabla_ce2,
            'tabla_bi': tabla_bi,
            'tabla_multa': tabla_multa,
            'tabla_detalle_personal': tabla_personal,
            'num_personal_total_texto': texto_con_numero(num_personal_total, 'f'),
            'multa_original_uit': f"{multa_uit:,.3f} UIT", # Multa original
            'mh_uit': f"{multa_final_del_hecho_uit:,.3f} UIT", # Multa final (con reducción/tope)
            'bi_uit': f"{bi_uit:,.3f} UIT",
            # --- AÑADIR ESTAS LÍNEAS ---
            'ce2_monto_soles': f"S/ {res_ce['ce2_soles_calculado']:,.3f}",
            'ce2_monto_dolares': f"US$ {res_ce['ce2_dolares_calculado']:,.3f}",
            # --- FIN ---
            # --- INICIO: ADICIÓN PRECIOS BASE (Simple) ---
            'ce2_precio_base_soles': f"S/ {ce2_base_precio_soles_simple:,.3f}",
            'ce2_precio_base_dolares': f"US$ {ce2_base_precio_dolares_simple:,.3f}",
            # --- FIN: ADICIÓN PRECIOS BASE (Simple) ---
            'fuente_cos': res_bi.get('fuente_cos', ''),
            'texto_explicacion_prorrateo': '', # Se genera en app.py

            # --- INICIO: (REQ 1) PLACEHOLDERS DE REDUCCIÓN Y TOPE ---
            'aplica_reduccion': aplica_reduccion_str == 'Sí',
            'porcentaje_reduccion': porcentaje_str,
            'texto_reduccion': datos_hecho_completos.get('texto_reduccion', ''),
            'memo_num': datos_hecho_completos.get('memo_num', ''),
            'memo_fecha': format_date(datos_hecho_completos.get('memo_fecha'), "d 'de' MMMM 'de' yyyy", locale='es') if datos_hecho_completos.get('memo_fecha') else '',
            'escrito_num': datos_hecho_completos.get('escrito_num', ''),
            'escrito_fecha': format_date(datos_hecho_completos.get('escrito_fecha'), "d 'de' MMMM 'de' yyyy", locale='es') if datos_hecho_completos.get('escrito_fecha') else '',
            'multa_con_reduccion_uit': f"{multa_con_reduccion_uit:,.3f} UIT", # Multa DESPUÉS de 50/30
            'se_aplica_tope': se_aplica_tope, # Booleano para {% if %}
            'tope_multa_uit': f"{tope_multa_uit:,.3f} UIT", # Valor del tope
            # --- FIN: (REQ 1) ---
            'bi_moneda_es_dolares': es_dolares,
            'ph_bi_moneda_texto': texto_moneda_bi,
            'ph_bi_moneda_simbolo': ph_bi_abreviatura_moneda,
        }
        doc_tpl_bi.render(contexto_word, autoescape=True, jinja_env=jinja_env)
        buf_final_hecho = io.BytesIO()
        doc_tpl_bi.save(buf_final_hecho)

        # 6. Generar Anexo CE (Simple)
        anexos_ce = []
        # Reutilizamos ce1_fmt que ya tiene el formato y el total
        tabla_ce1_anx = create_table_subdoc(
            tpl_anexo, 
            ["Descripción", "Cantidad", "Horas", "Precio asociado (S/)", "Factor de ajuste 3/", "Monto (*) (S/)", "Monto (*) (US$) 4/"],
            ce1_fmt, 
            ['descripcion', 'cantidad', 'horas', 'precio_soles', 'factor_ajuste', 'monto_soles', 'monto_dolares']
        )
        tabla_ce2_anx = None
        if aplicar_capacitacion and res_ce['ce2_data_raw']:
             tabla_ce2_anx = create_table_subdoc(tpl_anexo, ["Descripción", "Precio (US$)", "Precio (S/)", "Factor de ajuste 2/", "Monto (*) (S/)", "Monto (*) (US$) 3/"],
                                               ce2_fmt, ['descripcion', 'precio_dolares', 'precio_soles', 'factor_ajuste', 'monto_soles', 'monto_dolares'])
        
        resumen_anexo = [{'desc': 'Costo de sistematización y remisión de información - CE1', 'sol': f"S/ {res_ce['ce1_soles']:,.3f}", 'dol': f"US$ {res_ce['ce1_dolares']:,.3f}"}]
        if aplicar_capacitacion:
            resumen_anexo.append({'desc': 'Costo de Capacitación al personal - CE2', 'sol': f"S/ {res_ce['ce2_soles_calculado']:,.3f}", 'dol': f"US$ {res_ce['ce2_dolares_calculado']:,.3f}"})
        resumen_anexo.append({'desc': 'Costo Evitado Total', 'sol': f"S/ {res_ce['ce_soles_para_bi']:,.3f}", 'dol': f"US$ {res_ce['ce_dolares_para_bi']:,.3f}"})
        tabla_resumen_anx = create_table_subdoc(tpl_anexo, ["Componente", "Monto (*) (S/)", "Monto  (*) (US$)"], resumen_anexo, ['desc', 'sol', 'dol'])
        
        contexto_anx = {
            **contexto_word, # Usar contexto base
            'extremo': {
                 'tipo': f"Manifiesto {extremo.get('trimestre', 'N/A')} {extremo.get('anio', 'N/A')}",
                 'periodicidad': extremo.get('trimestre', ''),
                 'tipo_incumplimiento': tipo_inc,
                # --- INICIO DE LA CORRECCIÓN DE FORMATO ---
                 'fecha_incumplimiento': format_date(fecha_inc, "d 'de' MMMM 'de' yyyy", locale='es'),
                 'fecha_max_presentacion': format_date(extremo.get('fecha_maxima_presentacion'), "d 'de' MMMM 'de' yyyy", locale='es') if extremo.get('fecha_maxima_presentacion') else "N/A",
                 'fecha_extemporanea': format_date(fecha_ext, "d 'de' MMMM 'de' yyyy", locale='es') if fecha_ext else "N/A",
                 # --- FIN DE LA CORRECCIÓN DE FORMATO ---
            },
            'tabla_ce1_anexo': tabla_ce1_anx,
            'tabla_ce2_anexo': tabla_ce2_anx,
            'tabla_resumen_anexo': tabla_resumen_anx,
            'fuente_salario_ce1': fuentes_ce.get('ce1', {}).get('fuente_salario', ''),
            'pdf_salario_ce1': fuentes_ce.get('ce1', {}).get('pdf_salario', ''),
            'sustento_prof_ce1': fuentes_ce.get('ce1', {}).get('sustento_item_profesional', ''),
            'fuente_coti_ce1': fuentes_ce.get('ce1', {}).get('fuente_coti', ''),
            'fuente_salario_ce2': fuentes_ce.get('ce2', {}).get('fuente_salario', ''),
            'pdf_salario_ce2': fuentes_ce.get('ce2', {}).get('pdf_salario', ''),
            'fuente_coti_ce2': fuentes_ce.get('ce2', {}).get('fuente_coti', ''),
            'fi_mes': fuentes_ce.get('fi_mes', ''),
            'fi_ipc': fuentes_ce.get('fi_ipc', 0),
            'fi_tc': fuentes_ce.get('fi_tc', 0),
            # --- AÑADIR ESTAS LÍNEAS ---
            'ce2_monto_soles': f"S/ {res_ce['ce2_soles_calculado']:,.3f}",
            'ce2_monto_dolares': f"US$ {res_ce['ce2_dolares_calculado']:,.3f}",
            # --- FIN ---
            # --- INICIO DE LA ADICIÓN (Anexo Simple) ---
            'ce2_precio_base_soles': f"S/ {ce2_base_precio_soles_simple:,.3f}",
            'ce2_precio_base_dolares': f"US$ {ce2_base_precio_dolares_simple:,.3f}",
            # --- FIN DE LA ADICIÓN ---
        }
        tpl_anexo.render(contexto_anx, autoescape=True, jinja_env=jinja_env)
        buf_anexo_final = io.BytesIO()
        tpl_anexo.save(buf_anexo_final)
        anexos_ce.append(buf_anexo_final)

        # 7. Devolver Resultados
        resultados_app = {
             'extremos': [{
                  'tipo': f"Manifiesto {extremo.get('trimestre')} {extremo.get('anio')} ({tipo_inc})",
                  'ce1_data': res_ce['ce1_data_raw'], 'ce2_data': res_ce['ce2_data_raw'],
                  'ce1_soles': res_ce['ce1_soles'], 'ce1_dolares': res_ce['ce1_dolares'],
                  'ce2_soles_calculado': res_ce['ce2_soles_calculado'], 'ce2_dolares_calculado': res_ce['ce2_dolares_calculado'],
                  'ce_soles_para_bi': res_ce['ce_soles_para_bi'], 'ce_dolares_para_bi': res_ce['ce_dolares_para_bi'],
                  'bi_data': res_bi.get('table_rows', []), 'bi_uit': bi_uit,
                  'aplicar_ce2_a_bi': aplicar_capacitacion
             }],
             'totales': {
                  'ce1_total_soles': res_ce['ce1_soles'], 'ce1_total_dolares': res_ce['ce1_dolares'],
                  'ce2_total_soles_calculado': res_ce['ce2_soles_calculado'], 'ce2_total_dolares_calculado': res_ce['ce2_dolares_calculado'],
                  'ce_total_soles_para_bi': res_ce['ce_soles_para_bi'], 'ce_total_dolares_para_bi': res_ce['ce_dolares_para_bi'],
                  'beneficio_ilicito_uit': bi_uit,
                  'multa_final_uit': multa_uit, # Multa original
                  'bi_data_raw': res_bi.get('table_rows', []),
                  'multa_data_raw': res_multa.get('multa_data_raw', []),
                  'tabla_personal_data': tabla_pers_data,
                  'aplicar_ce2_a_bi': aplicar_capacitacion,
                  
                  # --- INICIO: (REQ 1) DATOS DE REDUCCIÓN PARA APP ---
                  'aplica_reduccion': aplica_reduccion_str,
                  'porcentaje_reduccion': porcentaje_str,
                  'multa_con_reduccion_uit': multa_con_reduccion_uit, # Multa después de 50/30
                  'multa_final_aplicada': multa_final_del_hecho_uit # Multa final CON tope
                  # --- FIN: (REQ 1) ---
             }
        }
        return {
            'doc_pre_compuesto': buf_final_hecho,
            'resultados_para_app': resultados_app,
            'es_extemporaneo': es_ext,
            'usa_capacitacion': aplicar_capacitacion,
            'anexos_ce_generados': anexos_ce,
            'ids_anexos': list(filter(None, res_ce.get('ids_anexos', set()))),
            'texto_explicacion_prorrateo': '',
            'tabla_detalle_personal': tabla_personal,
            'tabla_personal_data': tabla_pers_data
        }
    except Exception as e:
        import traceback; traceback.print_exc()
        st.error(f"Error _procesar_simple INF007: {e}")
        return {'error': f"Error _procesar_simple INF007: {e}"}

# ---------------------------------------------------------------------
# FUNCIÓN 6: PROCESAR HECHO MÚLTIPLE (Req 2, 3)
# ---------------------------------------------------------------------
def _procesar_hecho_multiple(datos_comunes, datos_hecho):
    """
    Procesa INF007 con múltiples extremos, usando la lógica de INF004/INF005.
    """
    try:
        jinja_env = Environment(trim_blocks=True, lstrip_blocks=True)
        
        # 1. Cargar Plantillas
        df_tipificacion, id_infraccion = datos_comunes['df_tipificacion'], 'INF007'
        filas_inf = df_tipificacion[df_tipificacion['ID_Infraccion'] == id_infraccion]
        if filas_inf.empty: return {'error': f"No se encontró ID '{id_infraccion}' en Tipificación."}
        fila_inf = filas_inf.iloc[0]
        id_tpl_principal = fila_inf.get('ID_Plantilla_BI_Extremo') # Plantilla de bucle
        id_tpl_anx = fila_inf.get('ID_Plantilla_CE_Extremo') # Anexo por extremo
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
        resultados_app = {'extremos': [], 'totales': {'ce1_total_soles': 0, 'ce1_total_dolares': 0, 'ce2_total_soles_calculado': 0, 'ce2_total_dolares_calculado': 0, 'ce_total_soles_para_bi': 0, 'ce_total_dolares_para_bi': 0, 'aplicar_ce2_a_bi': aplicar_capacitacion_general}}

        # 3. Generar Tabla Personal (si aplica)
        tabla_pers_subdoc_final = None
        tabla_pers_data = []
        num_pers_total_int = 0
        if aplicar_capacitacion_general:
            tabla_pers_render = datos_hecho.get('tabla_personal', [])
            tabla_pers_sin_total = []
            for fila in tabla_pers_render:
                perfil = fila.get('Perfil'); cantidad = pd.to_numeric(fila.get('Cantidad'), errors='coerce')
                if perfil and cantidad > 0: tabla_pers_sin_total.append({'Perfil': perfil, 'Descripción': fila.get('Descripción', ''), 'Cantidad': int(cantidad)})
            num_pers_total_int = int(datos_hecho.get('num_personal_capacitacion', 0))
            tabla_pers_data = tabla_pers_sin_total + [{'Perfil':'Total', 'Descripción':'', 'Cantidad': num_pers_total_int}]
            tabla_pers_subdoc_final = create_personal_table_subdoc(tpl_principal, ["Perfil (1)", "Descripción", "Cantidad"], tabla_pers_data, ['Perfil', 'Descripción', 'Cantidad'], column_widths=(2,3,1), texto_posterior="Elaboración: Subdirección de Sanción y Gestión de Incentivos (SSAG) - DFAI.")

        # 4. Iterar sobre cada extremo
        for j, extremo in enumerate(datos_hecho['extremos']):
            # a. Calcular CE
            res_ce = _calcular_costo_evitado_extremo_inf007(datos_comunes, datos_hecho, extremo)
            if res_ce.get('error'): st.error(f"Error CE Extremo {j+1}: {res_ce['error']}"); continue
            aplicar_ce2_bi_extremo = res_ce['aplicar_ce2_a_bi']

            # b. Calcular BI
            tipo_inc, fecha_inc, fecha_ext = extremo.get('tipo_extremo'), extremo.get('fecha_incumplimiento'), extremo.get('fecha_extemporanea')
            texto_bi = f"{datos_hecho.get('texto_hecho', 'Hecho no especificado')} - Extremo n.° {j + 1}"
            datos_bi_base = {**datos_comunes, 'ce_soles': res_ce['ce_soles_para_bi'], 'ce_dolares': res_ce['ce_dolares_para_bi'], 'fecha_incumplimiento': fecha_inc, 'texto_del_hecho': texto_bi}
            res_bi_parcial = calcular_beneficio_ilicito_extemporaneo({**datos_bi_base, 'fecha_cumplimiento_extemporaneo': fecha_ext, **calcular_beneficio_ilicito(datos_bi_base)}) if tipo_inc == "Presentó fuera de plazo" else calcular_beneficio_ilicito(datos_bi_base)
            if not res_bi_parcial or res_bi_parcial.get('error'): st.warning(f"Error BI Extremo {j+1}: {res_bi_parcial.get('error', 'Error')}"); continue

            # c. Acumular totales
            bi_uit = res_bi_parcial.get('beneficio_ilicito_uit', 0.0); total_bi_uit += bi_uit
            anexos_ids.update(res_ce.get('ids_anexos', set()))
            resultados_app['totales']['ce1_total_soles'] += res_ce.get('ce1_soles', 0.0)
            resultados_app['totales']['ce1_total_dolares'] += res_ce.get('ce1_dolares', 0.0)
            resultados_app['totales']['ce2_total_soles_calculado'] += res_ce.get('ce2_soles_calculado', 0.0)
            resultados_app['totales']['ce2_total_dolares_calculado'] += res_ce.get('ce2_dolares_calculado', 0.0)
            resultados_app['totales']['ce_total_soles_para_bi'] += res_ce['ce_soles_para_bi']
            resultados_app['totales']['ce_total_dolares_para_bi'] += res_ce['ce_dolares_para_bi']
            resultados_app['extremos'].append({ 
                'tipo': f"Manifiesto {extremo.get('trimestre')} {extremo.get('anio')} ({tipo_inc})",
                'ce1_data': res_ce['ce1_data_raw'], 'ce2_data': res_ce['ce2_data_raw'], 
                'ce1_soles': res_ce['ce1_soles'], 'ce1_dolares': res_ce['ce1_dolares'],
                'ce2_soles_calculado': res_ce['ce2_soles_calculado'], 'ce2_dolares_calculado': res_ce['ce2_dolares_calculado'],
                'ce_soles_para_bi': res_ce['ce_soles_para_bi'], 'ce_dolares_para_bi': res_ce['ce_dolares_para_bi'],
                'bi_data': res_bi_parcial.get('table_rows', []), 'bi_uit': bi_uit, 
                'aplicar_ce2_a_bi': aplicar_ce2_bi_extremo
            })
            lista_bi_resultados_completos.append(res_bi_parcial)

            # d. Generar Anexo CE del extremo (Formato INF004)
            tpl_anx_loop = DocxTemplate(io.BytesIO(buffer_anexo.getvalue()))
            
            # --- Formato CE1 Anexo ---
            ce1_fmt_anx = []
            for i, item in enumerate(res_ce['ce1_data_raw']):
                desc_orig = item.get('descripcion', '')
                texto_adicional = f"{i+1}/ "
                ce1_fmt_anx.append({
                    'descripcion': f"{desc_orig} {texto_adicional}", 'cantidad': format_decimal_dinamico(item.get('cantidad', 0)), 'horas': format_decimal_dinamico(item.get('horas', 0)),
                    'precio_soles': f"S/ {item.get('precio_soles', 0):,.3f}", 'factor_ajuste': f"{item.get('factor_ajuste', 0):,.3f}",
                    'monto_soles': f"S/ {item.get('monto_soles', 0):,.3f}", 'monto_dolares': f"US$ {item.get('monto_dolares', 0):,.3f}"
                })
            ce1_fmt_anx.append({'descripcion': 'Total', 'monto_soles': f"S/ {res_ce['ce1_soles']:,.3f}", 'monto_dolares': f"US$ {res_ce['ce1_dolares']:,.3f}"})
            
            tabla_ce1_anx = create_table_subdoc(
                tpl_anx_loop, 
                ["Descripción", "Cantidad", "Horas", "Precio asociado (S/)", "Factor de ajuste 3/", "Monto (*) (S/)", "Monto (*) (US$) 4/"],
                ce1_fmt_anx, 
                ['descripcion', 'cantidad', 'horas', 'precio_soles', 'factor_ajuste', 'monto_soles', 'monto_dolares']
            )

            # --- Formato CE2 Anexo ---
            tabla_ce2_anx = None
            ce2_fmt_anx = []
            if aplicar_ce2_bi_extremo and res_ce['ce2_data_raw']:
                for i, item in enumerate(res_ce['ce2_data_raw']):
                    ce2_fmt_anx.append({
                        'descripcion': f"{item.get('descripcion', '')} {i+1}/",
                        'precio_dolares': f"US$ {item.get('precio_dolares', 0):,.3f}", 'precio_soles': f"S/ {item.get('precio_soles', 0):,.3f}",
                        'factor_ajuste': f"{item.get('factor_ajuste', 0):,.3f}",
                        'monto_soles': f"S/ {item.get('monto_soles', 0):,.3f}", 'monto_dolares': f"US$ {item.get('monto_dolares', 0):,.3f}"
                    })
                ce2_fmt_anx.append({'descripcion': 'Total', 'monto_soles': f"S/ {res_ce['ce2_soles_calculado']:,.3f}", 'monto_dolares': f"US$ {res_ce['ce2_dolares_calculado']:,.3f}"})
                tabla_ce2_anx = create_table_subdoc(tpl_anx_loop, ["Descripción", "Precio (US$)", "Precio (S/)", "Factor de ajuste 2/", "Monto (*) (S/)", "Monto (*) (US$) 3/"],
                                                  ce2_fmt_anx, ['descripcion', 'precio_dolares', 'precio_soles', 'factor_ajuste', 'monto_soles', 'monto_dolares'])

            # --- Formato Resumen Anexo ---
            resumen_anexo = [{'desc': 'Costo de sistematización y remisión de información - CE1', 'sol': f"S/ {res_ce['ce1_soles']:,.3f}", 'dol': f"US$ {res_ce['ce1_dolares']:,.3f}"}]
            if aplicar_ce2_bi_extremo:
                resumen_anexo.append({'desc': 'Costo de capacitación al personal - CE2', 'sol': f"S/ {res_ce['ce2_soles_calculado']:,.3f}", 'dol': f"US$ {res_ce['ce2_dolares_calculado']:,.3f}"})
            resumen_anexo.append({'desc': 'Costo Evitado Total', 'sol': f"S/ {res_ce['ce_soles_para_bi']:,.3f}", 'dol': f"US$ {res_ce['ce_dolares_para_bi']:,.3f}"})
            tabla_resumen_anx = create_table_subdoc(tpl_anx_loop, ["Componente", "Monto (*) (S/)", "Monto (*) (US$)"], resumen_anexo, ['desc', 'sol', 'dol'])

            # --- Contexto Anexo ---
            fuentes_ce = res_ce.get('fuentes', {})

            # --- INICIO: Buscar Precios Base de CE2 (para Anexo Múltiple) ---
            ce2_base_precio_soles_anexo = 0.0
            ce2_base_precio_dolares_anexo = 0.0
            if aplicar_ce2_bi_extremo and res_ce.get('ce2_data_raw'):
                try:
                    # Toma el precio del primer ítem de capacitación de este extremo
                    primer_item_ce2_anexo = res_ce['ce2_data_raw'][0]
                    ce2_base_precio_soles_anexo = primer_item_ce2_anexo.get('precio_soles', 0.0)
                    ce2_base_precio_dolares_anexo = primer_item_ce2_anexo.get('precio_dolares', 0.0)
                except IndexError:
                    pass # ce2_data_raw existe pero está vacío
            # --- FIN: Buscar Precios Base de CE2 (para Anexo Múltiple) ---

            contexto_anx = {
                **datos_comunes['context_data'],
                **(fuentes_ce.get('ce1', {}).get('placeholders_dinamicos', {})),
                'acronyms': datos_comunes['acronym_manager'],
                'hecho': {'numero_imputado': num_hecho},
                'extremo': {
                    'numeral': j+1,
                    'tipo': f"Manifiesto {extremo.get('trimestre', 'N/A')} {extremo.get('anio', 'N/A')}",
                    'periodicidad': extremo.get('trimestre', ''),
                    'tipo_incumplimiento': tipo_inc,
                    # --- INICIO DE LA CORRECCIÓN DE FORMATO ---
                    'fecha_incumplimiento': format_date(fecha_inc, "d 'de' MMMM 'de' yyyy", locale='es'),
                    'fecha_max_presentacion': format_date(extremo.get('fecha_maxima_presentacion'), "d 'de' MMMM 'de' yyyy", locale='es') if extremo.get('fecha_maxima_presentacion') else "N/A",
                    'fecha_extemporanea': format_date(fecha_ext, "d 'de' MMMM 'de' yyyy", locale='es') if fecha_ext else "N/A",
                    # --- FIN DE LA CORRECCIÓN DE FORMATO ---
                },
                'tabla_ce1_anexo': tabla_ce1_anx,
                'tabla_ce2_anexo': tabla_ce2_anx,
                'tabla_resumen_anexo': tabla_resumen_anx,
                'aplicar_capacitacion': aplicar_ce2_bi_extremo,
                'fuente_salario_ce1': fuentes_ce.get('ce1', {}).get('fuente_salario', ''),
                'pdf_salario_ce1': fuentes_ce.get('ce1', {}).get('pdf_salario', ''),
                'sustento_prof_ce1': fuentes_ce.get('ce1', {}).get('sustento_item_profesional', ''),
                'fuente_coti_ce1': fuentes_ce.get('ce1', {}).get('fuente_coti', ''),
                'fuente_salario_ce2': fuentes_ce.get('ce2', {}).get('fuente_salario', ''),
                'pdf_salario_ce2': fuentes_ce.get('ce2', {}).get('pdf_salario', ''),
                'fuente_coti_ce2': fuentes_ce.get('ce2', {}).get('fuente_coti', ''),
                'fi_mes': fuentes_ce.get('fi_mes', ''),
                'fi_ipc': fuentes_ce.get('fi_ipc', 0),
                'fi_tc': fuentes_ce.get('fi_tc', 0),
                # --- INICIO DE LA ADICIÓN (Anexo Múltiple) ---
                'ce2_precio_base_soles': f"S/ {ce2_base_precio_soles_anexo:,.3f}",
                'ce2_precio_base_dolares': f"US$ {ce2_base_precio_dolares_anexo:,.3f}",
                # --- FIN DE LA ADICIÓN ---
            }
            tpl_anx_loop.render(contexto_anx, autoescape=True, jinja_env=jinja_env); 
            buf_anx = io.BytesIO(); tpl_anx_loop.save(buf_anx); anexos_ce.append(buf_anx)

            # e. Generar tablas CE para el CUERPO
            tabla_ce1_cuerpo = create_table_subdoc(
                tpl_anx_loop, 
                ["Descripción", "Cantidad", "Horas", "Precio asociado (S/)", "Factor de ajuste 3/", "Monto (*) (S/)", "Monto (*) (US$) 4/"],
                ce1_fmt_anx, # Reutiliza ce1_fmt_anx que ya está formateado
                ['descripcion', 'cantidad', 'horas', 'precio_soles', 'factor_ajuste', 'monto_soles', 'monto_dolares']
            )
            tabla_ce2_cuerpo = None
            if aplicar_ce2_bi_extremo and res_ce['ce2_data_raw']: 
                tabla_ce2_cuerpo = create_table_subdoc(tpl_principal, ["Descripción", "Precio (US$)", "Precio (S/)", "Factor de ajuste 2/", "Monto (*) (S/)", "Monto (*) (US$) 3/"],
                                                      ce2_fmt_anx, ['descripcion', 'precio_dolares', 'precio_soles', 'factor_ajuste', 'monto_soles', 'monto_dolares'])
            
            # --- Tabla BI (con superíndices) ---
            filas_bi_crudas_ext, fn_map_ext, fn_data_ext = res_bi_parcial.get('table_rows', []), res_bi_parcial.get('footnote_mapping', {}), res_bi_parcial.get('footnote_data', {})
            es_ext_iter = (tipo_inc == "Presentó fuera de plazo")
            fn_list_ext = [f"({l}) {obtener_fuente_formateada(k, fn_data_ext, id_infraccion, es_ext_iter)}" for l, k in sorted(fn_map_ext.items())]
            fn_data_dict_ext = {'list': fn_list_ext, 'elaboration': 'Elaboración: Subdirección de Sanción y Gestión de Incentivos (SSAG) - DFAI.', 'style': 'FuenteTabla'}
            filas_bi_con_superindice = []
            for fila in filas_bi_crudas_ext:
                nueva_fila = fila.copy(); ref_letra = nueva_fila.get('ref')
                texto_base = str(nueva_fila.get('descripcion_texto', '')); super_existente = str(nueva_fila.get('descripcion_superindice', ''))
                if ref_letra: super_existente += f"({ref_letra})"
                nueva_fila['descripcion_texto'] = texto_base; nueva_fila['descripcion_superindice'] = super_existente
                filas_bi_con_superindice.append(nueva_fila)
            tabla_bi_cuerpo = create_main_table_subdoc(tpl_principal, ["Descripción", "Monto"], filas_bi_con_superindice,
                                                     keys=['descripcion_texto', 'monto'], footnotes_data=fn_data_dict_ext, column_widths=(5, 1))

            # --- INICIO: Formatear Trimestre (Req. 1) ---
            trimestre_str_largo = extremo.get('trimestre', 'N/A')
            anio_str = extremo.get('anio', 'N/A')
            
            mapa_trimestre = {
                "Trimestre 1 (Ene-Mar)": "primer trimestre",
                "Trimestre 2 (Abr-Jun)": "segundo trimestre",
                "Trimestre 3 (Jul-Sep)": "tercer trimestre",
                "Trimestre 4 (Oct-Dic)": "cuarto trimestre"
            }
            
            trimestre_formateado = mapa_trimestre.get(trimestre_str_largo, trimestre_str_largo)
            placeholder_final_trimestre = f"{trimestre_formateado} de {anio_str}"
            # --- FIN: Formatear Trimestre (Req. 1) ---

            label_ce_extremo = "CE1" if aplicar_ce2_bi_extremo else "CE"
            # f. Añadir datos del extremo a la lista para el bucle
            lista_extremos_plantilla_word.append({
                'loop_index': j + 1,
                'numeral': f"{num_hecho}.{j + 1}",
                'descripcion': f"Cálculo para el Extremo {j+1}: Manifiesto {extremo.get('trimestre', 'N/A')} {extremo.get('anio', 'N/A')} ({tipo_inc})",
                'label_ce_principal': label_ce_extremo, # <-- AÑADIR
                
                # --- INICIO DE LA ADICIÓN (Req. 1) ---
                'trimestre_manifiesto': placeholder_final_trimestre,
                # --- FIN DE LA ADICIÓN (Req. 1) ---

                # --- INICIO DE LA ADICIÓN (INF007) ---
                'fecha_incumplimiento_larga': (format_date(fecha_inc, "d 'de' MMMM 'de' yyyy", locale='es') if fecha_inc else "N/A").replace("septiembre", "setiembre").replace("Septiembre", "Setiembre"),
                'fecha_extemporanea_larga': (format_date(fecha_ext, "d 'de' MMMM 'de' yyyy", locale='es') if fecha_ext else "N/A").replace("septiembre", "setiembre").replace("Septiembre", "Setiembre"),
                # --- FIN DE LA ADICIÓN ---
                'tabla_ce1': tabla_ce1_cuerpo,
                'tabla_ce2': tabla_ce2_cuerpo,
                'aplicar_capacitacion': aplicar_ce2_bi_extremo,
                'tabla_bi': tabla_bi_cuerpo,
                'bi_uit_extremo': f"{bi_uit:,.3f} UIT",
                'texto_razonabilidad': RichText(""), # Placeholder si INF007 necesita texto de razonabilidad
                # --- INICIO DE LA CORRECCIÓN DE FORMATO ---
                'fecha_max_presentacion': format_date(extremo.get('fecha_maxima_presentacion'), "d 'de' MMMM 'de' yyyy", locale='es') if extremo.get('fecha_maxima_presentacion') else "N/A",
                'fecha_extemporanea': format_date(fecha_ext, "d 'de' MMMM 'de' yyyy", locale='es') if fecha_ext else "N/A",
                # --- FIN DE LA CORRECCIÓN DE FORMATO ---
            })
        # --- FIN DEL BUCLE DE EXTREMOS ---

        # 5. Post-Cálculo: Multa Final
        if not lista_bi_resultados_completos: return {'error': 'No se pudo calcular BI para ningún extremo.'}
        res_multa_final = calcular_multa({**datos_comunes, 'beneficio_ilicito': total_bi_uit})
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

        tabla_multa_final_subdoc = create_main_table_subdoc( tpl_principal, ["Componentes", "Monto"], res_multa_final.get('multa_data_raw', []), ['Componentes', 'Monto'], texto_posterior="Elaboración: Subdirección de Sanción y Gestión de Incentivos (SSAG) - DFAI.", estilo_texto_posterior='FuenteTabla', column_widths=(5, 1) )

        # --- INICIO: Generar Texto Desglose BI (Req. 2) ---
        lista_desglose = []
        # Iterar sobre los resultados que ya calculamos
        for i, ext_res in enumerate(resultados_app.get('extremos', [])):
            bi_valor = ext_res.get('bi_uit', 0.0)
            # Usar el "tipo" como descripción (ej. "Manifiesto Trimestre 1 2023 (No presentó)")
            tipo_desc = f"extremo n.° {i+1}"
            lista_desglose.append(f"{bi_valor:,.3f} UIT del {tipo_desc}")
        
        texto_desglose_bi = ""
        num_extremos_bi = len(lista_desglose)
        if num_extremos_bi == 1:
            texto_desglose_bi = lista_desglose[0]
        elif num_extremos_bi == 2:
            texto_desglose_bi = " y ".join(lista_desglose)
        elif num_extremos_bi > 2:
            # Une todos menos el último con comas, luego añade ", y " y el último
            texto_desglose_bi = ", ".join(lista_desglose[:-1]) + ", y " + lista_desglose[-1]
        # --- FIN: Generar Texto Desglose BI (Req. 2) ---

        # --- INICIO: Buscar Precios Base de CE2 (Múltiple) ---
        ce2_base_precio_soles = 0.0
        ce2_base_precio_dolares = 0.0
        if aplicar_capacitacion_general:
            for extremo_res in resultados_app.get('extremos', []):
                if extremo_res.get('ce2_data'):
                    try:
                        # Tomar el precio del primer ítem de la primera capacitación encontrada
                        primer_item_ce2 = extremo_res['ce2_data'][0]
                        ce2_base_precio_soles = primer_item_ce2.get('precio_soles', 0.0)
                        ce2_base_precio_dolares = primer_item_ce2.get('precio_dolares', 0.0)
                        break # Salir en cuanto se encuentre el primer precio
                    except IndexError:
                        continue # Este extremo tiene 'ce2_data' pero está vacío, buscar en el siguiente
        # --- FIN: Buscar Precios Base de CE2 (Múltiple) ---

        # --- AÑADIR ESTA LÍNEA ---
        es_ext = any(e.get('tipo_extremo') == 'Presentó fuera de plazo' for e in datos_hecho['extremos'])
        # 6. Contexto Final y Renderizado
        contexto_final = {
            **datos_comunes['context_data'], 'acronyms': datos_comunes['acronym_manager'], 'es_extemporaneo': es_ext,
            'hecho': {
                'numero_imputado': num_hecho,
                'descripcion': RichText(datos_hecho.get('texto_hecho', '')),
                'lista_extremos': lista_extremos_plantilla_word,
             },
            'numeral_hecho': f"IV.{num_hecho + 1}",
            'bi_uit_total': f"{total_bi_uit:,.3f} UIT",
            # --- INICIO DE LA ADICIÓN (Req. 2) ---
            'texto_desglose_bi': texto_desglose_bi,
            # --- FIN DE LA ADICIÓN (Req. 2) ---
            'multa_original_uit': f"{multa_final_uit:,.3f} UIT", # Multa original
            'mh_uit': f"{multa_final_del_hecho_uit:,.3f} UIT", # Multa final (con reducción/tope)
            'tabla_multa_final': tabla_multa_final_subdoc,
            'tabla_detalle_personal': tabla_pers_subdoc_final,
            'se_usa_capacitacion': aplicar_capacitacion_general,
            'num_personal_total_texto': texto_con_numero(num_pers_total_int, 'f') if aplicar_capacitacion_general else '',
            'texto_explicacion_prorrateo': '', # Se genera en app.py
            
            # --- AÑADIR ESTAS LÍNEAS ---
            'ce2_monto_soles': f"S/ {resultados_app['totales']['ce2_total_soles_calculado']:,.3f}",
            'ce2_monto_dolares': f"US$ {resultados_app['totales']['ce2_total_dolares_calculado']:,.3f}",
            # --- FIN ---

            # --- INICIO: ADICIÓN PRECIOS BASE (Múltiple) ---
            'ce2_precio_base_soles': f"S/ {ce2_base_precio_soles:,.3f}",
            'ce2_precio_base_dolares': f"US$ {ce2_base_precio_dolares:,.3f}",
            # --- FIN: ADICIÓN PRECIOS BASE (Múltiple) ---

            # --- INICIO: (REQ 1) PLACEHOLDERS DE REDUCCIÓN Y TOPE ---
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
            # --- FIN: (REQ 1) ---
        }
        tpl_principal.render(contexto_final, autoescape=True, jinja_env=jinja_env)
        buf_final = io.BytesIO(); tpl_principal.save(buf_final)

        # 7. Preparar datos para App
        resultados_app['totales'] = {
            **resultados_app['totales'], 
            'beneficio_ilicito_uit': total_bi_uit, 
            'multa_data_raw': res_multa_final.get('multa_data_raw', []), 
            'multa_final_uit': multa_final_uit, # Multa original
            'bi_data_raw': lista_bi_resultados_completos, 
            'tabla_personal_data': tabla_pers_data,
            
            # --- INICIO: (REQ 1) DATOS DE REDUCCIÓN PARA APP ---
            'aplica_reduccion': aplica_reduccion_str,
            'porcentaje_reduccion': porcentaje_str,
            'multa_con_reduccion_uit': multa_con_reduccion_uit, # Multa después de 50/30
            'multa_reducida_uit': multa_reducida_uit, # <-- LA VARIABLE QUE FALTABA
            'multa_final_aplicada': multa_final_del_hecho_uit # Multa final CON tope
            # --- FIN: (REQ 1) ---
        }
        # 8. Devolver resultados
        return {
            'doc_pre_compuesto': buf_final,
            'resultados_para_app': resultados_app,
            'texto_explicacion_prorrateo': '',
            'tabla_detalle_personal': tabla_pers_subdoc_final,
            'usa_capacitacion': aplicar_capacitacion_general,
            'es_extemporaneo': any(e.get('tipo_extremo') == 'Presentó fuera de plazo' for e in datos_hecho['extremos']),
            'anexos_ce_generados': anexos_ce,
            'ids_anexos': list(filter(None, anexos_ids)),
            'tabla_personal_data': tabla_pers_data
            # --- NOTA: Los datos de reducción se leen desde resultados_app['totales'] ---
        }
    except Exception as e:
        import traceback; traceback.print_exc(); return {'error': f"Error _procesar_multiple INF007: {e}"}
