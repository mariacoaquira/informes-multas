# --- Archivo: infracciones/INF011.py ---

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

# ---------------------------------------------------------------------
# 1. FUNCIÓN DE FECHAS (PLAZO FIJO SEGÚN NORMATIVA O AÑO)
# ---------------------------------------------------------------------
def _calcular_fechas_inf011(fecha_supervision):
    """
    Calcula fechas para el Plan de Contingencia.
    La fecha de incumplimiento es igual a la fecha de supervisión.
    """
    if not fecha_supervision: return None, None
    
    fecha_limite = fecha_supervision
    fecha_incumplimiento = fecha_supervision
    return fecha_limite, fecha_incumplimiento

# ---------------------------------------------------------------------
# 2. MOTOR DE CÁLCULO CE (SIN CANTIDAD NI HORAS)
# ---------------------------------------------------------------------
def _calcular_costo_evitado_inf011_interno(datos_comunes, extremo_data):
    result = {'items_calculados': [], 'ce_soles': 0.0, 'ce_dolares': 0.0, 
              'fuentes': {'ce1': {'placeholders_dinamicos': {}}}, 'ids_anexos': set(), 'error': None,
              'fi_mes': '', 'fi_ipc': 0.0, 'fi_tc': 0.0} # Inicializar llaves en el root
    
    try:
        df_items_inf = datos_comunes.get('df_items_infracciones')
        df_costos = datos_comunes.get('df_costos_items')
        df_sal = datos_comunes.get('df_salarios_general')
        df_coti = datos_comunes.get('df_coti_general')
        df_ind = datos_comunes.get('df_indices')
        id_rubro = datos_comunes.get('id_rubro_seleccionado')
        id_inf = 'INF011'
        
        fecha_inc = extremo_data.get('fecha_incumplimiento')
        fecha_final_dt = pd.to_datetime(fecha_inc)

        # 1. CAPTURA GLOBAL DE DATOS DE INCUMPLIMIENTO (Igual a INF004)
        ipc_row_inc = df_ind[df_ind['Indice_Mes'].dt.to_period('M') == fecha_final_dt.to_period('M')]
        if ipc_row_inc.empty:
            result['error'] = f"Sin IPC/TC para {fecha_final_dt.strftime('%B %Y')}"
            return result
        ipc_inc, tc_inc = ipc_row_inc.iloc[0]['IPC_Mensual'], ipc_row_inc.iloc[0]['TC_Mensual']
        
        # Guardar en el root para que el procesador los vea
        result['fi_mes'] = format_date(fecha_final_dt, "MMMM 'de' yyyy", locale='es').replace("septiembre", "setiembre")
        result['fi_ipc'] = float(ipc_inc)
        result['fi_tc'] = float(tc_inc)

        receta = df_items_inf[df_items_inf['ID_Infraccion'] == id_inf]
        if receta.empty:
            result['error'] = f"No hay receta para {id_inf}"; return result

        items_final = []
        fuentes_ce1 = {'placeholders_dinamicos': {}}
        salario_capturado = False

        for _, item_receta in receta.iterrows():
            id_item = item_receta['ID_Item_Infraccion']
            desc_item = item_receta.get('Nombre_Item', 'N/A')
            
            costos_posibles = df_costos[df_costos['ID_Item_Infraccion'] == id_item].copy()
            if costos_posibles.empty: continue
            
            id_rubro_str = str(id_rubro) if id_rubro else ''
            df_candidatos = costos_posibles[costos_posibles['ID_Rubro'].astype(str).str.contains(fr'\b{id_rubro_str}\b', regex=True, na=False)].copy()
            if df_candidatos.empty:
                df_candidatos = costos_posibles[costos_posibles['ID_Rubro'].isin(['', 'nan', None])].copy()

            # Búsqueda de fecha más cercana
            fechas_f = []
            for _, cand in df_candidatos.iterrows():
                id_gen = cand['ID_General']; f_f = pd.NaT
                if pd.notna(id_gen):
                    if 'SAL' in id_gen:
                        f = df_sal[df_sal['ID_Salario'] == id_gen]
                        if not f.empty: f_f = pd.to_datetime(f"{int(f.iloc[0]['Costeo_Salario'])}-12-31")
                    elif 'COT' in id_gen:
                        f = df_coti[df_coti['ID_Cotizacion'] == id_gen]
                        if not f.empty: f_f = pd.to_datetime(f.iloc[0]['Fecha_Costeo'])
                fechas_f.append(f_f)
            
            df_candidatos['Fecha_Fuente'] = fechas_f
            df_candidatos.dropna(subset=['Fecha_Fuente'], inplace=True)
            if df_candidatos.empty: continue
            costo_final = df_candidatos.loc[(df_candidatos['Fecha_Fuente'] - fecha_final_dt).dt.days.abs().idxmin()]
            
            # IPC Fuente
            id_gen = costo_final['ID_General']; f_f = costo_final['Fecha_Fuente']
            if 'SAL' in id_gen:
                idx_anio = df_ind[df_ind['Indice_Mes'].dt.year == f_f.year]
                ipc_cost = float(idx_anio['IPC_Mensual'].mean()) if not idx_anio.empty else 1.0
                if not salario_capturado:
                    f_row = df_sal[df_sal['ID_Salario'] == id_gen]
                    fuentes_ce1['fuente_salario'] = f_row.iloc[0].get('Fuente_Salario','')
                    fuentes_ce1['pdf_salario'] = f_row.iloc[0].get('PDF_Salario','')
                    fuentes_ce1['placeholders_dinamicos']['ref_ipc_salario'] = f"Promedio {f_f.year}, IPC = {ipc_cost}"
                    salario_capturado = True
            else:
                idx_row = df_ind[df_ind['Indice_Mes'].dt.to_period('M') == f_f.to_period('M')]
                ipc_cost = float(idx_row.iloc[0]['IPC_Mensual']) if not idx_row.empty else 1.0

            precio_s = float(costo_final.get('Costo_Unitario_Item', 0.0))
            factor = redondeo_excel(ipc_inc / ipc_cost, 3)
            monto_s = redondeo_excel(precio_s * factor, 3)
            
            items_final.append({
                "descripcion": desc_item, "precio_soles": precio_s, "factor_ajuste": factor,
                "monto_soles": monto_s, "monto_dolares": redondeo_excel(monto_s / tc_inc, 3)
            })
            if costo_final.get('ID_Anexo_Drive'): result['ids_anexos'].add(costo_final.get('ID_Anexo_Drive'))

        result.update({
            'items_calculados': items_final, 
            'ce_soles': sum(x['monto_soles'] for x in items_final),
            'ce_dolares': sum(x['monto_dolares'] for x in items_final),
            'fuentes': {'ce1': fuentes_ce1}
        })
        return result
    except Exception as e:
        result['error'] = str(e); return result
    
# ---------------------------------------------------------------------
# 3. INTERFAZ DE USUARIO (CON SUPERVISIÓN Y EXTEMPORANEIDAD)
# ---------------------------------------------------------------------
# CÓDIGO CORREGIDO EN INF011.py
def renderizar_inputs_especificos(i, df_dias_no_laborables=None):
    st.markdown("##### Detalles del Incumplimiento: Plan de Contingencia (INF011)")
    datos_hecho = st.session_state.imputaciones_data[i]

    if 'extremos' not in datos_hecho: datos_hecho['extremos'] = [{}]
    
    # ... (botón añadir extremo igual) ...

    for j, extremo in enumerate(datos_hecho['extremos']):
        with st.container(border=True):
            st.markdown(f"**Extremo n.° {j + 1}**")
            
            col_fecha, col_tipo = st.columns(2)
            with col_fecha:
                # 1. Aseguramos que la llave sea consistente
                fecha_sup = st.date_input("Fecha de la supervisión", key=f"sup_fecha_{i}_{j}", value=extremo.get('fecha_supervision'))
                extremo['fecha_supervision'] = fecha_sup
                # ASIGNACIÓN CRÍTICA: La fecha de incumplimiento es la de supervisión
                extremo['fecha_incumplimiento'] = fecha_sup 

            with col_tipo:
                tipo = st.radio("Tipo de incumplimiento", ["No presentó", "Presentó fuera de plazo"], key=f"tipo_{i}_{j}", horizontal=True)
                extremo['tipo_extremo'] = tipo

            if fecha_sup: 
                if tipo == "Presentó fuera de plazo":
                    extremo['fecha_extemporanea'] = st.date_input("Fecha de cumplimiento extemporáneo", min_value=fecha_sup, key=f"ext_{i}_{j}", value=extremo.get('fecha_extemporanea'))
                else:
                    extremo['fecha_extemporanea'] = None
                st.info(f"🚨 **Fecha de incumplimiento detectada:** {fecha_sup.strftime('%d/%m/%Y')}")

    return datos_hecho

def validar_inputs(datos_hecho):
    if not datos_hecho.get('extremos'): return False
    for ex in datos_hecho['extremos']:
        # Validamos que exista la fecha de supervisión y el tipo de extremo
        if not all([ex.get('fecha_supervision'), ex.get('tipo_extremo')]): return False
        if ex['tipo_extremo'] == "Presentó fuera de plazo" and not ex.get('fecha_extemporanea'): return False
    return True

# ---------------------------------------------------------------------
# 5. PROCESADOR (SIMPLE Y MÚLTIPLE)
# ---------------------------------------------------------------------
def _procesar_hecho_simple(datos_comunes, datos_hecho):
    try:
        jinja_env = Environment(trim_blocks=True, lstrip_blocks=True)
        id_inf = 'INF011'
        fila_inf = datos_comunes['df_tipificacion'][datos_comunes['df_tipificacion']['ID_Infraccion'] == id_inf].iloc[0]
        
        doc_tpl_bi = DocxTemplate(descargar_archivo_drive(fila_inf.get('ID_Plantilla_BI')))
        tpl_anexo = DocxTemplate(descargar_archivo_drive(fila_inf.get('ID_Plantilla_CE')))

        extremo = datos_hecho['extremos'][0]
        res_ce = _calcular_costo_evitado_inf011_interno(datos_comunes, extremo)
        
        # 1. BI y Multa
        tipo_inc = extremo['tipo_extremo']
        es_extemporaneo = (tipo_inc == "Presentó fuera de plazo")
        
        # Obtenemos los totales directamente del resultado del CE
        total_ce_s = res_ce.get('ce_soles', 0)
        total_ce_d = res_ce.get('ce_dolares', 0)

        datos_bi = {
            **datos_comunes, 
            'ce_soles': total_ce_s, 
            'ce_dolares': total_ce_d, 
            'fecha_incumplimiento': extremo.get('fecha_incumplimiento'), 
            'texto_del_hecho': datos_hecho.get('texto_hecho', '')
        }

        if es_extemporaneo:
            res_bi = calcular_beneficio_ilicito_extemporaneo({**datos_bi, 'fecha_cumplimiento_extemporaneo': extremo['fecha_extemporanea'], **calcular_beneficio_ilicito(datos_bi)})
        else:
            res_bi = calcular_beneficio_ilicito(datos_bi)
            
        bi_uit = res_bi.get('beneficio_ilicito_uit', 0)
        res_multa = calcular_multa({**datos_comunes, 'beneficio_ilicito': bi_uit, 'factor_f': datos_hecho.get('factor_f_calculado', 1.0)})

        # Lógica de Moneda
        es_dolares = (res_bi.get('moneda_cos', 'USD') == 'USD')
        
        # Reducción y Tope
        datos_c = datos_comunes.get('datos_hecho_completos', {})
        multa_uit = res_multa.get('multa_final_uit', 0)
        multa_red = redondeo_excel(multa_uit * (0.5 if datos_c.get('porcentaje_reduccion') == '50%' else 0.7), 3) if datos_c.get('aplica_reduccion') == 'Sí' else multa_uit
        tope_uit = float(fila_inf.get('Tope_Multa_Infraccion', 1000))
        multa_final = min(multa_red, tope_uit)

        # --- LÓGICA DE COMPACTACIÓN DE NOTAS BI (IGUAL A INF004) ---
        filas_bi_crudas = res_bi.get('table_rows', [])
        fn_map_orig = res_bi.get('footnote_mapping', {})
        fn_data = res_bi.get('footnote_data', {})
        
        # 1. Identificar las letras/fuentes realmente usadas en las filas
        letras_usadas = sorted(list({r for f in filas_bi_crudas if f.get('ref') for r in f.get('ref').replace(" ", "").split(",") if r}))
        
        # 2. Crear mapa de traducción a letras correlativas (a, b, c...)
        letras_base = "abcdefghijklmnopqrstuvwxyz"
        map_traduccion = {v: letras_base[i] for i, v in enumerate(letras_usadas)}
        nuevo_fn_map = {map_traduccion[v]: fn_map_orig[v] for v in letras_usadas if v in fn_map_orig}

        # 3. Construir nuevas filas para la tabla incluyendo los superíndices en el texto
        filas_bi_para_tabla = []
        for fila in filas_bi_crudas:
            ref_orig = fila.get('ref', '')
            # Empezamos con el superíndice técnico del cálculo (ej: T para tiempo)
            super_final = str(fila.get('descripcion_superindice', ''))
            
            # Añadimos las letras de las fuentes entre paréntesis
            if ref_orig:
                nuevas_letras = [map_traduccion[r] for r in ref_orig.replace(" ", "").split(",") if r in map_traduccion]
                if nuevas_letras: 
                    super_final += f"({', '.join(nuevas_letras)})"
            
            filas_bi_para_tabla.append({
                'descripcion_texto': fila.get('descripcion_texto', ''),
                'descripcion_superindice': super_final, # Se pasa a la columna Descripción
                'monto': fila.get('monto', '')
            })

        # 4. Generar la lista de textos para el pie de página
        fn_list = [f"({l}) {obtener_fuente_formateada(k, fn_data, 'INF011', es_extemporaneo)}" for l, k in sorted(nuevo_fn_map.items())]
        fn_dict = {
            'list': fn_list, 
            'elaboration': 'Elaboración: Subdirección de Sanción y Gestión de Incentivos (SSAG) – DFAI.', 
            'style': 'FuenteTabla'
        }

        # Generación de la tabla con proporciones y fuentes
        tabla_bi_subdoc = create_main_table_subdoc(
            doc_tpl_bi, 
            ["Descripción", "Monto"], 
            filas_bi_para_tabla, # Usar las filas procesadas arriba
            ['descripcion_texto', 'monto'], 
            footnotes_data=fn_dict, 
            column_widths=(5, 1) # Proporción correcta (Descripción más ancha)
        )

        # --- GENERACIÓN DE TABLA CE (5 COLUMNAS) ---
        ce_table_formatted = []
        for item in res_ce['items_calculados']:
            ce_table_formatted.append({
                'descripcion': f"{item.get('descripcion', '')} 1/",
                'precio_soles': f"S/ {item.get('precio_soles', 0):,.3f}", 
                'factor_ajuste': f"{item.get('factor_ajuste', 0):,.3f}",
                'monto_soles': f"S/ {item.get('monto_soles', 0):,.3f}",
                'monto_dolares': f"US$ {item.get('monto_dolares', 0):,.3f}"
            })
        
        # Añadir Fila Total
        ce_table_formatted.append({
            'descripcion': 'Total', 'precio_soles': '', 'factor_ajuste': '',
            'monto_soles': f"S/ {total_ce_s:,.3f}",
            'monto_dolares': f"US$ {total_ce_d:,.3f}"
        })

        tabla_ce_subdoc = create_table_subdoc(
            tpl_anexo, # O doc_tpl_bi según uses
            headers=["Descripción", "Precio asociado (S/)", "Factor de ajuste 2/", "Monto (*) (S/)", "Monto (*) (US$) 3/"],
            data=ce_table_formatted,
            keys=['descripcion', 'precio_soles', 'factor_ajuste', 'monto_soles', 'monto_dolares']
        )

        numero_hecho = datos_comunes['numero_hecho_actual']
        ctx = {
            **datos_comunes['context_data'], 
            **res_ce['fuentes']['ce1'].get('placeholders_dinamicos', {}),
            'numeral_hecho': f"IV.{numero_hecho + 1}",
            'multa_original_uit': f"{multa_uit:,.3f} UIT",
            # CAPTURA DE DATOS DE TIEMPO (Igual a INF004)
            'fi_mes': res_ce.get('fi_mes', ''),
            'fi_ipc': f"{res_ce.get('fi_ipc', 0)}",
            'fi_tc': f"{res_ce.get('fi_tc', 0)}",
            'hecho': {
                'numero_imputado': numero_hecho, 
                'descripcion': RichText(datos_hecho.get('texto_hecho', '')),
                'tabla_ce': tabla_ce_subdoc,
                'tabla_bi': tabla_bi_subdoc,
                'tabla_multa': create_main_table_subdoc(
                    doc_tpl_bi, 
                    ["Componente", "Monto"], 
                    res_multa['multa_data_raw'], 
                    ['Componentes', 'Monto'], 
                    texto_posterior="Elaboración: Subdirección de Sanción y Gestión de Incentivos (SSAG) – DFAI.", 
                    estilo_texto_posterior='FuenteTabla', 
                    column_widths=(5, 1)
                )
            },
            # Resto de tus placeholders (anio_plan, bi_uit, mh_uit, etc.) se mantienen igual
            'anio_plan': extremo['fecha_supervision'].year, 
            'bi_uit': f"{bi_uit:,.3f} UIT", 
            'mh_uit': f"{multa_final:,.3f} UIT",
            'fecha_supervision': format_date(extremo['fecha_supervision'], "d 'de' MMMM 'de' yyyy", locale='es') if extremo.get('fecha_supervision') else "N/A",
            'fuente_salario_ce1': res_ce['fuentes']['ce1'].get('fuente_salario',''), 
            'pdf_salario_ce1': res_ce['fuentes']['ce1'].get('pdf_salario',''),
            'ph_anexo_ce_num': "3" if datos_hecho.get('aplica_graduacion') == 'Sí' else "2",
            'bi_moneda_es_dolares': es_dolares, 
            'ph_bi_moneda_simbolo': "US$" if es_dolares else "S/",
            'ph_bi_moneda_texto': "moneda extranjera (Dólares)" if es_dolares else "moneda nacional (Soles)"
        }

        doc_tpl_bi.render(ctx, jinja_env=jinja_env); b_bi = io.BytesIO(); doc_tpl_bi.save(b_bi)
        tpl_anexo.render(ctx, jinja_env=jinja_env); b_anx = io.BytesIO(); tpl_anexo.save(b_anx)

        # ESTRUCTURA DE RETORNO CORREGIDA PARA LA APP
        # ESTRUCTURA DE RETORNO PARA QUE LA APP DIBUJE LAS TABLAS
        resultados_app = {
             'extremos': [{
                  'tipo': tipo_inc, 
                  'ce_data': res_ce['items_calculados'], # Clave para que la App la vea
                  'ce1_data': res_ce['items_calculados'], # Backup
                  'ce_soles': total_ce_s, 
                  'ce_dolares': total_ce_d,
                  'bi_data': res_bi.get('table_rows', []), 
                  'bi_uit': bi_uit,
                  'aplicar_ce2_a_bi': False
             }],
             'totales': {
                  'multa_final_aplicada': multa_final,
                  'beneficio_ilicito_uit': bi_uit,
                  'multa_final_uit': multa_uit,
                  'bi_data_raw': res_bi['table_rows'],
                  'multa_data_raw': res_multa['multa_data_raw'],
                  'aplica_reduccion': datos_c.get('aplica_reduccion', 'No'),
                  'porcentaje_reduccion': datos_c.get('porcentaje_reduccion', '0%'),
                  'multa_con_reduccion_uit': multa_red,
                  'ce_total_soles': total_ce_s,
                  'ce_total_dolares': total_ce_d
             },
             # Para que INF004-style display funcione
             'ce_data_raw': res_ce['items_calculados']
        }
        
        return {
            'doc_pre_compuesto': b_bi, 
            'anexos_ce_generados': [b_anx], 
            'resultados_para_app': resultados_app,
            'es_extemporaneo': es_extemporaneo,
            'ids_anexos': list(res_ce.get('ids_anexos', []))
        }
    except Exception as e: return {'error': str(e)}

def procesar_infraccion(datos_comunes, datos_hecho):
    return _procesar_hecho_simple(datos_comunes, datos_hecho)