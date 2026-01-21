# --- Archivo: infracciones/INF006.py ---

# --- BIBLIOTECAS ---
import streamlit as st
import pandas as pd
from datetime import date
from docxtpl import DocxTemplate, RichText
from docxcompose.composer import Composer
import io
from num2words import num2words
from babel.dates import format_date

# --- IMPORTACIONES DE MÓDulos PROPIOS ---
from textos_manager import obtener_fuente_formateada
from funciones import (create_table_subdoc, create_main_table_subdoc, texto_con_numero,
                     create_footnotes_subdoc, create_consolidated_bi_table_subdoc,
                     create_personal_table_subdoc) # Funciones de formato de tablas/texto
from sheets import (calcular_beneficio_ilicito, calcular_multa,
                    descargar_archivo_drive) # Funciones de cálculo BI/Multa y Drive
# --- IMPORTAR EL MÓDULO DE CAPACITACIÓN ---
try:
    from modulos.calculo_capacitacion import calcular_costo_capacitacion
except ImportError:
    st.error("No se pudo importar 'calcular_costo_capacitacion'. Verifica la ruta del archivo.")
    def calcular_costo_capacitacion(*args, **kwargs):
        return {'error': 'Módulo de cálculo de capacitación no encontrado.'}

# --- 2. INTERFAZ DE USUARIO (Adaptada al estilo INF003) ---
def renderizar_inputs_especificos(i, df_dias_no_laborables=None):
    """
    Renderiza la interfaz para INF006:
    1. Lista dinámica de extremos, pidiendo la fecha de presentación de info. falsa.
    2. Tabla editable para personal a capacitar (anidada en el primer extremo).
    """
    st.markdown("##### Detalles de la Presentación de Información Falsa")
    datos_hecho = st.session_state.imputaciones_data[i] # Acceder a los datos de este hecho

    # --- SECCIÓN 1: EXTREMOS DEL INCUMPLIMIENTO ---
    st.markdown("###### **Extremos del incumplimiento**")
    
    # Inicializar lista de personal y extremos si no existen
    if 'tabla_personal' not in datos_hecho or not isinstance(datos_hecho['tabla_personal'], list):
        datos_hecho['tabla_personal'] = [{'Perfil': 'Gerente General', 'Descripción': 'Se considera la capacitación del Gerente General, dada su responsabilidad directa en la toma de decisiones estratégicas y el cumplimiento de las obligaciones legales de la empresa. Como máxima autoridad, su conocimiento en gestión ambiental y normativas aplicables es fundamental para garantizar que la documentación presentada a las autoridades sea veraz y cumpla con los estándares exigidos, minimizando así el riesgo de futuras infracciones.', 'Cantidad': 1}]
    if 'extremos' not in datos_hecho: 
        datos_hecho['extremos'] = [{}]

    # Botón para añadir nuevo extremo
    if st.button("+ Añadir Extremo", key=f"add_extremo_{i}"):
        datos_hecho['extremos'].append({}) # Añadir diccionario vacío
        st.rerun() # Recargar interfaz

    # Iterar sobre los extremos existentes
    for j, extremo in enumerate(datos_hecho['extremos']):
        with st.container(border=True): # Contenedor visual para cada extremo
            
            # --- Columna de Fecha y Métrica ---
            col1, col_display, col_button = st.columns([2, 2, 1])

            with col1:
                # --- CAMBIO CLAVE: Etiqueta de la Fecha ---
                fecha_presentacion_falsa = st.date_input(
                    f"Fecha de presentación de la información falsa", # Etiqueta específica de INF006
                    key=f"fecha_presentacion_falsa_{i}_{j}", # Clave única
                    value=extremo.get('fecha_incumplimiento'), # Guardar en 'fecha_incumplimiento'
                    format="DD/MM/YYYY",
                    max_value=date.today() # No permitir fechas futuras
                )
                extremo['fecha_incumplimiento'] = fecha_presentacion_falsa
                extremo['fecha_base'] = fecha_presentacion_falsa

            with col_display:
                if extremo.get('fecha_incumplimiento'):
                    fecha_str = extremo['fecha_incumplimiento'].strftime('%d/%m/%Y')
                    st.metric(label="Fecha de incumplimiento", value=fecha_str) # <-- Estilo métrica
                else:
                    st.metric(label="Fecha de incumplimiento", value="---")

            with col_button:
                if len(datos_hecho['extremos']) > 1:
                    st.write("") 
                    st.write("") 
                    if st.button(f"🗑️", key=f"del_extremo_{i}_{j}", help="Eliminar Extremo"):
                        datos_hecho['extremos'].pop(j)
                        st.rerun()
            
            st.divider() 
            
            # --- SECCIÓN 2: DATOS DE CAPACITACIÓN (Anidada) ---
            st.markdown("###### **Personal a capacitar**")

            df_personal = pd.DataFrame(datos_hecho['tabla_personal'])

            if j == 0:
                # Editor EDITABLE para el primer extremo (j==0)
                edited_df = st.data_editor(
                    df_personal,
                    num_rows="dynamic",
                    key=f"data_editor_personal_{i}",
                    hide_index=True,
                    use_container_width=True,
                    disabled=False, # <-- Editable
                    column_config={ 
                        "Perfil": st.column_config.TextColumn("Perfil", help="Ej: Personal operativo, Supervisor", required=True),
                        "Descripción": st.column_config.TextColumn("Descripción", help="Detalle de las funciones...", width="large"),
                        "Cantidad": st.column_config.NumberColumn("Cantidad", help="Número de personas con este perfil", min_value=0, step=1, required=True, format="%d"), # Formato Entero
                    }
                )
                datos_hecho['tabla_personal'] = edited_df.to_dict('records') # Guardar cambios
            
            else:
                # Editor DESHABILITADO para los siguientes extremos
                 st.data_editor(
                    df_personal,
                    num_rows="dynamic",
                    key=f"data_editor_personal_{i}_disabled_{j}",
                    hide_index=True,
                    use_container_width=True,
                    disabled=True, # <-- Deshabilitado
                    column_config={ "Perfil": {}, "Descripción": {}, "Cantidad": st.column_config.NumberColumn(format="%d") } # Formato Entero
                )

            # Calcular y mostrar el total en CADA extremo
            cantidades_num = [pd.to_numeric(p.get('Cantidad'), errors='coerce') for p in datos_hecho['tabla_personal']]
            total_personal = pd.Series(cantidades_num).fillna(0).sum()
            datos_hecho['num_personal_capacitacion'] = int(total_personal)
            st.metric("Total de Personal a Capacitar", f"{datos_hecho['num_personal_capacitacion']} persona(s)") # <-- Estilo métrica

    return datos_hecho # Devolver datos actualizados

# --- 3. VALIDACIÓN DE INPUTS (Idéntica a INF003) ---
def validar_inputs(datos_especificos):
    """
    Valida que se haya ingresado el personal y al menos una fecha de incumplimiento.
    """
    if not datos_especificos.get('num_personal_capacitacion', 0) > 0:
        return False
    if not datos_especificos.get('extremos'):
        return False
    for extremo in datos_especificos['extremos']:
        if not extremo.get('fecha_incumplimiento'):
            return False
    return True

# --- 4. DESPACHADOR PRINCIPAL (Idéntico a INF003) ---
def procesar_infraccion(datos_comunes, datos_especificos):
    """
    Decide si procesar como hecho simple (1 extremo) o múltiple (>1 extremo)
    y llama a la función correspondiente.
    """
    num_extremos = len(datos_especificos.get('extremos', []))

    if num_extremos == 0:
        return {'error': 'No se ha registrado ningún extremo (fecha) para este hecho.'}
    elif num_extremos == 1:
        return _procesar_hecho_simple(datos_comunes, datos_especificos)
    else: # num_extremos > 1
        return _procesar_hecho_multiple(datos_comunes, datos_especificos)
    
# --- 5. PROCESAR HECHO SIMPLE (Adaptado de INF003) ---
def _procesar_hecho_simple(datos_comunes, datos_especificos):
    """
    Procesa un hecho INF006 con un único extremo (fecha), usando plantilla simple.
    """
    try:
        # 1. Cargar plantilla simple BI y CE
        df_tipificacion, id_infraccion = datos_comunes['df_tipificacion'], datos_comunes['id_infraccion']
        fila_inf = df_tipificacion[df_tipificacion['ID_Infraccion'] == id_infraccion].iloc[0]
        id_tpl_bi, id_tpl_ce = fila_inf.get('ID_Plantilla_BI'), fila_inf.get('ID_Plantilla_CE')
        if not id_tpl_bi or not id_tpl_ce: return {'error': f'Faltan IDs plantilla simple (BI o CE) para {id_infraccion}.'}
        buf_bi, buf_ce = descargar_archivo_drive(id_tpl_bi), descargar_archivo_drive(id_tpl_ce)
        if not buf_bi or not buf_ce: return {'error': f'Fallo descarga plantilla simple para {id_infraccion}.'}
        doc_tpl_bi = DocxTemplate(buf_bi)
        tpl_anexo = DocxTemplate(buf_ce)

        # 2. Calcular CE (solo capacitación)
        extremo = datos_especificos['extremos'][0]
        fecha_inc = extremo['fecha_incumplimiento'] # Fecha única
        datos_para_ce = {**datos_comunes, 'fecha_incumplimiento': fecha_inc}
        res_ce = calcular_costo_capacitacion(num_personal=datos_especificos.get('num_personal_capacitacion', 1), datos_comunes=datos_para_ce)
        if res_ce.get('error'): return {'error': f"Error CE: {res_ce['error']}"}
        ce_data_raw = res_ce.get('items_calculados', [])
        total_soles = sum(item.get('monto_soles', 0) for item in ce_data_raw)
        total_dolares = sum(item.get('monto_dolares', 0) for item in ce_data_raw)

        # 3. Calcular BI y Multa
        # Usar el texto del hecho de la UI, como en INF003
        datos_bi_base = {**datos_comunes, 'ce_soles': total_soles, 'ce_dolares': total_dolares, 'fecha_incumplimiento': fecha_inc, 
                         'texto_del_hecho': datos_especificos.get('texto_hecho', 'Presentación de información falsa')}
        res_bi = calcular_beneficio_ilicito(datos_bi_base)
        if not res_bi or res_bi.get('error'): return res_bi or {'error': 'Error BI.'}
        bi_uit = res_bi.get('beneficio_ilicito_uit', 0)
        res_multa = calcular_multa({**datos_comunes, 'beneficio_ilicito': bi_uit})
        multa_uit = res_multa.get('multa_final_uit', 0)

        # 4. Generar Tablas para Cuerpo (CE, BI, Multa, Personal)
        # (Lógica de tablas copiada de INF003 para consistencia)
        # Tabla CE
        ce_fmt = [{**item, 'precio_dolares': f"US$ {item.get('precio_dolares', 0):,.3f}", 'precio_soles': f"S/ {item.get('precio_soles', 0):,.3f}", 'factor_ajuste': f"{item.get('factor_ajuste', 0):,.3f}", 'monto_soles': f"S/ {item.get('monto_soles', 0):,.3f}", 'monto_dolares': f"US$ {item.get('monto_dolares', 0):,.3f}"} for item in ce_data_raw]
        ce_fmt.append({'descripcion': 'Total', 'monto_soles': f"S/ {total_soles:,.3f}", 'monto_dolares': f"US$ {total_dolares:,.3f}"})
        tabla_ce = create_table_subdoc(doc_tpl_bi, ["Descripción", "Precio (US$)", "Precio (S/)", "Factor de ajuste", "Monto (S/)", "Monto (US$)"], ce_fmt, ['descripcion', 'precio_dolares', 'precio_soles', 'factor_ajuste', 'monto_soles', 'monto_dolares'])
        # Tabla BI
        fn_map, fn_data = res_bi.get('footnote_mapping', {}), res_bi.get('footnote_data', {})
        fn_list = [f"({l}) {obtener_fuente_formateada(k, fn_data, id_infraccion, False)}" for l, k in sorted(fn_map.items())] # es_extemporaneo=False
        fn_data_dict = {'list': fn_list, 'elaboration': 'Elaboración: SSAG - DFAI.', 'style': 'FuenteTabla'}
        tabla_bi = create_main_table_subdoc(doc_tpl_bi, ["Descripción", "Monto"], res_bi.get('table_rows', []), keys=['descripcion', 'monto'], footnotes_data=fn_data_dict)
        # Tabla Multa
        tabla_multa = create_main_table_subdoc(doc_tpl_bi, ["Componentes", "Monto"], res_multa.get('multa_data_raw', []), ['Componentes', 'Monto'])

        # Tabla Personal (con enteros y fuente)
        tabla_pers_render = datos_especificos.get('tabla_personal', [])
        tabla_pers_sin_total = []
        for fila in tabla_pers_render:
            perfil = fila.get('Perfil')
            cantidad = pd.to_numeric(fila.get('Cantidad'), errors='coerce')
            if perfil and cantidad > 0: tabla_pers_sin_total.append({'Perfil': perfil, 'Descripción': fila.get('Descripción', ''), 'Cantidad': int(cantidad)})
        num_pers_total_int = int(datos_especificos.get('num_personal_capacitacion', 0))
        tabla_pers_data = tabla_pers_sin_total + [{'Perfil':'Total', 'Descripción':'', 'Cantidad': num_pers_total_int}]
        tabla_personal = create_personal_table_subdoc(doc_tpl_bi, ["Perfil", "Descripción", "Cantidad"], tabla_pers_data, ['Perfil', 'Descripción', 'Cantidad'], column_widths=(2,3,1), texto_posterior="Elaboración: Subdirección de Sanción y Gestión de Incentivos (SSAG) - DFAI.")

        # 5. Contexto y Renderizado Cuerpo
        num_hecho_actual = datos_comunes['numero_hecho_actual']
        contexto_word = {
            **datos_comunes['context_data'], 'acronyms': datos_comunes['acronym_manager'],
            'hecho': {'numero_imputado': num_hecho_actual, 'descripcion': RichText(datos_especificos.get('texto_hecho', '')), 'tabla_ce': tabla_ce, 'tabla_bi': tabla_bi, 'tabla_multa': tabla_multa},
            'numeral_hecho': f"IV.{num_hecho_actual}",
            'mh_uit': f"{multa_uit:,.3f} UIT", 'bi_uit': f"{bi_uit:,.3f} UIT",
            'fuente_cos': res_bi.get('fuente_cos', ''),
            # --- Lógica de texto movida a app.py ---
            'texto_explicacion_prorrateo': '', 
            'tabla_detalle_personal': tabla_personal,
            # --- Placeholders para anexo CE (si la plantilla simple los usa) ---
             'nro_personal': texto_con_numero(num_pers_total_int, 'f'),
             'precio_dolares': f"US$ {res_ce.get('precio_dolares', 0):,.3f}",
             'fi_mes': res_ce.get('fi_mes', ''),
             'fi_ipc': f"{res_ce.get('fi_ipc', 0):,.3f}",
             'fi_tc': f"{res_ce.get('fi_tc', 0):,.3f}",
        }
        doc_tpl_bi.render(contexto_word, autoescape=True); buf_final_hecho = io.BytesIO(); doc_tpl_bi.save(buf_final_hecho)

        # 6. Generar Anexo CE (Simple)
        anexos_ce = []
        # Recrear tabla CE para anexo (idéntica a la del cuerpo)
        tabla_ce_anx = create_table_subdoc(tpl_anexo, ["Descripción", "Precio (US$)", "Precio (S/)", "Factor de ajuste", "Monto (S/)", "Monto (US$)"], ce_fmt, ['descripcion', 'precio_dolares', 'precio_soles', 'factor_ajuste', 'monto_soles', 'monto_dolares'])
        contexto_anx = {**contexto_word, 'tabla_ce_anexo': tabla_ce_anx} # Usar contexto base + tabla
        tpl_anexo.render(contexto_anx, autoescape=True); buf_anexo_final = io.BytesIO(); tpl_anexo.save(buf_anexo_final); anexos_ce.append(buf_anexo_final)

        # 7. Devolver Resultados
        resultados_app = {
             'ce_data_raw': ce_data_raw, 
             'totales': {
                 'ce_total_soles': total_soles, 'ce_total_dolares': total_dolares, 
                 'beneficio_ilicito_uit': bi_uit, 'multa_final_uit': multa_uit, 
                 'bi_data_raw': res_bi.get('table_rows', []), 'multa_data_raw': res_multa.get('multa_data_raw', []), 
                 'tabla_personal_data': tabla_pers_data
            }
        }
        return {
            'contexto_final_word': contexto_word,
            'doc_pre_compuesto': buf_final_hecho,
            'resultados_para_app': resultados_app,
            'es_extemporaneo': False, 'usa_capacitacion': True,
            'anexos_ce_generados': anexos_ce,
            'ids_anexos': list(filter(None, res_ce.get('ids_anexos', []))),
            # --- Devolver para app.py ---
            'texto_explicacion_prorrateo': '', 
            'tabla_detalle_personal': tabla_personal,
            'tabla_personal_data': tabla_pers_data
        }
    except Exception as e:
        import traceback; traceback.print_exc(); return {'error': f"Error _procesar_simple INF006: {e}"}
    
# --- 6. PROCESAR HECHO MÚLTIPLE (Adaptado de INF003) ---
def _procesar_hecho_multiple(datos_comunes, datos_especificos):
    """
    Procesa INF006 con múltiples extremos (fechas), prorrateando costo base.
    """
    try:
        # a. Cargar plantillas compuestas
        df_tipificacion, id_infraccion = datos_comunes['df_tipificacion'], datos_comunes['id_infraccion']
        fila_inf = df_tipificacion[df_tipificacion['ID_Infraccion'] == id_infraccion].iloc[0]
        id_tpl_bi_ext, id_tpl_ce_ext = fila_inf.get('ID_Plantilla_BI_Extremo'), fila_inf.get('ID_Plantilla_CE_Extremo')
        if not id_tpl_bi_ext or not id_tpl_ce_ext: return {'error': f'Faltan IDs plantillas extremo (BI o CE) para {id_infraccion}.'}
        buf_bi_ext, buf_ce_ext = descargar_archivo_drive(id_tpl_bi_ext), descargar_archivo_drive(id_tpl_ce_ext)
        if not buf_bi_ext or not buf_ce_ext: return {'error': f'Fallo descarga plantillas extremo para {id_infraccion}.'}
        tpl_bi_final = DocxTemplate(buf_bi_ext)

        # b. Inicializar
        total_bi_uit = 0.0; lista_bi = []; lista_ce_raw = []
        anexos_ce = []; anexos_ids = set(); lista_ext_plantilla = []
        resultados_app = {'extremos': [], 'totales': {}}

        # c. Generar tabla personal (con enteros y fuente)
        tabla_pers_render = datos_especificos.get('tabla_personal', [])
        tabla_pers_sin_total = []
        for fila in tabla_pers_render:
            perfil = fila.get('Perfil')
            cantidad = pd.to_numeric(fila.get('Cantidad'), errors='coerce')
            if perfil and cantidad > 0: tabla_pers_sin_total.append({'Perfil': perfil, 'Descripción': fila.get('Descripción', ''), 'Cantidad': int(cantidad)})
        num_pers_total_int = int(datos_especificos.get('num_personal_capacitacion', 0))
        tabla_pers_data = tabla_pers_sin_total + [{'Perfil':'Total', 'Descripción':'', 'Cantidad': num_pers_total_int}]
        tabla_pers_subdoc = create_personal_table_subdoc(tpl_bi_final, ["Perfil", "Descripción", "Cantidad"], tabla_pers_data, ['Perfil', 'Descripción', 'Cantidad'], column_widths=(2,3,1), texto_posterior="Elaboración: Subdirección de Sanción y Gestión de Incentivos (SSAG) - DFAI.")

        # d. Lógica de Prorrateo
        grupos_anio = {}; costos_base_prorr = {}
        for ext in datos_especificos['extremos']: anio=ext['fecha_incumplimiento'].year; grupos_anio.setdefault(anio, []).append(ext)
        for anio, grupo in grupos_anio.items():
            fecha_ref = grupo[0]['fecha_incumplimiento']
            costo_total = calcular_costo_capacitacion(num_pers_total_int, {**datos_comunes, 'fecha_incumplimiento': fecha_ref})
            if costo_total.get('error'): return costo_total
            num_ext_grupo = len(grupo)
            costos_base_prorr[anio] = { "precio_soles": costo_total.get('precio_base_soles_con_igv', 0)/num_ext_grupo, "precio_dolares": costo_total.get('precio_base_dolares_con_igv', 0)/num_ext_grupo, "ipc_costeo": costo_total.get('ipc_costeo', 0), "descripcion": costo_total.get('descripcion', ''), "ids_anexos": costo_total.get('ids_anexos', []) }
            if costo_total.get('ids_anexos'): anexos_ids.update(costo_total.get('ids_anexos'))

        # e. Bucle principal: Calcular CE/BI por extremo, generar anexos
        texto_hecho_principal = datos_especificos.get('texto_hecho', 'Presentación de información falsa')
        num_hecho = datos_comunes['numero_hecho_actual']

        for j, extremo in enumerate(datos_especificos['extremos']):
            anio_ext = extremo['fecha_incumplimiento'].year
            costo_base = costos_base_prorr.get(anio_ext)
            if not costo_base: continue
            fecha_inc = extremo['fecha_incumplimiento']
            ipc_row = datos_comunes['df_indices'][datos_comunes['df_indices']['Indice_Mes'].dt.to_period('M') == pd.to_datetime(fecha_inc).to_period('M')]
            if ipc_row.empty: continue
            ipc_inc, tc_inc = ipc_row.iloc[0]['IPC_Mensual'], ipc_row.iloc[0]['TC_Mensual']
            if pd.isna(ipc_inc) or pd.isna(tc_inc) or tc_inc==0: continue
            factor = round(ipc_inc / costo_base['ipc_costeo'], 3) if costo_base['ipc_costeo'] > 0 else 0
            monto_s = costo_base['precio_soles'] * factor
            monto_d = monto_s / tc_inc if tc_inc > 0 else 0
            ce_raw_ext = [{"descripcion": costo_base['descripcion'], "precio_soles": costo_base['precio_soles'], "precio_dolares": costo_base['precio_dolares'], "factor_ajuste": factor, "monto_soles": monto_s, "monto_dolares": monto_d}]
            lista_ce_raw.extend(ce_raw_ext)

            # Generar Anexo CE del extremo
            tpl_anx_loop = DocxTemplate(io.BytesIO(buf_ce_ext.getvalue()))
            ce_fmt_anx = [{**item, 'monto_soles': f"S/ {item['monto_soles']:,.3f}", 'monto_dolares': f"US$ {item['monto_dolares']:,.3f}"} for item in ce_raw_ext]
            ce_fmt_anx.append({'descripcion':'Total Extremo', 'monto_soles':f"S/ {monto_s:,.3f}", 'monto_dolares':f"US$ {monto_d:,.3f}"})
            tabla_ce_anx = create_table_subdoc(tpl_anx_loop, ["Descripción", "Precio (US$)", "Precio (S/)", "Factor de ajuste", "Monto (S/)", "Monto (US$)"], ce_fmt_anx, ['descripcion', 'precio_dolares', 'precio_soles', 'factor_ajuste', 'monto_soles', 'monto_dolares'])
            desc_anexo = f"Presentación de Info. Falsa del {fecha_inc.strftime('%d/%m/%Y')} (Prorrateado)"
            contexto_anx = {
                **datos_comunes['context_data'], 'hecho': {'numero_imputado': num_hecho},
                'extremo': {'numeral': j+1, 'tipo': desc_anexo, 'tabla_ce': tabla_ce_anx},
                'fi_mes': format_date(fecha_inc, "MMMM 'de' yyyy", locale='es'), 'fi_ipc': f"{ipc_inc:,.3f}", 'fi_tc': f"{tc_inc:,.3f}",
                'nro_personal': texto_con_numero(num_pers_total_int, 'f'), 
                'precio_dolares': f"US$ {costo_base.get('precio_dolares', 0):,.3f}"
            }
            tpl_anx_loop.render(contexto_anx, autoescape=True); buf_anx = io.BytesIO(); tpl_anx_loop.save(buf_anx); anexos_ce.append(buf_anx)

            # Calcular BI del extremo
            texto_para_bi = f"{texto_hecho_principal} [Extremo {j + 1}]"
            datos_bi_ext = {**datos_comunes, 'fecha_incumplimiento': fecha_inc, 'ce_soles': monto_s, 'ce_dolares': monto_d, 'texto_del_hecho': texto_para_bi}
            res_bi_ext = calcular_beneficio_ilicito(datos_bi_ext)
            if res_bi_ext.get('error'): continue
            total_bi_uit += res_bi_ext.get('beneficio_ilicito_uit', 0)
            lista_bi.append(res_bi_ext)
            resultados_app['extremos'].append({'tipo': desc_anexo, 'ce_data': ce_raw_ext, 'bi_data': res_bi_ext.get('table_rows', []), 'bi_uit': res_bi_ext.get('beneficio_ilicito_uit', 0)})

        # f. Post-Cálculo: Multa, Tablas Consolidadas BI
        if not lista_bi: return {'error': 'No se pudo calcular BI para ningún extremo.'}
        res_multa = calcular_multa({**datos_comunes, 'beneficio_ilicito': total_bi_uit})
        multa_uit = res_multa.get('multa_final_uit', 0)

        # Lógica remapeo de fuentes (idéntica a INF003)
        notas_map = {}; map_clave_txt = {}; datos_gen_notas = lista_bi[0].get('footnote_data', {})
        for i, res_bi in enumerate(lista_bi):
            datos_ext_notas = res_bi.get('footnote_data', {})
            for letra_orig, clave_orig in res_bi.get('footnote_mapping', {}).items():
                datos_fmt = {**datos_gen_notas, **datos_ext_notas}
                txt_nota = obtener_fuente_formateada(clave_orig, datos_fmt, id_infraccion, False)
                if txt_nota not in notas_map: notas_map[txt_nota] = set()
                if not txt_nota.startswith("Error:"): notas_map[txt_nota].add(clave_orig)
                map_clave_txt[(clave_orig, i)] = txt_nota
        key_order = ['ce_anexo', 'cok', 'periodo_bi', 'bcrp', 'ipc_fecha', 'sunat']
        map_txt_letra = {}; letra_code = ord('a'); txt_mapeados = set()
        for clave in key_order:
            txts_clave = sorted(list(set(txt for (k, idx), txt in map_clave_txt.items() if k == clave)))
            for txt in txts_clave:
                if txt not in txt_mapeados: map_txt_letra[txt] = chr(letra_code); txt_mapeados.add(txt); letra_code += 1
        fn_list_final = []; txt_agregados = set()
        for clave in key_order:
            txts_clave = sorted(list(set(txt for (k, idx), txt in map_clave_txt.items() if k == clave)))
            for txt in txts_clave:
                if txt in map_txt_letra and txt not in txt_agregados: letra = map_txt_letra[txt]; fn_list_final.append(f"({letra}) {txt}"); txt_agregados.add(txt)
        fn_data_final = {'list': fn_list_final, 'elaboration': 'Elaboración: SSAG - DFAI.', 'style': 'FuenteTabla'}

        # Crear tablas consolidadas BI y Multa
        tabla_bi_cons = create_consolidated_bi_table_subdoc(tpl_bi_final, lista_bi, total_bi_uit, footnotes_data=fn_data_final, map_texto_a_letra=map_txt_letra, map_clave_a_texto=map_clave_txt)
        tabla_multa_final = create_main_table_subdoc(tpl_bi_final, ["Componentes", "Monto"], res_multa.get('multa_data_raw', []), ['Componentes', 'Monto'])

        # g. Contexto Final y Renderizado
        contexto_final = {
            **datos_comunes['context_data'],
            'acronyms': datos_comunes['acronym_manager'],
            'hecho': { 'numero_imputado': num_hecho, 'descripcion': RichText(datos_especificos.get('texto_hecho', '')),
                       'tabla_bi': tabla_bi_cons, 'tabla_multa': tabla_multa_final },
            'numeral_hecho': f"IV.{num_hecho}",
            'fuente_cos': lista_bi[0].get('fuente_cos', ''),
            'mh_uit': f"{multa_uit:,.3f} UIT", 'bi_uit': f"{total_bi_uit:,.3f} UIT",
            # --- Lógica de texto movida a app.py ---
            'texto_explicacion_prorrateo': '', 
            'tabla_detalle_personal': tabla_pers_subdoc,
        }

        # Preparar datos para App (consolidado BI)
        filas_bi_app = []
        if lista_bi:
            # ... (Lógica idéntica a INF003 para construir filas_bi_app) ...
            primer_resultado_bi = lista_bi[0]
            cos_anual_row = next((row for row in primer_resultado_bi['table_rows'] if 'COS (anual)' in row['descripcion']), None)
            cosm_row = next((row for row in primer_resultado_bi['table_rows'] if 'COSm (mensual)' in row['descripcion']), None)
            tc_row = next((row for row in primer_resultado_bi['table_rows'] if 'Tipo de cambio' in row['descripcion']), None)
            uit_row = next((row for row in primer_resultado_bi['table_rows'] if 'Unidad Impositiva' in row['descripcion']), None)
            costos_ajustados_s = []
            for i, res in enumerate(lista_bi): 
                ce_row = res['table_rows'][0]; filas_bi_app.append({'descripcion': ce_row['descripcion'], 'monto': ce_row['monto'], 'ref': ce_row.get('ref')})
                aj_row = next((row for row in res['table_rows'] if 'Costo evitado ajustado' in row['descripcion']), None);
                if aj_row: monto_str = aj_row['monto'].replace('S/','').replace('US$','').replace(',','').strip(); costos_ajustados_s.append(float(monto_str))
            if cos_anual_row: filas_bi_app.append(cos_anual_row)
            if cosm_row: filas_bi_app.append(cosm_row)
            for i, res in enumerate(lista_bi): 
                t_row = next((row for row in res['table_rows'] if 'T: meses' in row['descripcion']), None);
                if t_row: filas_bi_app.append({'descripcion': f"{t_row['descripcion']} [Extremo {i+1}]", 'monto': t_row['monto'], 'ref': t_row.get('ref')})
            total_aj_s = sum(costos_ajustados_s)
            filas_bi_app.append({'descripcion': 'Costo evitado ajustado total (S/)', 'monto': f"S/ {total_aj_s:,.3f}", 'ref': None})
            if tc_row: filas_bi_app.append(tc_row)
            filas_bi_app.append({'descripcion': 'Beneficio ilícito (S/)', 'monto': f"S/ {total_aj_s:,.3f}", 'ref': None})
            if uit_row: filas_bi_app.append(uit_row)
            filas_bi_app.append({'descripcion': 'Beneficio Ilícito (UIT)', 'monto': f"{total_bi_uit:,.3f} UIT", 'ref': None})

        resultados_app['totales'] = { 'beneficio_ilicito_uit': total_bi_uit, 'multa_data_raw': res_multa.get('multa_data_raw', []), 'multa_final_uit': multa_uit, 'bi_data_raw': filas_bi_app, 'tabla_personal_data': tabla_pers_data }

        # Renderizar plantilla final y devolver
        tpl_bi_final.render(contexto_final, autoescape=True); buf_final = io.BytesIO(); tpl_bi_final.save(buf_final)
        return {
            'doc_pre_compuesto': buf_final, 
            'resultados_para_app': resultados_app,
            'texto_explicacion_prorrateo': '', # Se genera en app.py
            'tabla_detalle_personal': tabla_pers_subdoc, # app.py lo usará
            'usa_capacitacion': True, 'es_extemporaneo': False,
            'anexos_ce_generados': anexos_ce,
            'ids_anexos': list(filter(None, anexos_ids))
        }
    except Exception as e:
        import traceback; traceback.print_exc(); return {'error': f"Error _procesar_multiple INF006: {e}"}