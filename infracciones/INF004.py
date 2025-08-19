import streamlit as st
import pandas as pd
import io
from num2words import num2words
from docxtpl import DocxTemplate, RichText
from datetime import date, timedelta
import holidays
from textos_manager import obtener_fuente_formateada
from funciones import create_main_table_subdoc, create_table_subdoc
from sheets import calcular_beneficio_ilicito, calcular_multa, descargar_archivo_drive, calcular_beneficio_ilicito_extemporaneo

# --- CÁLCULO DEL COSTO EVITADO PARA INF004 ---
def _calcular_costo_evitado_inf004(datos_comunes, datos_especificos):
    """
    Calcula el Costo Evitado para INF004, restaurando la lógica de cálculo
    original y completa para encontrar el costo más adecuado.
    """
    try:
        # 1. Desempaquetado de datos
        df_items_infracciones = datos_comunes['df_items_infracciones']
        df_costos_items = datos_comunes['df_costos_items']
        df_coti_general = datos_comunes['df_coti_general']
        df_salarios_general = datos_comunes['df_salarios_general']
        df_indices = datos_comunes['df_indices']
        id_rubro = datos_comunes.get('id_rubro_seleccionado')
        id_infraccion = datos_comunes['id_infraccion']
        fecha_incumplimiento = datos_especificos['fecha_incumplimiento']
        dias_habiles = datos_especificos.get('dias_habiles_plazo', 0)

        # 2. Preparación de datos base
        fecha_incumplimiento_dt = pd.to_datetime(fecha_incumplimiento)
        ipc_row_inc = df_indices[df_indices['Indice_Mes'].dt.to_period('M') == fecha_incumplimiento_dt.to_period('M')]
        ipc_incumplimiento = ipc_row_inc.iloc[0]['IPC_Mensual']
        tipo_cambio_incumplimiento = ipc_row_inc.iloc[0]['TC_Mensual']

        items_calculados_final = []
        ids_anexos_final = []
        sustentos_final = []
        lineas_resumen_fuentes = []
        fuente_salario_final, pdf_salario_final = '', ''
        sustentos_de_cotizaciones = []
        salario_info_capturado = False

        receta_df = df_items_infracciones[df_items_infracciones['ID_Infraccion'] == id_infraccion]

        # 3. Bucle principal sobre la "receta"
        for _, item_receta in receta_df.iterrows():
            id_item_a_buscar = item_receta['ID_Item_Infraccion']
            posibles_costos = df_costos_items[df_costos_items['ID_Item_Infraccion'] == id_item_a_buscar].copy()
            if posibles_costos.empty: continue

            # Lógica de filtrado por Tipo de Ítem (Fijo/Variable)
            tipo_item_receta = item_receta.get('Tipo_Item')
            df_candidatos = pd.DataFrame()
            if tipo_item_receta == 'Variable':
                df_candidatos = posibles_costos[posibles_costos['ID_Rubro'] == id_rubro].copy()
            elif tipo_item_receta == 'Fijo':
                df_candidatos = posibles_costos.copy()
            if df_candidatos.empty: continue

            # Lógica para encontrar la fecha de costeo más cercana
            fechas_fuente = []
            for _, candidato in df_candidatos.iterrows():
                id_general = candidato['ID_General']
                fecha_fuente = pd.NaT
                if pd.notna(id_general):
                    if 'SAL' in id_general:
                        fuente = df_salarios_general[df_salarios_general['ID_Salario'] == id_general]
                        if not fuente.empty:
                            anio = fuente.iloc[0]['Costeo_Salario']
                            fecha_fuente = pd.to_datetime(f'{int(anio)}-12-31')
                    elif 'COT' in id_general:
                        fuente = df_coti_general[df_coti_general['ID_Cotizacion'] == id_general]
                        if not fuente.empty:
                            fecha_fuente = fuente.iloc[0]['Fecha_Costeo']
                fechas_fuente.append(fecha_fuente)
            
            df_candidatos['Fecha_Fuente'] = fechas_fuente
            df_candidatos.dropna(subset=['Fecha_Fuente'], inplace=True)
            if df_candidatos.empty: continue

            df_candidatos['Diferencia_Dias'] = (df_candidatos['Fecha_Fuente'] - fecha_incumplimiento_dt).dt.days.abs()
            fila_costo_final = df_candidatos.loc[df_candidatos['Diferencia_Dias'].idxmin()]
            # --- FIN DE LA LÓGICA ---
            id_anexo_item = fila_costo_final.get('ID_Anexo_Drive')
            if pd.notna(id_anexo_item):
                ids_anexos_final.append(id_anexo_item)

            # Lógica de índices (IPC y TC) para la fecha de costeo
            id_general = fila_costo_final['ID_General']
            fecha_fuente_dt = fila_costo_final['Fecha_Fuente']
            ipc_costeo = 0
            tc_costeo = 0

            if pd.notna(id_general) and 'SAL' in id_general:
                anio_costeo = fecha_fuente_dt.year
                indices_del_anio = df_indices[df_indices['Indice_Mes'].dt.year == anio_costeo]
                if not indices_del_anio.empty:
                    ipc_costeo = indices_del_anio['IPC_Mensual'].mean()
                    tc_costeo = indices_del_anio['TC_Mensual'].mean()
            elif pd.notna(id_general) and 'COT' in id_general:
                ipc_costeo_row = df_indices[df_indices['Indice_Mes'].dt.to_period('M') == fecha_fuente_dt.to_period('M')]
                if not ipc_costeo_row.empty:
                    ipc_costeo = ipc_costeo_row.iloc[0]['IPC_Mensual']
                    tc_costeo = ipc_costeo_row.iloc[0]['TC_Mensual']

            if ipc_costeo == 0: continue

            # --- INICIO: Captura de datos para placeholders ---
            if pd.notna(fila_costo_final.get('Sustento_Item')):
                sustentos_final.append(fila_costo_final['Sustento_Item'])
            if pd.notna(fila_costo_final.get('ID_Anexo_Drive')):
                ids_anexos_final.append(fila_costo_final['ID_Anexo_Drive'])

            id_general = fila_costo_final.get('ID_General')
            if id_general:
                if 'COT' in id_general:
                    fuente_row = df_coti_general[df_coti_general['ID_Cotizacion'] == id_general]
                    if not fuente_row.empty:
                        sustento = fuente_row.iloc[0].get('Sustento_Cotizacion')
                        if sustento: sustentos_de_cotizaciones.append(sustento)
                elif 'SAL' in id_general and not salario_info_capturado:
                    fuente_row = df_salarios_general[df_salarios_general['ID_Salario'] == id_general]
                    if not fuente_row.empty:
                        fuente_salario_final = fuente_row.iloc[0].get('Fuente_Salario', '')
                        pdf_salario_final = fuente_row.iloc[0].get('PDF_Salario', '')
                        salario_info_capturado = True
            
            # Construcción del resumen de fuentes para este ítem
            descripcion_item = fila_costo_final.get('Descripcion_Item', 'Ítem no especificado')
            fc_texto = fila_costo_final['Fecha_Fuente'].strftime('%B %Y').lower()
            ipc_fc_row = df_indices[df_indices['Indice_Mes'].dt.to_period('M') == fila_costo_final['Fecha_Fuente'].to_period('M')]
            ipc_fc_texto = f"{ipc_fc_row.iloc[0]['IPC_Mensual']:,.3f}" if not ipc_fc_row.empty else "N/A"
            lineas_resumen_fuentes.append(f"{descripcion_item}: {fc_texto}, IPC={ipc_fc_texto}")
            # --- FIN: Captura de datos ---

            # 4. Cálculo de montos para el ítem seleccionado
            costo_original = float(fila_costo_final['Costo_Unitario_Item'])
            moneda_original = fila_costo_final['Moneda_Item']
            precio_base_soles = costo_original if moneda_original == 'S/' else costo_original * tc_costeo
            factor_ajuste = round(ipc_incumplimiento / ipc_costeo, 3) if ipc_costeo > 0 else 0
            horas = dias_habiles * 8
            cantidad = float(item_receta.get('Cantidad_Recursos', 1))
            monto_soles = cantidad * horas * precio_base_soles * factor_ajuste
            monto_dolares = monto_soles / tipo_cambio_incumplimiento if tipo_cambio_incumplimiento > 0 else 0

            items_calculados_final.append({
                "descripcion": fila_costo_final.get('Descripcion_Item', 'N/A'),
                "cantidad": cantidad, "horas": horas, "precio_soles": precio_base_soles,
                "factor_ajuste": factor_ajuste, "monto_soles": monto_soles,
                "monto_dolares": monto_dolares
            })

        if not items_calculados_final:
            return {'error': "No se encontraron costos aplicables."}
        
        fuente_coti_texto = "\n".join([f"- {s}" for s in list(set(sustentos_de_cotizaciones))])
        resumen_final_texto = "\n".join(lineas_resumen_fuentes)

        # Devuelve la lista de IDs recolectada
        return {
    "items_calculados": items_calculados_final,
    "ids_anexos": list(set(ids_anexos_final)),
    "sustentos": list(set(sustentos_final)),
    "fuente_salario": fuente_salario_final,
    "pdf_salario": pdf_salario_final,
    "fuente_coti": fuente_coti_texto,
    "resumen_fuentes_costo": resumen_final_texto,
    "error": None
}

    except Exception as e:
        return {'error': f"Error crítico en el cálculo del CE: {e}"}

# --- FUNCIONES PÚBLICAS DEL MÓDULO ---

def renderizar_inputs_especificos(i):
    # ... (Esta función está bien, no necesita cambios)
    st.markdown("##### Detalles del Requerimiento")
    datos_especificos = {}

    col_req1, col_req2 = st.columns(2)
    with col_req1:
        fecha_solicitud = st.date_input("Fecha de solicitud", key=f"fecha_sol_{i}", format="DD/MM/YYYY")
        min_fecha_entrega = fecha_solicitud if fecha_solicitud else date.today()
        fecha_entrega = st.date_input("Fecha máxima de entrega", min_value=min_fecha_entrega, key=f"fecha_ent_{i}", format="DD/MM/YYYY")

        if fecha_solicitud and fecha_entrega:
            if fecha_entrega < fecha_solicitud:
                st.error("La fecha de entrega no puede ser anterior a la solicitud.")
            else:
                feriados_pe = holidays.PE()
                rango_dias = pd.date_range(start=fecha_solicitud, end=fecha_entrega)
                dias_habiles = sum(1 for dia in rango_dias[1:] if dia.weekday() < 5 and dia not in feriados_pe)
                
                fecha_incumplimiento = fecha_entrega
                while True:
                    fecha_incumplimiento += timedelta(days=1)
                    if fecha_incumplimiento.weekday() < 5 and fecha_incumplimiento not in feriados_pe:
                        break
                
                datos_especificos['dias_habiles_plazo'] = dias_habiles
                datos_especificos['fecha_incumplimiento'] = fecha_incumplimiento
                
                st.metric(label="Días Hábiles de Plazo", value=dias_habiles)
                st.info(f"Fecha de Incumplimiento: **{fecha_incumplimiento.strftime('%d/%m/%Y')}**")

    with col_req2:
        num_items = st.number_input("Número de requerimientos de información", min_value=1, step=1, key=f"num_requerimientos_{i}")
        datos_especificos['num_items_solicitados'] = num_items
        
        estado_remision = st.radio(
            "Estado de la remisión:",
            options=["No remitió la información", "Remitió completo pero tardío", "Remitió parcial", "Remitió parcial pero tardío"],
            index=None, key=f"estado_remision_{i}"
        )
        datos_especificos['estado_entrega'] = estado_remision

        # --- INICIO DE LA CORRECCIÓN ---
        # El input para la fecha extemporánea AHORA ESTÁ DENTRO de la columna
        if estado_remision == "Remitió completo pero tardío":
            fecha_extemporanea = st.date_input(
                "Fecha de cumplimiento extemporáneo",
                min_value=fecha_entrega,
                key=f"fecha_ext_{i}",
                format="DD/MM/YYYY"
            )
            datos_especificos['fecha_cumplimiento_extemporaneo'] = fecha_extemporanea
                        # --- DEBUG 1: VERIFICA EL DATO AQUÍ ---
            st.warning(f"DEBUG UI: Fecha extemporánea capturada: {fecha_extemporanea}")
        # --- FIN DE LA CORRECCIÓN ---
    
        # --- INICIO DE LA MODIFICACIÓN ---
    # Añadimos el cargador de archivos aquí
    st.divider()
    datos_especificos['doc_adjunto_hecho'] = st.file_uploader(
        "Adjuntar archivo de análisis del hecho (Word .docx)",
        type=['docx'],
        key=f"upload_analisis_{i}"
    )
    # --- FIN DE LA MODIFICACIÓN ---
    
    datos_especificos['fecha_solicitud'] = fecha_solicitud
    datos_especificos['fecha_max_entrega'] = fecha_entrega
    
    return datos_especificos


def validar_inputs(datos_especificos):
    """
    Valida que TODOS los inputs específicos de INF004 estén completos.
    """
    # 1. Validar que las fechas principales estén seleccionadas
    if not datos_especificos.get('fecha_solicitud'):
        return False
    if not datos_especificos.get('fecha_max_entrega'):
        return False

    # 2. Validar que el número de requerimientos sea mayor que cero
    if not datos_especificos.get('num_items_solicitados'):
        return False

    # 3. Validar que se haya seleccionado una opción en el estado de remisión
    if not datos_especificos.get('estado_entrega'):
        return False

    # 4. Validación condicional: si es tardío, debe tener la fecha extemporánea
    if datos_especificos.get('estado_entrega') == "Remitió completo pero tardío":
        if not datos_especificos.get('fecha_cumplimiento_extemporaneo'):
            return False
            
    # Si todas las validaciones anteriores pasan, todo está OK
    return True


def procesar_infraccion(datos_comunes, datos_especificos):
    # --- FUNCIÓN CORREGIDA ---
    res_ce = _calcular_costo_evitado_inf004(datos_comunes, datos_especificos)
    if res_ce.get('error'): 
        return {'error': res_ce['error']}

    # Lee la clave correcta ("items_calculados") que devuelve la función de cálculo
    ce_data_raw = res_ce.get('items_calculados', []) 
    
    # Calcula los totales a partir de los datos crudos recibidos
    total_soles = sum(item.get('monto_soles', 0) for item in ce_data_raw)
    total_dolares = sum(item.get('monto_dolares', 0) for item in ce_data_raw)

    # --- INICIO DE LA MODIFICACIÓN ---
    # 2. Decidir qué función de Beneficio Ilícito usar
    estado_entrega = datos_especificos.get('estado_entrega')
    
    datos_bi_base = {**datos_comunes, 'ce_soles': total_soles, 'ce_dolares': total_dolares, 'fecha_incumplimiento': datos_especificos['fecha_incumplimiento']}
    
    if estado_entrega == "Remitió completo pero tardío":
        fecha_extemporanea = datos_especificos.get('fecha_cumplimiento_extemporaneo')
        
        # --- INICIO DEL BLOQUE DE SEGURIDAD A AÑADIR ---
        if not fecha_extemporanea:
            return {'error': "Para el estado 'Remitió completo pero tardío', es obligatorio seleccionar la 'Fecha de cumplimiento extemporáneo'."}
        # --- FIN DEL BLOQUE DE SEGURIDAD ---

        # A. Pre-cálculo para obtener los valores de COS
        pre_calculo_bi = calcular_beneficio_ilicito(datos_bi_base)
        if pre_calculo_bi.get('error'): return pre_calculo_bi

        # B. Enriquecer los datos con los resultados del pre-cálculo
        # (Esta parte ahora funcionará correctamente gracias al Paso 1)
        datos_bi_ext = {
            **datos_bi_base,
            'fecha_cumplimiento_extemporaneo': fecha_extemporanea,
            'cos_anual': pre_calculo_bi.get('cos_anual', 0),
            'cos_mensual': pre_calculo_bi.get('cos_mensual', 0),
            'moneda_cos': pre_calculo_bi.get('moneda_cos', 'S/'),
            'fuente_cos': pre_calculo_bi.get('fuente_cos', '')
        }
        
        # C. Llamar a la función extemporánea con los datos completos
        res_bi = calcular_beneficio_ilicito_extemporaneo(datos_bi_ext)
    else:
        res_bi = calcular_beneficio_ilicito(datos_bi_base)

    if res_bi.get('error'): return res_bi
    beneficio_ilicito_uit = res_bi.get('beneficio_ilicito_uit', 0)

    res_multa = calcular_multa({**datos_comunes, 'beneficio_ilicito': beneficio_ilicito_uit})
    multa_uit = res_multa.get('multa_final_uit', 0)

    doc_tpl = datos_comunes['doc_tpl']

    # --- INICIO: Preparación de placeholders para el contexto ---
    dias_habiles = datos_especificos.get('dias_habiles_plazo', 0)
    horas_numero = dias_habiles * 8
    horas_texto = num2words(horas_numero, lang='es')
    
    fecha_inc_dt = datos_especificos['fecha_incumplimiento']
    fi_mes = fecha_inc_dt.strftime('%B de %Y').lower()
    ipc_row_inc = datos_comunes['df_indices'][datos_comunes['df_indices']['Indice_Mes'].dt.to_period('M') == pd.to_datetime(fecha_inc_dt).to_period('M')]
    fi_ipc = f"{ipc_row_inc.iloc[0]['IPC_Mensual']:,.3f}" if not ipc_row_inc.empty else "N/A"
    fi_tc = f"{ipc_row_inc.iloc[0]['TC_Mensual']:,.3f}" if not ipc_row_inc.empty else "N/A"
    
    texto_sustentos_rt = RichText()
    for sustento in res_ce.get('sustentos', []):
        texto_sustentos_rt.add(f'- {sustento}\n')
    # --- FIN: Preparación ---

    # --- INICIO DE LA CORRECCIÓN ---
    # Tabla Costo Evitado (CE)
    tabla_ce_subdoc = None
    if ce_data_raw:
        # 1. Primero, formateamos los datos numéricos a texto para la tabla
        ce_table_formatted = []
        for item in ce_data_raw:
            ce_table_formatted.append({
                'descripcion': item.get('descripcion', ''),
                'cantidad': f"{item.get('cantidad', 0):.0f}",
                'horas': f"{item.get('horas', 0):.0f}",
                'precio_soles': f"S/ {item.get('precio_soles', 0):,.3f}",
                'factor_ajuste': f"{item.get('factor_ajuste', 0):,.3f}",
                'monto_soles': f"S/ {item.get('monto_soles', 0):,.3f}",
                'monto_dolares': f"US$ {item.get('monto_dolares', 0):,.3f}"
            })
        
        # 2. Añadimos la fila de Total con el formato correcto
        ce_table_formatted.append({
            'descripcion': 'Total', 'cantidad': '', 'horas': '', 'precio_soles': '',
            'factor_ajuste': '', 'monto_soles': f"S/ {total_soles:,.3f}",
            'monto_dolares': f"US$ {total_dolares:,.3f}"
        })
        
        # 3. Llamamos a la función con las cabeceras y claves correctas para INF004
        tabla_ce_subdoc = create_table_subdoc(
            doc_tpl,
            ["Descripción", "Cantidad", "Horas", "Precio (S/)", "Factor de ajuste", "Monto (S/)", "Monto (US$)"],
            ce_table_formatted,
            ['descripcion', 'cantidad', 'horas', 'precio_soles', 'factor_ajuste', 'monto_soles', 'monto_dolares']
        )
    
    # 1. Obtenemos los datos crudos y para formatear las fuentes
    filas_bi_crudas = res_bi.get('table_rows', [])
    datos_para_fuentes = res_bi.get('footnote_data', {})

    # 2. Asignamos letras a las fuentes (esto se mantiene igual)
    refs_usadas = [fila['ref'] for fila in filas_bi_crudas if fila.get('ref') and fila['ref'] not in locals().get('refs_usadas', [])]
    letras = 'abcdefghijklmnopqrstuvwxyz'
    mapa_ref_letra = {ref: letras[i] for i, ref in enumerate(refs_usadas)}

    # 3. Construimos las fuentes (esto se mantiene igual, pero SIN RichText)
    footnotes_bi_formateadas = []
    for ref_key in refs_usadas:
        texto_formateado = obtener_fuente_formateada(ref_key, datos_para_fuentes)
        # Creamos un texto simple, el formato se dará en la plantilla si es necesario
        footnotes_bi_formateadas.append(f"({mapa_ref_letra[ref_key]}) {texto_formateado}")
    
    # 4. Preparamos las filas de la tabla con el texto y el superíndice SEPARADOS
    filas_bi_para_tabla = []
    for fila_data in filas_bi_crudas:
        ref_key = fila_data.get('ref')
        letra = mapa_ref_letra.get(ref_key) # Obtenemos la letra (ej: 'a') o None
        
        # Formateamos el superíndice con paréntesis solo si existe la letra
        superindice_formateado = f"({letra})" if letra else ""

        filas_bi_para_tabla.append({
            'descripcion_texto': fila_data['descripcion'],
            'descripcion_superindice': superindice_formateado, # <-- Usamos la nueva variable
            'monto': fila_data['monto']
        })

    # 5. Creamos la tabla de BI (pasamos las claves nuevas)
    tabla_bi_subdoc = create_main_table_subdoc(
        datos_comunes['doc_tpl'],
        ["Descripción", "Monto"],
        filas_bi_para_tabla,
        # Las claves ahora son 'descripcion_texto' y 'monto'
        # El superíndice se manejará dentro de la función
        ['descripcion_texto', 'monto'] 
    )
    tabla_multa_subdoc = create_main_table_subdoc(doc_tpl, ["Componentes", "Monto"], res_multa.get('multa_data_raw', []), ['Componentes', 'Monto'])

    # -- 4. Ensamblaje del diccionario final para el hecho --
    datos_para_hecho = {
        'numero_imputado': datos_comunes['numero_hecho_actual'],
        'descripcion': RichText(datos_especificos.get('texto_hecho', '')),
        'tabla_ce': tabla_ce_subdoc,
        'tabla_bi': tabla_bi_subdoc,
        'bi_footnotes': footnotes_bi_formateadas, # <-- La nueva lista de fuentes
        'tabla_multa': tabla_multa_subdoc,
    }

    contexto_final = {
        **datos_comunes['context_data'],
        'hecho': datos_para_hecho,
        'mh_uit': f"{multa_uit:,.3f} UIT",
        'bi_uit': f"{beneficio_ilicito_uit:,.3f} UIT",
        # --- INICIO: Inclusión de placeholders en el contexto ---
        'horas_texto': horas_texto,
        'horas_numero': horas_numero,
        'horas_dias': dias_habiles, # Asumiendo que 'horas_dias' es lo mismo que 'dias_habiles'
        'fuente_cos': res_bi.get('fuente_cos', ''),
        'fecha_incumplimiento_texto': fecha_inc_dt.strftime('%d de %B de %Y').lower(),
        'fuente_salario': res_ce.get('fuente_salario', ''),
        'pdf_salario': res_ce.get('pdf_salario', ''),
        'fuente_coti': res_ce.get('fuente_coti', ''),
        'fi_mes': fi_mes,
        'fi_ipc': fi_ipc,
        'fi_tc': fi_tc,
        'resumen_fuentes_costo': res_ce.get('resumen_fuentes_costo', '')
        # --- FIN: Inclusión ---
    }

    # --- INICIO DE LA MODIFICACIÓN ---
    # 4. Generar el Anexo de Costo Evitado (AHORA se hace al final)
    anexos_ce_generados = []
    fila_infraccion = datos_comunes['df_tipificacion'][datos_comunes['df_tipificacion']['ID_Infraccion'] == datos_comunes['id_infraccion']]
    id_plantilla_anexo_ce = fila_infraccion.iloc[0].get('ID_Plantilla_CE')
    
    if id_plantilla_anexo_ce:
        buffer_anexo = descargar_archivo_drive(id_plantilla_anexo_ce)
        if buffer_anexo:
            anexo_tpl = DocxTemplate(buffer_anexo)
            # Ahora 'contexto_final' ya existe y puede ser usado aquí
            anexo_tpl.render(contexto_final)
            buffer_final_anexo = io.BytesIO()
            anexo_tpl.save(buffer_final_anexo)
            anexos_ce_generados.append(buffer_final_anexo)
    # --- FIN DE LA MODIFICACIÓN ---
    
    return {
        'contexto_final_word': contexto_final,
        'resultados_para_app': {
            'ce_data_raw': ce_data_raw, 'ce_total_soles': total_soles, 'ce_total_dolares': total_dolares,
            'bi_data_raw': res_bi.get('table_rows', []), 'beneficio_ilicito_uit': beneficio_ilicito_uit,
            'multa_data_raw': res_multa.get('multa_data_raw', []), 'multa_final_uit': multa_uit
        },
        'anexos_ce_generados': anexos_ce_generados, # <-- Devuelve el anexo Word generado
        'ids_anexos': res_ce.get('ids_anexos', []) # <-- Devuelve los IDs de sustento
    }
