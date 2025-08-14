# --- 1. IMPORTACIONES ---
from docxtpl import RichText
from num2words import num2words
from datetime import date, timedelta
import streamlit as st
import holidays
import pandas as pd

# Importa las funciones constructoras de tablas
from funciones import create_table_subdoc, create_main_table_subdoc
from sheets import calcular_costo_evitado, calcular_beneficio_ilicito

# --- 2. FUNCIÓN "CONTRATO" ---  

def preparar_contexto_especifico(doc_tpl, datos_hecho, datos_generales):
    """
    Prepara y devuelve el diccionario de contexto COMPLETO y específico
    para la infracción tipo INF004.
    """
    print("-> Cargando lógica y contexto específico de INF004...")
    
    # --- 3. CREACIÓN DE LAS TABLAS PARA ESTA INFRACCIÓN ---

    # --- Tabla Costo Evitado (CE) ---
    tabla_ce_subdoc = None
    ce_data_cruda = datos_hecho.get('ce_data_raw', [])
    if ce_data_cruda:
        ce_table_formatted = []
        for item in ce_data_cruda:
            ce_table_formatted.append({
                'descripcion': item['descripcion'],
                'cantidad': f"{int(float(item.get('cantidad', 0)))}",
                'horas': f"{int(float(item.get('horas', 0)))}",
                'precio_soles': f"{item.get('precio_soles', 0):,.3f}",
                'factor_ajuste': f"{item.get('factor_ajuste', 0):,.3f}",
                'monto_soles': f"S/ {item.get('monto_soles', 0):,.3f}",
                'monto_dolares': f"US$ {item.get('monto_dolares', 0):,.3f}"
            })
        total_s = datos_hecho.get('ce_total_soles', 0)
        total_d = datos_hecho.get('ce_total_dolares', 0)
        ce_table_formatted.append({
            'descripcion': 'Total', 'cantidad': '', 'horas': '', 'precio_soles': '',
            'factor_ajuste': '', 'monto_soles': f"S/ {total_s:,.3f}",
            'monto_dolares': f"US$ {total_d:,.3f}"
        })
        tabla_ce_subdoc = create_table_subdoc(
            doc_tpl,
            ["Descripción", "Cantidad", "Horas", "Precio (S/)", "Factor de ajuste",
             "Monto(*) (S/)", "Monto(*) (US$)"],
            ce_table_formatted,
            ['descripcion', 'cantidad', 'horas', 'precio_soles', 'factor_ajuste',
             'monto_soles', 'monto_dolares']
        )

    # --- Tabla Beneficio Ilícito (BI) ---
    tabla_bi_subdoc = None
    bi_data_cruda = datos_hecho.get('bi_data_raw', [])
    if bi_data_cruda:
        bi_table_formatted = [{'descripcion': item.get('Descripción', ''), 'monto': item.get('Monto', '')} for item in bi_data_cruda]
        tabla_bi_subdoc = create_main_table_subdoc(
            doc_tpl, ["Descripción", "Monto"], bi_table_formatted, ['descripcion', 'monto']
        )

    # --- Tabla Multa ---
    tabla_multa_subdoc = None
    multa_data_cruda = datos_hecho.get('multa_data_raw', [])
    if multa_data_cruda:
        multa_table_formatted = [{'componente': item.get('Componentes', ''), 'monto': item.get('Monto', '')} for item in multa_data_cruda]
        tabla_multa_subdoc = create_main_table_subdoc(
            doc_tpl, ["Componentes", "Monto"], multa_table_formatted, ['componente', 'monto']
        )

    lista_sustentos = datos_hecho.get('lista_sustentos', [])
    texto_sustentos_rt = RichText()
    for sustento in lista_sustentos:
        texto_sustentos_rt.add(f'- {sustento}\n')

    # Primero trato de obtenerlo desde datos_hecho (que es específico de este hecho)
    dias_habiles = datos_hecho.get('dias_habiles', None)

    # Si no existe, uso el valor global de session_state (por compatibilidad)
    if dias_habiles is None:
        dias_habiles = st.session_state.get('dias_habiles_plazo', 0)

    # --- 4. AÑADIR PLACEHOLDERS DE TEXTO ADICIONALES ---

    horas_totales = dias_habiles * 8
    horas_numero = int(horas_totales)
    horas_texto = num2words(horas_numero, lang='es')
    dias_calculados = horas_totales / 8
    if dias_calculados == int(dias_calculados):
        horas_dias = f"{int(dias_calculados)}"
    else:
        horas_dias = f"{dias_calculados:,.3f}"

    # Lógica para placeholders de fechas y BI
    fecha_inc_hecho = datos_hecho.get('fecha_incumplimiento')
    if fecha_inc_hecho:
        fecha_incumplimiento_texto = fecha_inc_hecho.strftime('%d de %B de %Y').lower()
    else:
        fecha_incumplimiento_texto = "No aplica"
    
    # --- Placeholders de fuentes y costos (priorizar datos_hecho, luego session_state) ---
    fuente_cos = datos_hecho.get('fuente_cos') or st.session_state.get('fuente_cos', '')
    fuente_salario = datos_hecho.get('fuente_salario') or st.session_state.get('fuente_salario', '')
    pdf_salario = datos_hecho.get('pdf_salario') or st.session_state.get('pdf_salario', '')
    fc_texto = datos_hecho.get('fc_texto') or st.session_state.get('fc_texto', '')
    prov_cotizacion = datos_hecho.get('prov_cotizacion') or st.session_state.get('prov_cotizacion', '')
    incluye_igv = datos_hecho.get('incluye_igv') or st.session_state.get('incluye_igv', '')
    fi_mes = datos_hecho.get('fi_mes') or st.session_state.get('fi_mes', '')
    fi_ipc = datos_hecho.get('fi_ipc') or st.session_state.get('fi_ipc', '')
    fi_tc = datos_hecho.get('fi_tc') or st.session_state.get('fi_tc', '')
    fc1 = datos_hecho.get('fc1') or st.session_state.get('fc1', '')
    ipc_fc1 = datos_hecho.get('ipc_fc1') or st.session_state.get('ipc_fc1', '')
    fc2 = datos_hecho.get('fc2') or st.session_state.get('fc2', '')
    ipc_fc2 = datos_hecho.get('ipc_fc2') or st.session_state.get('ipc_fc2', '')


    # --- 5. CONSTRUCCIÓN Y RETORNO DEL DICCIONARIO DE CONTEXTO ---
    
    # El diccionario anidado para los placeholders que usan 'hecho.'
    datos_para_hecho = {
        'numero_imputado': datos_generales['numero_hecho_actual'],
        'descripcion': RichText(datos_hecho.get('texto_hecho', '')),
        'tabla_ce': tabla_ce_subdoc,
        'tabla_bi': tabla_bi_subdoc,
        'tabla_multa': tabla_multa_subdoc,
        'texto_sustentos': texto_sustentos_rt,
        # ...etc...
    }
    
    # --- NUEVO: Crear un subdocumento para el resumen de fuentes ---
    resumen_texto_plano = datos_hecho.get('resumen_fuentes_costo', '')
    lineas_resumen = resumen_texto_plano.split('\n')
    
    # Se crea un subdocumento vacío
    subdoc_resumen = doc_tpl.new_subdoc()
    # Se añade cada línea como un párrafo separado
    for linea in lineas_resumen:
        subdoc_resumen.add_paragraph(linea)
    # --- FIN DEL NUEVO CÓDIGO ---

    fecha_inc_obj = datos_hecho.get('fecha_incumplimiento')
    if fecha_inc_obj:
        fecha_incumplimiento_texto = fecha_inc_obj.strftime('%d de %B de %Y').lower()
    else:
        fecha_incumplimiento_texto = "No aplica"

    # El diccionario final con TODOS los placeholders
    contexto_final = {
        # Pasa TODOS los datos generales (expediente, nombres, etc.)
        **datos_generales['context_data'],
        
        # Pasa TODOS los datos específicos del hecho que ya calculamos
        # (incluye fi_mes, fi_ipc, fuente_coti, resumen_fuentes_costo, etc.)
        **datos_hecho,

        # Asigna el diccionario que creamos a la clave 'hecho'
        'hecho': datos_para_hecho,
        
        # Placeholders que se construyen aquí mismo
        'tabla_ce': tabla_ce_subdoc,
        'tabla_bi': tabla_bi_subdoc,
        'tabla_multa': tabla_multa_subdoc,
        'texto_sustentos': texto_sustentos_rt,
        'horas_texto': horas_texto,
        'horas_numero': horas_numero,
        'horas_dias': horas_dias,
        'fecha_incumplimiento_texto': fecha_incumplimiento_texto,
        'mes_hoy': date.today().strftime('%B %Y').lower()
    }

    return contexto_final