import ssl
ssl._create_default_https_context = ssl._create_unverified_context
import streamlit as st
import io
import locale
from babel.dates import format_date 
from datetime import datetime, date
import pandas as pd
import importlib
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH                  
from docxcompose.composer import Composer
from docxtpl import DocxTemplate, RichText
from num2words import num2words
import requests # <--- AÑADIR
import traceback # <--- AÑADIR
from jinja2 import Environment
import copy #
from ifi import generar_documento_ifi

# Importaciones de nuestros módulos
from sheets import (conectar_gsheet, cargar_hoja_a_df, 
                    get_person_details_by_base_name, descargar_archivo_drive,
                    calcular_beneficio_ilicito, calcular_multa,
                    actualizar_hoja_con_df,
                    guardar_datos_caso, cargar_datos_caso) # <-- AÑADIR ESTAS DOS
from funciones import (combinar_con_composer, create_table_subdoc, 
                     create_main_table_subdoc, create_summary_table_subdoc, create_personal_table_subdoc, NumberingManager,
                     texto_con_numero, format_decimal_dinamico, get_initials_from_name, formatear_lista_hechos, redondeo_excel, create_capacitacion_table_subdoc)

def preparar_datos_para_json(obj):
    """
    Limpia recursivamente un objeto para que sea compatible con JSON.
    Convierte fechas a string y elimina objetos binarios o complejos.
    """
    if isinstance(obj, dict):
        # Eliminamos claves conocidas que causan errores de guardado (objetos binarios/clases)
        claves_invalidas = {'resultados', 'anexos_ce', 'tabla_detalle_personal', 'acronyms', 'numbering', 'doc_tpl'}
        return {k: preparar_datos_para_json(v) for k, v in obj.items() if k not in claves_invalidas}
    elif isinstance(obj, list):
        return [preparar_datos_para_json(i) for i in obj]
    elif isinstance(obj, (datetime, date)):
        return obj.isoformat() # "2024-05-20"
    elif isinstance(obj, (RichText, io.BytesIO)):
        return None # No se pueden guardar estos objetos
    return obj


def actualizar_y_recalcular_prorrateos(cliente_gspread, NOMBRE_GSHEET_MAESTRO):
    """
    Analiza todos los hechos calculados, detecta colisiones de capacitación por año,
    actualiza los factores de prorrateo y vuelve a procesar los resultados para que
    la interfaz muestre los montos actualizados (1/N).
    """
    if 'imputaciones_data' not in st.session_state:
        return

    # 1. Contar cuántos extremos hay por cada año (solo para los que usan capacitación)
    conteo_por_anio = {}
    for datos in st.session_state.imputaciones_data:
        id_inf = datos.get('id_infraccion', '')
        for extremo in datos.get('extremos', []):
            tipo = extremo.get('tipo_extremo', '') or extremo.get('tipo_presentacion', '')
            # Lógica: INF003 siempre cuenta, otros solo si es "No presentó"
            if 'INF003' in id_inf or tipo in ["No presentó", "No remitió", "No remitió información / Remitió incompleto"]:
                fecha = extremo.get('fecha_incumplimiento') or extremo.get('fecha_incumplimiento_extremo')
                if fecha:
                    conteo_por_anio[fecha.year] = conteo_por_anio.get(fecha.year, 0) + 1

    # 2. Actualizar factores y disparar recálculo de los hechos afectados
    for i, datos in enumerate(st.session_state.imputaciones_data):
        id_inf = datos.get('id_infraccion')
        if not id_inf or 'resultados' not in datos:
            continue

        mapa_nuevo = {}
        for extremo in datos.get('extremos', []):
            fecha = extremo.get('fecha_incumplimiento') or extremo.get('fecha_incumplimiento_extremo')
            if fecha and fecha.year in conteo_por_anio:
                mapa_nuevo[fecha.year] = 1.0 / conteo_por_anio[fecha.year]
        
        # Si el factor cambió, recalculamos este hecho individual
        if datos.get('mapa_factores_prorrateo') != mapa_nuevo:
            datos['mapa_factores_prorrateo'] = mapa_nuevo
            
            # --- RE-EJECUCIÓN DEL PROCESAMIENTO ---
            try:
                modulo = importlib.import_module(f"infracciones.{id_inf}")
                # Reconstruir datos comunes (similar al botón de calcular del Paso 3)
                df_tip = st.session_state.datos_calculo['df_tipificacion']
                id_plantilla = df_tip[df_tip['ID_Infraccion'] == id_inf].iloc[0]['ID_Plantilla_BI']
                
                datos_comunes = {
                    **st.session_state.datos_calculo,
                    'datos_hecho_completos': datos,
                    'fecha_emision_informe': st.session_state.get('fecha_emision_informe', date.today()),
                    'id_infraccion': id_inf,
                    'rubro': st.session_state.rubro_seleccionado,
                    'id_rubro_seleccionado': st.session_state.get('id_rubro_seleccionado'),
                    'numero_hecho_actual': i + 1,
                    'context_data': st.session_state.get('context_data', {}),
                    'acronym_manager': st.session_state.context_data.get('acronyms')
                }
                
                nuevos_res = modulo.procesar_infraccion(datos_comunes, datos)
                if not nuevos_res.get('error'):
                    st.session_state.imputaciones_data[i]['resultados'] = nuevos_res
            except Exception as e:
                st.error(f"Error al actualizar prorrateo del Hecho {i+1}: {e}")

# --- CONSTANTES: FACTORES DE GRADUACIÓN (f1 a f7) ---
FACTORES_GRADUACION = {
    "f1": {
        "titulo": "GRAVEDAD DEL DAÑO AL AMBIENTE",
        "criterios": {
            "1.1 Componentes Ambientales": {
                "No determinado / No aplica": 0.0,
                "El daño afecta a un (01) componente ambiental.": 0.10,
                "El daño afecta a dos (02) componentes ambientales.": 0.20,
                "El daño afecta a tres (03) componentes ambientales.": 0.30,
                "El daño afecta a cuatro (04) componentes ambientales.": 0.40,
                "El daño afecta a cinco (05) componentes ambientales.": 0.50
            },
            "1.2 Incidencia en la calidad": {
                "No determinado / No aplica": 0.0,
                "Impacto mínimo.": 0.06,
                "Impacto regular.": 0.12,
                "Impacto alto.": 0.18,
                "Impacto total.": 0.24
            },
            "1.3 Extensión geográfica": {
                "No determinado / No aplica": 0.0,
                "El impacto está localizado en el área de influencia directa.": 0.10,
                "El impacto está localizado en el área de influencia indirecta.": 0.20
            },
            "1.4 Reversibilidad/Recuperabilidad": {
                "No determinado / No aplica": 0.0,
                "Reversible en el corto plazo.": 0.06,
                "Recuperable en el corto plazo.": 0.12,
                "Recuperable en el mediano plazo.": 0.18,
                "Recuperable en el largo plazo o irrecuperable.": 0.24
            },
            "1.5 Afectación recursos/áreas protegidas": {
                "No existe afectación o esta es indeterminable...": 0.0,
                "El impacto se ha producido en un área natural protegida...": 0.40
            },
             "1.6 Afectación comunidades": {
                "No afecta a comunidades nativas o campesinas.": 0.0,
                "Afecta a una comunidad nativa o campesina.": 0.15,
                "Afecta a más de una comunidad nativa o campesina.": 0.30
            },
            "1.7 Afectación salud": {
                "No afecta a la salud de las personas...": 0.0,
                "Afecta la salud de las personas.": 0.60
            }
        }
    },
    "f2": {
        "titulo": "PERJUICIO ECONÓMICO CAUSADO (Pobreza)",
        "criterios": {
            "Incidencia de pobreza total": {
                "No determinado / No aplica": 0.0,
                "Incidencia de pobreza total hasta 19,6%.": 0.04,
                "Incidencia de pobreza total mayor a 19,6% hasta 39,1%.": 0.08,
                "Incidencia de pobreza total mayor a 39,1% hasta 58,7%.": 0.12,
                "Incidencia de pobreza total mayor a 58,7% hasta 78,2%.": 0.16,
                "Incidencia de pobreza total mayor a 78,2%.": 0.20
            }
        }
    },
    "f3": {
        "titulo": "ASPECTOS AMBIENTALES O FUENTES DE CONTAMINACIÓN",
        "criterios": {
             "Cantidad de aspectos": {
                "No determinado / No aplica": 0.0,
                "El impacto involucra un (01) aspecto ambiental...": 0.06,
                "El impacto involucra dos (02) aspectos ambientales...": 0.12,
                "El impacto involucra tres (03) aspectos ambientales...": 0.18,
                "El impacto involucra cuatro (04) aspectos ambientales...": 0.24,
                "El impacto involucra cinco (05) aspectos ambientales...": 0.30
             }
        }
    },
    "f4": {
        "titulo": "REINCIDENCIA",
        "criterios": {
            "Reincidencia": {
                "No existe reincidencia": 0.0,
                "Existe reincidencia (comisión de misma infracción en 1 año)": 0.20
            }
        }
    },
    "f5": {
        "titulo": "CORRECCIÓN DE LA CONDUCTA INFRACTORA (Atenuante)",
        "criterios": {
             "Subsanación/Corrección": {
                "No corrige / No aplica": 0.0,
                "Subsana voluntariamente antes del inicio del PAS (Eximente)": "Eximente",
                "Corrige (leve) a requerimiento, antes del inicio del PAS (Eximente)": "Eximente",
                "Corrige (trascendente) a requerimiento, antes del inicio del PAS (-40%)": -0.40,
                "Corrige luego del inicio del PAS, antes de resolución (-20%)": -0.20
             }
        }
    },
    "f6": {
        "titulo": "ADOPCIÓN DE MEDIDAS PARA REVERTIR CONSECUENCIAS",
        "criterios": {
            "Medidas adoptadas": {
                "No ejecutó ninguna medida (+30%)": 0.30,
                "Ejecutó medidas tardías (+20%)": 0.20,
                "Ejecutó medidas parciales (+10%)": 0.10,
                "No aplica / Neutro": 0.0,
                "Ejecutó medidas necesarias e inmediatas (-10%)": -0.10
            }
        }
    },
    "f7": {
        "titulo": "INTENCIONALIDAD",
        "criterios": {
             "Intencionalidad": {
                 "No se acredita intencionalidad": 0.0,
                 "Se acredita intencionalidad (+72%)": 0.72
             }
        }
    }
}

st.write(st.session_state)

# --- INICIALIZACIÓN DE LA APLICACIÓN ---
st.set_page_config(layout="wide", page_title="Asistente de Multas")
st.title("🤖 Asistente para la elaboración de informes de multa")

# --- INICIO: Lógica de Actualización BCRP ---
def actualizar_datos_bcrp(cliente_gspread):
    """
    Función principal para conectarse a la API del BCRP,
    descargar datos y actualizar la hoja de Google Sheets.
    """
    
    # 1. --- ¡DEBES REEMPLAZAR ESTOS CÓDIGOS! ---
    # Búscalos en la web de BCRPData. (Ej: 'PN01288PM', 'PN01270PM')
    CODIGO_IPC_MENSUAL = "PN38705PM" 
    CODIGO_TC_MENSUAL = "PN01210PM"

    # Nombres de tus hojas y columnas
    NOMBRE_ARCHIVO = "Base de datos"
    NOMBRE_HOJA = "Indices_BCRP"
    COLUMNAS_API = [CODIGO_IPC_MENSUAL, CODIGO_TC_MENSUAL]
    COLUMNAS_HOJA = ['IPC_Mensual', 'TC_Mensual']
    
    # Unir códigos para la API [cite: 74]
    series_a_pedir = "-".join(COLUMNAS_API)
    
    with st.spinner("Actualizando datos del BCRP..."):
        try:
            # 2. Cargar datos existentes de Google Sheets
            df_existente = cargar_hoja_a_df(cliente_gspread, NOMBRE_ARCHIVO, NOMBRE_HOJA)
            if df_existente is None:
                st.error("No se pudo cargar la hoja 'Indices_BCRP'. Abortando.")
                return

            df_existente['Indice_Mes'] = pd.to_datetime(df_existente['Indice_Mes'], dayfirst=True, errors='coerce')
            ultima_fecha = df_existente['Indice_Mes'].max()
            
            # 3. Determinar el rango de fechas para la API
            fecha_hoy_str = pd.to_datetime('today').strftime('%Y-%m')
            
            if pd.isna(ultima_fecha):
                # Si la hoja está vacía, pedimos los últimos 5 años
                periodo_inicial_str = (pd.to_datetime('today') - pd.DateOffset(years=5)).strftime('%Y-%m')
            else:
                # Pedimos desde el mes SIGUIENTE al último que tenemos
                periodo_inicial_str = (ultima_fecha + pd.DateOffset(months=1)).strftime('%Y-%m')

            if periodo_inicial_str > fecha_hoy_str:
                st.success("¡Datos actualizados! No se encontraron nuevos registros.")
                st.cache_data.clear() # Limpiar caché por si acaso
                return

            # 4. Construir y llamar a la API del BCRP [cite: 71, 97]
            url_api = f"https://estadisticas.bcrp.gob.pe/estadisticas/series/api/{series_a_pedir}/json/{periodo_inicial_str}/{fecha_hoy_str}/"
            
            response = requests.get(url_api)
            response.raise_for_status() # Lanza un error si la petición falla
            
            data = response.json()
            
            # 5. Procesar y parsear la respuesta JSON
            nuevos_registros = []
            meses_map = {'Ene': 1, 'Feb': 2, 'Mar': 3, 'Abr': 4, 'May': 5, 'Jun': 6,
                         'Jul': 7, 'Ago': 8, 'Set': 9, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dic': 12}
            
            # --- INICIO: Lógica robusta de mapeo de API (v2 - por NOMBRE) ---
            # 1. Definir palabras clave ÚNICAS para cada serie
            MAP_PALABRA_CLAVE_A_COLUMNA = {
                'IPC': 'IPC_Mensual', # Si 'name' contiene "IPC"
                'Tipo de cambio': 'TC_Mensual' # Si 'name' contiene "Tipo de cambio"
            }
            
            # 2. Leer la configuración de la API para saber el orden real de los valores
            map_index_a_columna = {}
            api_series_config = data.get('config', {}).get('series', [])
            
            for index, series_info in enumerate(api_series_config):
                nombre_api = series_info.get('name', '') # <-- OBTENER EL NOMBRE
                
                # Buscar la palabra clave en el nombre
                for palabra_clave, nombre_columna_hoja in MAP_PALABRA_CLAVE_A_COLUMNA.items():
                    if palabra_clave in nombre_api:
                        map_index_a_columna[index] = nombre_columna_hoja # Ej: {0: 'TC_Mensual', 1: 'IPC_Mensual'}
                        break # Salir del bucle interno una vez encontrada
            # --- FIN: Lógica robusta de mapeo (v2) ---

            # --- INICIO: Verificación de Mapeo ---
            if not map_index_a_columna or len(map_index_a_columna) < len(COLUMNAS_API):
                st.error("Error: La respuesta de la API del BCRP no incluyó una configuración de series válida o faltan códigos.")
                st.warning(f"Mapeo de columnas resultante: {map_index_a_columna}")
                st.warning(f"Respuesta de 'config' de la API: {data.get('config', 'No encontrada')}")
                return # Detener la ejecución de la función
            # --- FIN: Verificación de Mapeo ---

            for periodo in data.get('periods', []):
                try:
                    # Parsear la fecha "Ene.2024"
                    mes_str, anio_str = periodo['name'].split('.')
                    mes_num = meses_map[mes_str.capitalize()]
                    fecha_registro = pd.to_datetime(f"{anio_str}-{mes_num}-01")
                    
                    # Crear el diccionario de datos
                    registro = {'Indice_Mes': fecha_registro}
                    valores = periodo['values'] # Ej: ["3.5432", "115.5868"]
                    
                    # --- INICIO: Asignación corregida ---
                    # Asignar valores usando el mapeo que creamos
                    for index, col_nombre in map_index_a_columna.items():
                        try:
                            # Convertir a float, manejando comas si las hubiera
                            valor_limpio = valores[index].replace(',', '.') if isinstance(valores[index], str) else valores[index]
                            registro[col_nombre] = float(valor_limpio)
                        except (IndexError, ValueError, TypeError):
                            registro[col_nombre] = None # Poner nulo si el valor no es numérico
                    # --- FIN: Asignación corregida ---
                    
                    nuevos_registros.append(registro)

                except Exception as e_parse:
                    st.warning(f"No se pudo procesar el periodo '{periodo.get('name', 'N/A')}'. Error: {e_parse}")
            
            if not nuevos_registros:
                st.success("¡Datos actualizados! No se encontraron nuevos registros.")
                st.cache_data.clear()
                return
                
            # 6. Convertir a DataFrame y filtrar duplicados (por si acaso)
            df_nuevos = pd.DataFrame(nuevos_registros)
            df_nuevos = df_nuevos.dropna(subset=COLUMNAS_HOJA) # Eliminar filas donde falten datos
            df_nuevos_filtrados = df_nuevos[~df_nuevos['Indice_Mes'].isin(df_existente['Indice_Mes'])]

            if df_nuevos_filtrados.empty:
                st.success("¡Datos actualizados! No se encontraron nuevos registros.")
                st.cache_data.clear()
                return

            # 7. Enviar los datos nuevos a Google Sheets
            # (Solo enviamos las 3 columnas que nos importan)
            df_final_a_subir = df_nuevos_filtrados[['Indice_Mes', 'IPC_Mensual', 'TC_Mensual']]
            
            num_filas_anadidas = actualizar_hoja_con_df(cliente_gspread, NOMBRE_ARCHIVO, NOMBRE_HOJA, df_final_a_subir)
            
            if num_filas_anadidas > 0:
                st.success(f"¡Éxito! Se añadieron {num_filas_anadidas} nuevos registros a '{NOMBRE_HOJA}'.")
                st.cache_data.clear() # MUY IMPORTANTE: Limpia el caché para que el resto de la app lea los datos nuevos.
            else:
                st.error("No se pudo añadir los nuevos registros a la hoja de cálculo.")

        except requests.exceptions.HTTPError as e_http:
            st.error(f"Error al contactar la API del BCRP: {e_http}")
            st.error(f"URL consultada: {url_api}")
        except Exception as e:
            st.error(f"Ocurrió un error inesperado durante la actualización: {e}")
            traceback.print_exc()

# --- FIN: Lógica de Actualización BCRP ---

# --- INICIO: Botón de Actualización BCRP ---
if st.button("Sincronizar datos del BCRP"):
    # Esta función la definiremos a continuación
    actualizar_datos_bcrp(conectar_gsheet()) 
# --- FIN: Botón ---

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
        # En app.py, dentro del Paso 1
        hojas_disponibles = [
            format_date(datetime.now() - pd.DateOffset(months=i), "MMMM yyyy", locale='es').capitalize().replace(
                "Septiembre", "Setiembre") for i in range(3)]
        mes_seleccionado = st.selectbox("Selecciona el mes de la asignación:", options=hojas_disponibles)
    with col2:
        df_asignaciones = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_ASIGNACIONES, mes_seleccionado)
        if df_asignaciones is not None:
            num_expediente_input = st.text_input("Ingresa el N° de Expediente:", placeholder="Ej: 1234-2023 o 1234-2023-OEFA/DFAI/PAS")
            
            if st.button("Buscar Expediente", type="primary"):
                # --- LIMPIEZA PROFUNDA DE SESIÓN ---
                claves_a_limpiar = [
                    'info_expediente', 'imputaciones_data', 'context_data', 
                    'num_expediente_formateado', 'rubro_seleccionado', 
                    'id_rubro_seleccionado', 'id_subdireccion_seleccionada',
                    'confiscatoriedad', 'numero_rsd_base', 'fecha_rsd',
                    'numero_ifi', 'fecha_ifi', 'num_informe_multa_ifi',
                    'monto_multa_ifi', 'num_imputaciones_ifi'
                ]
                for clave in claves_a_limpiar:
                    if clave in st.session_state:
                        del st.session_state[clave]
                st.rerun() # Forzamos recarga para asegurar limpieza
            
            if num_expediente_input:
                num_formateado = ""
                if "OEFA" in num_expediente_input.upper():
                    num_formateado = num_expediente_input
                elif "-" in num_expediente_input:
                    num_formateado = f"{num_expediente_input}-OEFA/DFAI/PAS"

                if num_formateado:
                    resultado = df_asignaciones[df_asignaciones['EXPEDIENTE'] == num_formateado]
                    if not resultado.empty:
                        # --- CORRECCIÓN: Actualizar si el número es diferente al actual ---
                        if st.session_state.get('num_expediente_formateado') != num_formateado:
                            st.session_state.num_expediente_formateado = num_formateado
                            st.session_state.info_expediente = resultado.iloc[0].to_dict()
                            
                            # Limpiar datos de hechos anteriores al cambiar de expediente
                            if 'imputaciones_data' in st.session_state:
                                del st.session_state['imputaciones_data']
                            
                            st.success(f"¡Expediente '{num_formateado}' cargado correctamente!")
                        # --- BLOQUE DE CARGA OPTIMIZADO ---
                        if st.button("📥 Cargar Avance Guardado"):
                            expediente_a_cargar = st.session_state.num_expediente_formateado
                            with st.spinner("Buscando avance..."):
                                datos_cargados, mensaje = cargar_datos_caso(cliente_gspread, expediente_a_cargar)
                            
                            if datos_cargados:
                                # Función interna para reconstruir fechas
                                def restaurar_fechas(obj):
                                    if isinstance(obj, dict):
                                        for k, v in obj.items():
                                            if isinstance(v, str) and len(v) == 10 and v.count('-') == 2:
                                                try:
                                                    obj[k] = date.fromisoformat(v)
                                                except: pass
                                            else: restaurar_fechas(v)
                                    elif isinstance(obj, list):
                                        for i in obj: restaurar_fechas(i)

                                restaurar_fechas(datos_cargados)
                                
                                # Inyectar en el session_state
                                for key, value in datos_cargados.items():
                                    st.session_state[key] = value
                                
                                st.success("Datos cargados. Los cálculos se actualizarán al presionar 'Generar Informe'.")
                                st.rerun()
                            else:
                                st.warning(mensaje)

                        
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

# Reemplaza todo desde aquí en tu app.py

    # --- PASO 2 Y LÓGICA SUBSIGUIENTE ---
    if st.session_state.get('info_expediente'):
        st.header("Paso 2: Detalles del Expediente")
        info_caso = st.session_state.info_expediente

        # --- Subsección: Datos del Expediente ---
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
            st.text_input("Nombre o Razón Social del administrado", value=info_caso.get('ADMINISTRADO'), disabled=True)
            st.text_input("Producto", value=info_caso.get('PRODUCTO'), disabled=True)

        with col_info2:
            st.text_input("Analista Económico", value=nombre_completo_analista, disabled=True)
            st.text_input("Sector", value=info_caso.get('SECTOR'), disabled=True)

            df_sector_subdireccion = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Sector_Subdireccion")
            if df_sector_subdireccion is not None and 'ID_Rubro' in df_sector_subdireccion.columns:
                sector_del_caso = info_caso.get('SECTOR')
                if sector_del_caso:
                    rubros_filtrados_df = df_sector_subdireccion[df_sector_subdireccion['Sector_Base'] == sector_del_caso]
                    if not rubros_filtrados_df.empty:
                        lista_rubros = rubros_filtrados_df['Sector_Rubro'].tolist()
                        index_seleccionado = None
                        rubro_guardado = st.session_state.get('rubro_seleccionado')
                        if rubro_guardado and rubro_guardado in lista_rubros:
                            index_seleccionado = lista_rubros.index(rubro_guardado)
                        nombre_rubro_seleccionado = st.selectbox("Elige el subsector/rubro", options=lista_rubros, index=index_seleccionado, placeholder="Selecciona una opción...")
                        if nombre_rubro_seleccionado:
                            st.session_state.rubro_seleccionado = nombre_rubro_seleccionado
                            fila_rubro = rubros_filtrados_df[rubros_filtrados_df['Sector_Rubro'] == nombre_rubro_seleccionado]
                            if not fila_rubro.empty:
                                id_rubro = fila_rubro.iloc[0]['ID_Rubro']
                                id_subdireccion = fila_rubro.iloc[0]['ID_Subdireccion'] # <--- AÑADIR
                                st.session_state.id_rubro_seleccionado = id_rubro
                                st.session_state.id_subdireccion_seleccionada = id_subdireccion # <--- AÑADIR
                    else:
                        st.warning(f"No hay rubros para el sector '{sector_del_caso}'.")
            else:
                st.error("No se pudo cargar la hoja 'Sector_Subdireccion'.")

        st.subheader("Fecha del Informe")

        # Inicializamos el valor en el estado solo si no existe
        if 'fecha_emision_informe' not in st.session_state:
            st.session_state['fecha_emision_informe'] = date.today()

        # El widget ahora solo usa el 'key', Streamlit manejará el valor automáticamente
        fecha_emision_informe = st.date_input(
            "Selecciona la fecha de emisión del informe:",
            key='fecha_emision_informe',
            format="DD/MM/YYYY"
        )

        # --- Subsección: Resolución Subdirectoral (RSD) ---
        producto_caso = st.session_state.info_expediente.get('PRODUCTO', '')

        # Solo mostramos estos campos si el producto requiere una RSD de base (IFI y RD)
        if producto_caso != 'COERCITIVA':
            st.subheader("Resolución Subdirectoral (RSD)")
            col_rsd1, col_rsd2 = st.columns([1, 1])
            
            with col_rsd1:
                if 'numero_rsd_base' not in st.session_state:
                    st.session_state['numero_rsd_base'] = ''

                st.text_input(
                    "N° de RSD:", 
                    key='numero_rsd_base',
                    placeholder="Ej. 00245-2025"
                )

                numero_rsd_base = st.session_state.get('numero_rsd_base', '')
                numero_rsd_completo = ""
                if numero_rsd_base:
                    suffix_sub = st.session_state.get('id_subdireccion_seleccionada', 'ERROR_SUB')
                    numero_rsd_completo = f"{numero_rsd_base}-OEFA/DFAI-{suffix_sub}"
                
                if numero_rsd_completo:
                    st.info(f"**RSD:** {numero_rsd_completo}")
                st.session_state.numero_rsd = numero_rsd_completo 

            with col_rsd2:
                if 'fecha_rsd' not in st.session_state:
                    st.session_state['fecha_rsd'] = date.today()

                st.date_input(
                    "Fecha de notificación de la RSD:", 
                    key='fecha_rsd',
                    format="DD/MM/YYYY"
                )

        # --- INICIO: SECCIÓN EXCLUSIVA PARA RD (IFI + INFORME MULTA) ---
        producto_caso = st.session_state.info_expediente.get('PRODUCTO', '')
        
        # --- NUEVO ORDEN: Primero capturamos los inputs del IFI ---
        if producto_caso == 'RD':
            st.divider()
            st.subheader("Informe Final de Instrucción (IFI)")
            
            col_ifi1, col_ifi2 = st.columns(2)
            with col_ifi1:
                if 'numero_ifi_base' not in st.session_state: 
                    st.session_state.numero_ifi_base = ''
                
                st.text_input("N° del IFI:", key="numero_ifi_base", placeholder="Ej: 00123-2023")
                
                # Lógica de autocompletado
                ifi_base = st.session_state.get('numero_ifi_base', '')
                ifi_completo = ""
                if ifi_base:
                    suffix_sub = st.session_state.get('id_subdireccion_seleccionada', 'SECTOR')
                    ifi_completo = f"{ifi_base}-OEFA/DFAI-{suffix_sub}"
                
                if ifi_completo:
                    st.info(f"**IFI:** {ifi_completo}")
                
                # Guardamos el resultado final para que lo use el informe
                st.session_state.numero_ifi = ifi_completo
            with col_ifi2:
                if 'fecha_ifi' not in st.session_state: st.session_state.fecha_ifi = date.today()
                st.date_input("Fecha de notificación del IFI:", key="fecha_ifi", format="DD/MM/YYYY")
            
            with st.container(border=True):
                st.markdown("###### Datos del Informe de Multa (IFI)")
                col_im1, col_im2, col_im3 = st.columns(3)
                with col_im1:
                    if 'num_informe_multa_ifi_base' not in st.session_state: 
                        st.session_state.num_informe_multa_ifi_base = ''
                    
                    st.text_input("N° de Informe de Multa:", key="num_informe_multa_ifi_base", placeholder="Ej: 00045-2024")
                    
                    # Lógica de autocompletado
                    im_base = st.session_state.get('num_informe_multa_ifi_base', '')
                    im_completo = ""
                    if im_base:
                        im_completo = f"{im_base}-OEFA/DFAI-SSAG"
                    
                    if im_completo:
                        st.info(f"**Inf. Multa:** {im_completo}")
                    
                    # Guardamos el resultado final para que lo use el informe
                    st.session_state.num_informe_multa_ifi = im_completo
                with col_im2:
                    if 'monto_multa_ifi' not in st.session_state: st.session_state.monto_multa_ifi = 0.0
                    st.number_input("Monto Total Propuesto (UIT):", key="monto_multa_ifi", format="%.3f")
                with col_im3:
                    if 'num_imputaciones_ifi' not in st.session_state: st.session_state.num_imputaciones_ifi = 1
                    st.number_input("Nº Imputaciones:", key="num_imputaciones_ifi", min_value=1)

        # --- RE-CALCULAR resolucion_ok AQUÍ ---
        resolucion_ok = False
        if producto_caso == 'RD':
            if (st.session_state.get('numero_rsd') and st.session_state.get('fecha_rsd') and 
                st.session_state.get('numero_ifi') and st.session_state.get('fecha_ifi')):
                resolucion_ok = True
        elif producto_caso == 'COERCITIVA':
            resolucion_ok = True
        else: # IFI
            if st.session_state.get('numero_rsd') and st.session_state.get('fecha_rsd'):
                resolucion_ok = True

        resolucion_ok = False
        if st.session_state.get('info_expediente'):
            producto_caso = st.session_state.info_expediente.get('PRODUCTO', '')
            
            if producto_caso == 'RD':
                # Para RD seguimos exigiendo todo (IFI y RSD)
                if (st.session_state.get('numero_rsd') and 
                    st.session_state.get('fecha_rsd') and 
                    st.session_state.get('numero_ifi') and 
                    st.session_state.get('fecha_ifi')):
                    resolucion_ok = True
            
            elif producto_caso == 'COERCITIVA':
                # No es necesario ingresar RSD para avanzar
                resolucion_ok = True
                # Aseguramos que las variables no sean None para evitar errores en el contexto
                if 'numero_rsd' not in st.session_state: st.session_state.numero_rsd = ""
                if 'fecha_rsd' not in st.session_state: st.session_state.fecha_rsd = date.today()
            
            else: # IFI
                # Para IFI exigimos solo la RSD
                if st.session_state.get('numero_rsd') and st.session_state.get('fecha_rsd'):
                    resolucion_ok = True

        if st.session_state.get('rubro_seleccionado') and resolucion_ok:
            with st.spinner("Preparando datos generales..."):

                from funciones import AcronymManager

                df_analistas = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Analistas")
                df_subdirecciones = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Subdirecciones")
                df_sector_sub = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Sector_Subdireccion")
                df_producto_asunto = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Producto_Asunto")
                df_administrados = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Administrados")
                df_indices_final = cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Indices_BCRP")

                # --- INICIO: (REQ 1) USAR FECHA DE EMISIÓN ---
                fecha_actual = st.session_state.get('fecha_emision_informe', date.today())
                # --- FIN: (REQ 1) ---
                
                fecha_rsd_dt = st.session_state.get('fecha_rsd') or fecha_actual
                context_data = {
                    'fecha_hoy': format_date(fecha_actual, "d 'de' MMMM 'de' yyyy", locale='es').replace("septiembre", "setiembre").replace("Septiembre", "Setiembre"),
                    'mes_hoy': format_date(fecha_actual, "MMMM 'de' yyyy", locale='es').lower().replace("septiembre", "setiembre"),
                    'fecha_rsd_texto': format_date(fecha_rsd_dt, "d 'de' MMMM 'de' yyyy", locale='es').replace("septiembre", "setiembre").replace("Septiembre", "Setiembre"),
                    'acronyms': AcronymManager(),
                    'numbering': NumberingManager(),
                    'nombre_administrado': info_caso.get('ADMINISTRADO', '')
                }
                nombre_base_administrado = info_caso.get('ADMINISTRADO', '')
                nombre_final_administrado = nombre_base_administrado
                if df_administrados is not None:
                    admin_info = df_administrados[df_administrados['Nombre_Administrado_Base'] == nombre_base_administrado]
                    if not admin_info.empty:
                        nombre_final_administrado = admin_info.iloc[0].get('Nombre_Administrado', nombre_base_administrado)
                context_data['administrado'] = nombre_final_administrado
                analista_details = get_person_details_by_base_name(info_caso.get('ANALISTA ECONÓMICO'), df_analistas)
                revisor_details = get_person_details_by_base_name(info_caso.get('REVISOR'), df_analistas)
                context_data.update({
                    'titulo_analista': analista_details['titulo'], 'nombre_analista': analista_details['nombre'],
                    'cargo_analista': analista_details['cargo'], 'colegiatura_analista': analista_details['colegiatura'],
                    'titulo_revisor': revisor_details['titulo'], 'nombre_revisor': revisor_details['nombre'],
                    'cargo_revisor': revisor_details['cargo'], 'colegiatura_revisor': revisor_details['colegiatura']
                })
                num_imputaciones = int(info_caso.get('IMPUTACIONES', 1))
                inf_texto = "una" if num_imputaciones == 1 else num2words(num_imputaciones, lang='es')
                
                # --- INICIO DE LA ADICIÓN ---
                texto_incumplimiento_plural = "el incumplimiento" if num_imputaciones == 1 else "los incumplimientos"
                # --- FIN DE LA ADICIÓN ---

                # --- INICIO DE LA ADICIÓN (REQ. 1) ---
                plural_infraccion_analizada = "la infracción analizada" if num_imputaciones == 1 else "las infracciones analizadas"
                plural_cada_infraccion = "la infracción" if num_imputaciones == 1 else "cada infracción"
                # --- FIN DE LA ADICIÓN ---

                # --- INICIO: OPTIMIZACIÓN 1 (Placeholders S/P) ---
                ph_hecho_imputado = "el hecho imputado" if num_imputaciones == 1 else "los hechos imputados"
                ph_conducta_infractora = "la conducta infractora" if num_imputaciones == 1 else "las conductas infractoras"
                ph_calculo_multa = "cálculo de multa" if num_imputaciones == 1 else "cálculo de multas"
                ph_la_infraccion = "la infracción" if num_imputaciones == 1 else "las infracciones"
                ph_hecho_imputado_ant = "del hecho imputado referido" if num_imputaciones == 1 else "de los hechos imputados referidos"
                # --- FIN: OPTIMIZACIÓN 1 ---
                ph_multa_propuesta = "la multa propuesta resulta" if num_imputaciones == 1 else "las multas propuestas resultan"
                
                context_data.update({
                    'inf_numero': num_imputaciones, 'inf_texto': inf_texto,
                    'inf_oracion': "presunta infracción administrativa" if num_imputaciones == 1 else "presuntas infracciones administrativas",
                    'hechos_plural_objeto': "al hecho imputado mencionado" if num_imputaciones == 1 else "a los hechos imputados mencionados", # <-- AÑADIR ESTA LÍNEA
                    'texto_incumplimiento': texto_incumplimiento_plural, # <-- AÑADIR ESTA LÍNEA
                    # --- AÑADIR ESTAS LÍNEAS ---
                    'plural_infraccion_analizada': plural_infraccion_analizada,
                    'plural_cada_infraccion': plural_cada_infraccion,
                    # --- FIN ---
                    # --- INICIO: AÑADIR NUEVOS PLACEHOLDERS ---
                    'ph_hecho_imputado': ph_hecho_imputado,
                    'ph_conducta_infractora': ph_conducta_infractora,
                    'ph_calculo_multa': ph_calculo_multa,
                    'ph_la_infraccion': ph_la_infraccion,
                    'ph_hecho_imputado_ant': ph_hecho_imputado_ant,
                    'ph_multa_propuesta': ph_multa_propuesta
                    # --- FIN: AÑADIR NUEVOS PLACEHOLDERS ---
                })
                context_data.update({
                    'expediente': st.session_state.get('num_expediente_formateado', ''),
                    'ht': info_caso.get('HT', ''),
                    'numero_rsd': st.session_state.get('numero_rsd', ''),
                    'numero_ifi': st.session_state.get('numero_ifi', ''),
                    'fecha_ifi': format_date(st.session_state.get('fecha_ifi'), "d 'de' MMMM 'de' yyyy", locale='es') if st.session_state.get('fecha_ifi') else '',
                    
                    # Datos del Informe de Multa previo
                    'num_informe_multa_ifi': st.session_state.get('num_informe_multa_ifi', ''),
                    'monto_multa_ifi': f"{st.session_state.get('monto_multa_ifi', 0.0):,.3f} UIT",
                    'num_imputaciones_ifi': st.session_state.get('num_imputaciones_ifi', 0)
                })
                # Lógica para obtener datos de la subdirección y SSAG
                # --- Lógica de Encargado Principal (sub1) ---
                # Para RD y COERCITIVA siempre se dirige a la DFAI (Director)
                if producto_caso in ['RD', 'COERCITIVA']:
                    sub_row = df_subdirecciones[df_subdirecciones['ID_Subdireccion'].astype(str).str.strip().str.upper() == 'DFAI']
                else:
                    # Para IFI y otros, se mantiene el encargado del sector (Subdirector)
                    id_sub_row = df_sector_sub[df_sector_sub['Sector_Rubro'] == st.session_state.rubro_seleccionado]
                    sub_row = pd.DataFrame()
                    if not id_sub_row.empty:
                        id_sub_id = id_sub_row.iloc[0].get('ID_Subdireccion')
                        sub_row = df_subdirecciones[df_subdirecciones['ID_Subdireccion'] == id_sub_id]

                if not sub_row.empty:
                    # Estos placeholders ahora tendrán los datos de la DFAI en RD/COERCITIVA
                    context_data['nombre_encargado_sub1'] = sub_row.iloc[0].get('Encargado_Sub', '')
                    context_data['cargo_encargado_sub1'] = sub_row.iloc[0].get('Cargo_Encargado_Sub', '')
                    context_data['titulo_encargado_sub1'] = sub_row.iloc[0].get('Titulo_Encargado_Sub', '')
                    context_data['subdireccion'] = sub_row.iloc[0].get('Subdireccion', '')
                    context_data['id_subdireccion'] = sub_row.iloc[0].get('ID_Subdireccion', '')

                ssag_row = df_subdirecciones[df_subdirecciones['ID_Subdireccion'].astype(str).str.strip().str.upper() == 'SSAG']
                if not ssag_row.empty:
                    nombre_enc_ssag = ssag_row.iloc[0].get('Encargado_Sub')
                    context_data.update({
                        'nombre_encargado_sub2': nombre_enc_ssag,
                        'titulo_encargado_sub2': ssag_row.iloc[0].get('Titulo_Encargado_Sub', ''),
                        'cargo_encargado_sub2': ssag_row.iloc[0].get('Cargo_Encargado_Sub', '')
                    })
                    
                    # --- INICIO REQ 4: Placeholders de Iniciales ---
                    ssag_base_name = ''
                    if nombre_enc_ssag and df_analistas is not None:
                        enc_ssag_analista_row = df_analistas[df_analistas['Nombre_Analista'] == nombre_enc_ssag]
                        if not enc_ssag_analista_row.empty:
                            context_data['colegiatura_encargado_sub2'] = enc_ssag_analista_row.iloc[0].get('Colegiatura_Analista', '')
                            # Extraer el Nombre_Base_Analista (ej: RMACHUCA)
                            ssag_base_name = enc_ssag_analista_row.iloc[0].get('Nombre_Base_Analista', '')
                    
                    # --- INICIO REQ 4: Placeholders de Iniciales (Corrección v3) ---
                    
                    # Obtener nombres completos
                    nombre_completo_sub = nombre_enc_ssag
                    nombre_completo_rev = revisor_details.get('nombre', '')
                    nombre_completo_ana = analista_details.get('nombre', '')

                    # --- Placeholder 1: [RMACHUCA] ---
                    # Requerimiento: [PrimeraLetraNombre][PrimerApellido]
                    placeholder_corchetes = "[SSAG_SUBDIRECTOR]" # Default
                    if nombre_completo_sub:
                        parts = nombre_completo_sub.split()
                        # Asegurarse de que parts[0] no esté vacío
                        primera_letra = parts[0][0].upper() if parts and parts[0] else ''
                        apellido = ""

                        # Lógica para encontrar el primer apellido
                        if len(parts) == 2: # Ej: Ricardo Machuca
                            apellido = parts[1].upper()
                        elif len(parts) == 3: # Ej: Ricardo Machuca Bravo
                            apellido = parts[1].upper() # Asume que el 2do es el primer apellido
                        elif len(parts) >= 4: # Ej: Ricardo Oscar Machuca Bravo
                            apellido = parts[2].upper() # Asume que el 3ro es el primer apellido
                        
                        if primera_letra and apellido:
                            placeholder_corchetes = f"[{primera_letra}{apellido}]"
                        elif primera_letra: # Fallback si solo hay un nombre/palabra
                            placeholder_corchetes = f"[{parts[0].upper()}]"
                            
                    context_data['ssag_iniciales_corchetes'] = placeholder_corchetes
                    
                    # --- Placeholder 2: ROMB/tv/ec ---
                    # Requerimiento: INICIALES_COMPLETAS_SUB (MAYUS) / iniciales_completas_rev (minus) / iniciales_completas_ana (minus)
                    
                    # Calcular iniciales completas usando la función de funciones.py
                    ssag_iniciales = get_initials_from_name(nombre_completo_sub, to_lower=False) # MAYUS
                    revisor_iniciales = get_initials_from_name(nombre_completo_rev, to_lower=True) # minus
                    analista_iniciales = get_initials_from_name(nombre_completo_ana, to_lower=True) # minus
                    
                    context_data['ssag_iniciales_linea'] = f"{ssag_iniciales or 'SSAG'}/{revisor_iniciales or 'rev'}/{analista_iniciales or 'ana'}"
                    # --- FIN REQ 4 (Corrección v3) ---

                if producto_caso and df_producto_asunto is not None:
                    asunto_row = df_producto_asunto[df_producto_asunto['Producto'] == producto_caso]
                    if not asunto_row.empty:
                        context_data['asunto'] = asunto_row.iloc[0].get('Producto_Asunto', '')
                mes_indice_texto = "No disponible"
                if df_indices_final is not None and not df_indices_final.empty:
                    try:
                        df_indices_final['Indice_Mes_dt'] = pd.to_datetime(df_indices_final['Indice_Mes'], dayfirst=True, errors='coerce')
                        latest_date = df_indices_final['Indice_Mes_dt'].max()
                        if pd.notna(latest_date):
                            mes_indice_texto = format_date(latest_date, 'MMMM yyyy', locale='es').lower().replace("septiembre", "setiembre")
                    except Exception: pass
                context_data['mes_indice'] = mes_indice_texto
                st.session_state.context_data = context_data
                st.success("Datos generales preparados.")

        st.divider()

        # --- PASO 3: LÓGICA CONDICIONAL POR TIPO DE PRODUCTO ---
        if st.session_state.get('context_data'):
            
            # Obtener el tipo de producto del expediente para decidir qué hacer
            producto_caso = st.session_state.info_expediente.get('PRODUCTO', '')

            # ----------------------------------------------------
            # ---- CASO 1: El producto es "COERCITIVA" ----
            # ----------------------------------------------------
            if producto_caso == 'COERCITIVA':
                
                # Importar las funciones del nuevo módulo
                from producto_coercitiva import renderizar_inputs_coercitiva, validar_inputs_coercitiva, procesar_coercitiva

                # Inicializar el diccionario de estado para coercitiva si no existe
                if 'datos_informe_coercitiva' not in st.session_state:
                    st.session_state.datos_informe_coercitiva = {}

                # 1. RENDERIZAR LA INTERFAZ
                st.session_state.datos_informe_coercitiva = renderizar_inputs_coercitiva(st.session_state.datos_informe_coercitiva)

                # 2. VALIDAR LOS INPUTS
                boton_habilitado = validar_inputs_coercitiva(st.session_state.datos_informe_coercitiva)

                st.divider()
                st.header("Paso 4: Generar Informe Coercitivo")

                # 3. BOTÓN PARA PROCESAR Y GENERAR
                if st.button("🚀 Generar Informe Coercitivo", type="primary", disabled=(not boton_habilitado)):
                    with st.spinner("Generando informe..."):
                        # Preparar datos comunes que necesita el módulo de coercitiva
                        datos_comunes_coercitiva = {
                            'cliente_gspread': cliente_gspread,
                            'context_data': st.session_state.get('context_data', {}),
                            'df_productos': cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Productos"),
                        }
                        # Guardar resultados en el estado de la sesión
                        st.session_state.resultados_coercitiva = procesar_coercitiva(datos_comunes_coercitiva, st.session_state.datos_informe_coercitiva)

                # Mostrar resultados y botón de descarga si existen
                if 'resultados_coercitiva' in st.session_state:
                    resultados_finales = st.session_state.resultados_coercitiva
                    if resultados_finales and not resultados_finales.get('error'):
                        st.success("¡Informe Coercitivo generado con éxito!")
                        
                        # Mostrar tabla resumen en la app
                        resultados_app = resultados_finales.get('resultados_para_app', {})
                        if resultados_app.get('tabla_resumen_coercitivas'):
                            st.markdown("###### Resumen de Multas Coercitivas Calculadas")
                            df_resumen = pd.DataFrame(resultados_app['tabla_resumen_coercitivas'])
                            # Formatear columnas para visualización
                            df_resumen_display = df_resumen.rename(columns={
                                'num_medida': 'N° Medida',
                                'multa_base_uit': 'Multa Base (UIT)',
                                'num_coercitiva_texto': 'N° Coercitiva',
                                'multa_coercitiva_final_uit': 'Coercitiva a Imponer (UIT)'
                            })
                            st.dataframe(
                                df_resumen_display[['N° Medida', 'Multa Base (UIT)', 'N° Coercitiva', 'Coercitiva a Imponer (UIT)']],
                                column_config={
                                    "Multa Base (UIT)": st.column_config.NumberColumn(format="%.3f"),
                                    "Coercitiva a Imponer (UIT)": st.column_config.NumberColumn(format="%.3f"),
                                },
                                use_container_width=True, 
                                hide_index=True
                            )
                            st.metric("Multa Coercitiva Total (UIT)", f"{resultados_app.get('total_uit', 0):,.3f}")

                        # Botón de descarga
                        nombre_exp = st.session_state.get('num_expediente_formateado', 'EXPEDIENTE')
                        st.download_button(
                            label="✅ Descargar Informe Coercitivo (.docx)",
                            data=resultados_finales['doc_pre_compuesto'].getvalue(),
                            file_name=f"Informe_Coercitiva_{nombre_exp.replace('/', '-')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="primary"
                        )
                        # Limpiar estado para la siguiente ejecución
                        del st.session_state.resultados_coercitiva
                    elif resultados_finales.get('error'):
                        st.error(f"Error: {resultados_finales['error']}")
                        del st.session_state.resultados_coercitiva

            # ----------------------------------------------------
            # ---- LÓGICA EXISTENTE PARA IFI/RD (INFRACCIONES) ----
            # ----------------------------------------------------
            elif producto_caso in ['IFI', 'RD']: # O los productos que ya manejas
                st.header("Paso 3: Detalles de Hechos Imputados")

                # --- INICIO DE LA OPTIMIZACIÓN ---
                # Cargamos todos los DataFrames necesarios para los cálculos UNA SOLA VEZ
                with st.spinner("Cargando datos para cálculos..."):
                    if 'datos_calculo' not in st.session_state:
                        st.session_state.datos_calculo = {
                            'df_tipificacion': cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Tipificacion_Infracciones"),
                            'df_items_infracciones': cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Items_Infracciones"),
                            'df_costos_items': cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Costos_Items"),
                            'df_salarios_general': cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Salarios_General"),
                            'df_cos': cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "COS"),
                            'df_uit': cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "UIT"),
                            'df_coti_general': pd.to_datetime(cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Cotizaciones_General")['Fecha_Costeo'], dayfirst=True, errors='coerce').to_frame().join(cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Cotizaciones_General").drop('Fecha_Costeo', axis=1)),
                            'df_indices': pd.to_datetime(cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Indices_BCRP")['Indice_Mes'], dayfirst=True, errors='coerce').to_frame().join(cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Indices_BCRP").drop('Indice_Mes', axis=1)),
                            'df_dias_no_laborables': cargar_hoja_a_df(cliente_gspread, NOMBRE_GSHEET_MAESTRO, "Dias_No_Laborables")
                        }
                
                # Obtenemos los dataframes desde el estado de la sesión para usarlos
                datos_calculo = st.session_state.datos_calculo
                df_tipificacion = datos_calculo['df_tipificacion']
                # --- FIN DE LA OPTIMIZACIÓN ---
                
                for i in range(len(st.session_state.get('imputaciones_data', []))):
                    with st.expander(f"Hecho imputado n.° {i + 1}", expanded=(i == 0)):
                        st.subheader(f"Detalles del Hecho {i + 1}")
                        
                        # --- Lógica de widgets con estado corregida ---
                        texto_guardado = st.session_state.imputaciones_data[i].get('texto_hecho', '')
                        texto_ingresado = st.text_area("Redacta aquí el hecho imputado:", value=texto_guardado, key=f"texto_hecho_{i}", height=150)
                        st.session_state.imputaciones_data[i]['texto_hecho'] = texto_ingresado
                        st.divider()

                        if df_tipificacion is not None:
                            try:
                                lista_tipos_infraccion = df_tipificacion['Tipo_Infraccion'].unique().tolist()
                                index_tipo = None
                                tipo_guardado = st.session_state.imputaciones_data[i].get('tipo_seleccionado')
                                if tipo_guardado in lista_tipos_infraccion:
                                    index_tipo = lista_tipos_infraccion.index(tipo_guardado)
                                tipo_seleccionado = st.radio("**Selecciona el tipo de infracción:**", options=lista_tipos_infraccion, index=index_tipo, horizontal=True, key=f"radio_tipo_infraccion_{i}")
                                st.session_state.imputaciones_data[i]['tipo_seleccionado'] = tipo_seleccionado

                                if tipo_seleccionado:
                                    subtipos_df = df_tipificacion[df_tipificacion['Tipo_Infraccion'] == tipo_seleccionado]
                                    lista_subtipos = subtipos_df['Descripcion_Infraccion'].tolist()
                                    index_subtipo = None
                                    subtipo_guardado = st.session_state.imputaciones_data[i].get('subtipo_seleccionado')
                                    if subtipo_guardado in lista_subtipos:
                                        index_subtipo = lista_subtipos.index(subtipo_guardado)
                                    subtipo_seleccionado = st.selectbox("**Selecciona la descripción de la infracción:**", options=lista_subtipos, index=index_subtipo, placeholder="Elige una descripción específica...", key=f"subtipo_infraccion_{i}")
                                    st.session_state.imputaciones_data[i]['subtipo_seleccionado'] = subtipo_seleccionado

                                    # En app.py, dentro del bucle for del Paso 3

                                    if subtipo_seleccionado:
                                        # --- INICIO DE LA CORRECCIÓN ---
                                        # Paso 1: Buscamos la fila y guardamos el resultado (que puede estar vacío)
                                        fila_encontrada = subtipos_df[subtipos_df['Descripcion_Infraccion'] == subtipo_seleccionado]
                                        
                                        # Paso 2: Solo si la búsqueda tuvo éxito (no está vacía), extraemos los datos
                                        if not fila_encontrada.empty:
                                            id_infraccion = fila_encontrada.iloc[0]['ID_Infraccion']
                                            st.session_state.imputaciones_data[i]['id_infraccion'] = id_infraccion
                                        # --- FIN DE LA CORRECCIÓN ---
                            except KeyError as e:
                                st.error(f"Error en la hoja 'Tipificacion_Infracciones'. Falta la columna: {e}")
                        else:
                            st.error("Error crítico: No se pudo cargar la hoja 'Tipificacion_Infracciones'.")
                            
                        # --- Lógica de cálculo del hecho (ya corregida) ---
                        id_infraccion = st.session_state.imputaciones_data[i].get('id_infraccion')
                        if id_infraccion:
                            try:
                                modulo_especialista = importlib.import_module(f"infracciones.{id_infraccion}")
                                datos_especificos = modulo_especialista.renderizar_inputs_especificos(i, datos_calculo.get('df_dias_no_laborables'))
                                st.session_state.imputaciones_data[i].update(datos_especificos)
                                
                                # --- INICIO: SECCIÓN GRADUACIÓN DE SANCIONES (Ubicación Correcta) ---
                                st.divider()
                                st.subheader("Factores de graduación")
                                
                                datos_hecho_actual = st.session_state.imputaciones_data[i]
                                
                                # 1. Botón de activación inicial
                                aplica_graduacion = st.radio(
                                    "¿Existen factores agravantes o atenuantes para graduar la sanción?",
                                    ["No", "Sí"],
                                    key=f"aplica_grad_{i}",
                                    index=0 if datos_hecho_actual.get('aplica_graduacion', 'No') == 'No' else 1,
                                    horizontal=True
                                )
                                datos_hecho_actual['aplica_graduacion'] = aplica_graduacion
                                
                                factor_f_final = 1.0 # Valor por defecto (Neutro)
                                es_eximente = False
                                
                                if aplica_graduacion == "Sí":
                                    with st.expander("Configurar Factores de Graduación", expanded=True):
                                        st.info("Seleccione los criterios. Por defecto, todos inician en 0%.")
                                        
                                        # Recuperar o inicializar el diccionario de graduación
                                        if 'graduacion' not in datos_hecho_actual:
                                            datos_hecho_actual['graduacion'] = {}
                                        seleccion_graduacion = datos_hecho_actual['graduacion']
                                        
                                        total_porcentaje_f = 0.0

                                        # Iterar sobre f1, f2, ... f7 (Usando la constante FACTORES_GRADUACION)
                                        for codigo_f, data_f in FACTORES_GRADUACION.items():
                                            st.markdown(f"**{codigo_f.upper()}: {data_f['titulo']}**")
                                            
                                            subtotal_f = 0.0
                                            
                                            for subcriterio, opciones in data_f['criterios'].items():
                                                key_widget = f"grad_{i}_{codigo_f}_{subcriterio}"
                                                val_guardado = seleccion_graduacion.get(key_widget)
                                                
                                                # Asegurar que la opción por defecto sea la primera (que suele ser 0.0)
                                                lista_opciones = list(opciones.keys())
                                                idx_sel = 0
                                                if val_guardado in lista_opciones:
                                                    idx_sel = lista_opciones.index(val_guardado)

                                                opcion_sel = st.selectbox(
                                                    subcriterio,
                                                    options=lista_opciones,
                                                    key=key_widget,
                                                    index=idx_sel
                                                )
                                                
                                                valor = opciones[opcion_sel]
                                                
                                                # Guardar selección
                                                seleccion_graduacion[key_widget] = opcion_sel
                                                
                                                # Lógica Eximente
                                                if valor == "Eximente":
                                                    es_eximente = True
                                                    valor_num = 0.0
                                                    seleccion_graduacion[f"{key_widget}_valor"] = "Eximente"
                                                else:
                                                    valor_num = float(valor)
                                                    seleccion_graduacion[f"{key_widget}_valor"] = valor_num

                                                subtotal_f += valor_num
                                            
                                            # Visualización del subtotal
                                            if subtotal_f != 0:
                                                color_txt = "red" if subtotal_f > 0 else "green" # Rojo=Agrava, Verde=Atenúa
                                                st.caption(f":{color_txt}[Subtotal {codigo_f}: {subtotal_f:+.0%}]")
                                            
                                            total_porcentaje_f += subtotal_f
                                            seleccion_graduacion[f"subtotal_{codigo_f}"] = subtotal_f

                                        # Cálculo final de F
                                        if es_eximente:
                                            st.warning("⚠️ **SE APLICA EXIMENTE DE RESPONSABILIDAD** (Multa = 0)")
                                            factor_f_final = 0.0
                                        else:
                                            factor_f_final = 1.0 + total_porcentaje_f
                                            if factor_f_final < 0: factor_f_final = 0.0 # No puede ser negativo
                                            
                                            col_res1, col_res2 = st.columns([3, 1])
                                            with col_res1:
                                                st.markdown(f"**Ajuste total acumulado:** {total_porcentaje_f:+.0%}")
                                            with col_res2:
                                                st.metric("Factor F Final", f"{factor_f_final:.2%}")
                                else:
                                    # Si dice "No", limpiamos o reseteamos
                                    # Opcional: datos_hecho_actual['graduacion'] = {} 
                                    pass

                                # Guardar en el estado para el cálculo posterior
                                datos_hecho_actual['factor_f_calculado'] = factor_f_final
                                datos_hecho_actual['es_eximente'] = es_eximente
                                # --- FIN: SECCIÓN GRADUACIÓN ---

                                # --- INICIO: (REQ 1) AÑADIR INPUTS DE REDUCCIÓN DE MULTA ---
                                st.divider()
                                st.subheader("Reconocimiento de responsabilidad")
                                datos_hecho_actual = st.session_state.imputaciones_data[i]
                                
                                aplica_reduccion = st.radio(
                                    "¿Aplica reducción de multa por reconocimiento de responsabilidad?",
                                    ["No", "Sí"],
                                    key=f"aplica_reduccion_{i}",
                                    index=0 if datos_hecho_actual.get('aplica_reduccion', 'No') == 'No' else 1,
                                    horizontal=True
                                )
                                datos_hecho_actual['aplica_reduccion'] = aplica_reduccion
                                
                                if aplica_reduccion == "Sí":
                                    porcentaje_reduccion = st.radio(
                                        "Seleccione el porcentaje:",
                                        ["50%", "30%"],
                                        key=f"porcentaje_reduccion_{i}",
                                        index=0 if datos_hecho_actual.get('porcentaje_reduccion') == '50%' else 1,
                                        horizontal=True
                                    )
                                    datos_hecho_actual['porcentaje_reduccion'] = porcentaje_reduccion
                                    
                                    # Definir el texto del placeholder basado en el porcentaje
                                    if porcentaje_reduccion == "50%":
                                        datos_hecho_actual['texto_reduccion'] = "en la presentación de los descargos contra la imputación de cargos"
                                    else: # 30%
                                        datos_hecho_actual['texto_reduccion'] = "mediante los descargos al IFI"
                                    
                                    st.info(f"Se aplicará reducción del **{porcentaje_reduccion}** ({datos_hecho_actual['texto_reduccion']}).")
                                    
                                    col_memo1, col_memo2 = st.columns(2)
                                    with col_memo1:
                                        datos_hecho_actual['memo_num'] = st.text_input("N.° de Memorando:", value=datos_hecho_actual.get('memo_num', ''), key=f"memo_num_{i}")
                                    with col_memo2:
                                        datos_hecho_actual['memo_fecha'] = st.date_input("Fecha del Memorando:", value=datos_hecho_actual.get('memo_fecha'), key=f"memo_fecha_{i}", format="DD/MM/YYYY")
                                        
                                    col_esc1, col_esc2 = st.columns(2)
                                    with col_esc1:
                                        datos_hecho_actual['escrito_num'] = st.text_input("N.° de Escrito (Administrado):", value=datos_hecho_actual.get('escrito_num', ''), key=f"escrito_num_{i}")
                                    with col_esc2:
                                        datos_hecho_actual['escrito_fecha'] = st.date_input("Fecha del Escrito:", value=datos_hecho_actual.get('escrito_fecha'), key=f"escrito_fecha_{i}", format="DD/MM/YYYY")
                                
                                # --- FIN: (REQ 1) ---

                                datos_generales_ok = st.session_state.imputaciones_data[i].get('texto_hecho') and st.session_state.imputaciones_data[i].get('subtipo_seleccionado')
                                datos_especificos_ok = modulo_especialista.validar_inputs(st.session_state.imputaciones_data[i])
                                
                                # --- INICIO: (REQ 1) VALIDACIÓN DE REDUCCIÓN ---
                                datos_reduccion_ok = True
                                datos_hecho_actual = st.session_state.imputaciones_data[i]
                                if datos_hecho_actual.get('aplica_reduccion') == 'Sí':
                                    if not all([
                                        datos_hecho_actual.get('porcentaje_reduccion'),
                                        datos_hecho_actual.get('memo_num'),
                                        datos_hecho_actual.get('memo_fecha'),
                                        datos_hecho_actual.get('escrito_num'),
                                        datos_hecho_actual.get('escrito_fecha')
                                    ]):
                                        datos_reduccion_ok = False
                                        st.warning("Para aplicar la reducción, debe completar todos los campos del memorando y del escrito.")
                                # --- FIN: (REQ 1) ---
                                
                                boton_habilitado = datos_generales_ok and datos_especificos_ok and datos_reduccion_ok

                                st.divider()
                                if st.button(f"Calcular Hecho {i+1}", key=f"calc_btn_{i}", disabled=(not boton_habilitado)):
                                    with st.spinner(f"Calculando hecho {i+1}..."):
                                        
                                        # --- VERSIÓN LIMPIA Y RÁPIDA SIN PLANTILLAS DE WORD ---
                                        acronym_manager = st.session_state.context_data.get('acronyms')
                                        
                                        # Preparamos los datos comunes para el cálculo matemático
                                        datos_comunes = {
                                            **datos_calculo, # Usamos todos los DFs precargados (UIT, COS, etc.)
                                            'datos_hecho_completos': st.session_state.imputaciones_data[i],
                                            'fecha_emision_informe': st.session_state.get('fecha_emision_informe', date.today()),
                                            'id_infraccion': id_infraccion,
                                            'rubro': st.session_state.rubro_seleccionado,
                                            'id_rubro_seleccionado': st.session_state.get('id_rubro_seleccionado'),
                                            'numero_hecho_actual': i + 1,
                                            'context_data': st.session_state.get('context_data', {}),
                                            'acronym_manager': acronym_manager
                                            # ELIMINAMOS 'doc_tpl' y toda la descarga de Drive.
                                        }

                                        resultados_completos = modulo_especialista.procesar_infraccion(
                                            datos_comunes, 
                                            st.session_state.imputaciones_data[i]
                                        )
                                        if resultados_completos.get('error'):
                                            st.error(f"Error en el cálculo del Hecho {i+1}: {resultados_completos['error']}")
                                        else:
                                            st.session_state.imputaciones_data[i]['resultados'] = resultados_completos
                                            st.session_state.imputaciones_data[i]['es_extemporaneo'] = resultados_completos.get('es_extemporaneo', False)
                                            st.session_state.imputaciones_data[i]['usa_capacitacion'] = resultados_completos.get('usa_capacitacion', False)
                                            st.success(f"Hecho {i+1} calculado.")
                                            st.session_state.imputaciones_data[i]['anexos_ce'] = resultados_completos.get('anexos_ce_generados', [])
                                            # --- NUEVA LÍNEA: Actualizar todos los hechos para reflejar el prorrateo ---
                                            actualizar_y_recalcular_prorrateos(cliente_gspread, NOMBRE_GSHEET_MAESTRO)
                                            
                                            st.success(f"Hecho {i+1} calculado y montos prorrateados actualizados.")
                                            st.rerun() # Forzar refresco para ver los cambios
                            except ImportError:
                                st.error(f"El módulo para '{id_infraccion}' no está implementado.")


                            # --- Sección para mostrar resultados ya calculados ---
                            if 'resultados' in st.session_state.imputaciones_data[i]:
                                resultados_app = st.session_state.imputaciones_data[i]['resultados'].get('resultados_para_app', {}) # Use .get for safety
                                id_infraccion_actual = st.session_state.imputaciones_data[i].get('id_infraccion', '')
                                st.subheader(f"Resultados del Cálculo para el Hecho {i + 1}")

                                totales_finales = {}

                                # Lógica específica para INF003
                                if 'INF003' in id_infraccion_actual or 'INF006' in id_infraccion_actual:
                                    if 'extremos' in resultados_app and len(resultados_app['extremos']) > 1:
                                        st.markdown("#### Desglose de Costo Evitado (CE) por Extremo")
                                        for j, extremo_data in enumerate(resultados_app['extremos']):
                                            st.markdown(f"##### Extremo {j + 1}: {extremo_data.get('tipo', 'N/A')}")
                                            df_ce_extremo = pd.DataFrame(extremo_data.get('ce_data', []))
                                            if not df_ce_extremo.empty:
                                                # Asumiendo que ce_data tiene las columnas correctas para INF003
                                                df_ce_extremo = df_ce_extremo.drop(columns=['id_anexo'], errors='ignore')
                                                df_ce_renamed = df_ce_extremo.rename(columns={
                                                    'descripcion': 'Descripción',
                                                    'precio_dolares': 'Precio Base (US$)', # Ajustar nombre si es diferente
                                                    'precio_soles': 'Precio Base (S/)',   # Ajustar nombre si es diferente
                                                    'factor_ajuste': 'Factor de Ajuste',
                                                    'monto_soles': 'Monto (S/)',
                                                    'monto_dolares': 'Monto (US$)'
                                                })
                                                numeric_cols_ce = ['Precio Base (US$)', 'Precio Base (S/)', 'Factor de Ajuste', 'Monto (S/)', 'Monto (US$)']
                                                for col in numeric_cols_ce:
                                                    if col in df_ce_renamed.columns: df_ce_renamed[col] = pd.to_numeric(df_ce_renamed[col], errors='coerce')
                                                st.dataframe(df_ce_renamed.style.format("{:,.3f}", subset=[c for c in numeric_cols_ce if c in df_ce_renamed.columns], na_rep='').hide(axis="index"), use_container_width=True)

                                        st.markdown("#### Beneficio Ilícito (BI) Consolidado")
                                        df_bi_total = pd.DataFrame(resultados_app.get('totales', {}).get('bi_data_raw', []))
                                        if not df_bi_total.empty:
                                            # Usar descripcion_texto si existe (para superíndices), sino 'descripcion'
                                            desc_col = 'descripcion_texto' if 'descripcion_texto' in df_bi_total.columns else 'descripcion'
                                            df_bi_display = df_bi_total.rename(columns={desc_col: 'Descripción', 'monto': 'Monto'})
                                            # Seleccionar solo las columnas a mostrar
                                            cols_to_show = ['Descripción', 'Monto']
                                            if 'descripcion_superindice' in df_bi_total.columns: # Añadir superíndice si existe
                                                df_bi_display['Descripción'] = df_bi_display['Descripción'] + df_bi_total['descripcion_superindice'].fillna('')
                                            st.dataframe(df_bi_display[cols_to_show].style.hide(axis="index"), use_container_width=True)

                                        totales_finales = resultados_app.get('totales', {})
                                    else:
                                        totales_finales = resultados_app # Caso simple INF003
                                        st.markdown("###### Costo Evitado (CE)")
                                        df_ce_total = pd.DataFrame(totales_finales.get('ce_data_raw', []))
                                        if not df_ce_total.empty:
                                            df_ce_total = df_ce_total.drop(columns=['id_anexo'], errors='ignore')
                                            df_ce_renamed = df_ce_total.rename(columns={
                                                'descripcion': 'Descripción',
                                                'precio_dolares': 'Precio Base (US$)', # Ajustar nombre si es diferente
                                                'precio_soles': 'Precio Base (S/)',   # Ajustar nombre si es diferente
                                                'factor_ajuste': 'Factor de Ajuste',
                                                'monto_soles': 'Monto (S/)',
                                                'monto_dolares': 'Monto (US$)'
                                            })
                                            numeric_cols_ce = ['Precio Base (US$)', 'Precio Base (S/)', 'Factor de Ajuste', 'Monto (S/)', 'Monto (US$)']
                                            for col in numeric_cols_ce:
                                                if col in df_ce_renamed.columns: df_ce_renamed[col] = pd.to_numeric(df_ce_renamed[col], errors='coerce')
                                            st.dataframe(df_ce_renamed.style.format("{:,.3f}", subset=[c for c in numeric_cols_ce if c in df_ce_renamed.columns], na_rep='').hide(axis="index"), use_container_width=True)

                                        st.markdown("###### Beneficio Ilícito (BI)")
                                        df_bi_total = pd.DataFrame(totales_finales.get('bi_data_raw', []))
                                        if not df_bi_total.empty:
                                            # Usar descripcion_texto si existe (para superíndices), sino 'descripcion'
                                            desc_col = 'descripcion_texto' if 'descripcion_texto' in df_bi_total.columns else 'descripcion'
                                            df_bi_display = df_bi_total.rename(columns={desc_col: 'Descripción', 'monto': 'Monto'})
                                            cols_to_show = ['Descripción', 'Monto']
                                            if 'descripcion_superindice' in df_bi_total.columns:
                                                df_bi_display['Descripción'] = df_bi_display['Descripción'] + df_bi_total['descripcion_superindice'].fillna('')
                                            st.dataframe(df_bi_display[cols_to_show].style.hide(axis="index"), use_container_width=True)

                                # Lógica específica para INF002
                                elif 'INF002' in id_infraccion_actual:
                                    if 'extremos' in resultados_app and isinstance(resultados_app['extremos'], list):
                                        st.markdown("#### Desglose por Extremo de Monitoreo")

                                        columnas_map_ce1 = {'descripcion': 'Descripción', 'horas': 'Horas', 'cantidad': 'Cantidad', 'precio_unitario': 'Precio Unitario (S/)', 'factor_ajuste': 'Factor Ajuste', 'monto_soles': 'Monto (S/)', 'monto_dolares': 'Monto (US$)'}
                                        columnas_map_ce2 = {'descripcion': 'Descripción', 'unidad': 'Unidad', 'cantidad': 'Cantidad', 'precio_unitario': 'Precio Unitario (S/)', 'factor_ajuste': 'Factor Ajuste', 'monto_soles': 'Monto (S/)', 'monto_dolares': 'Monto (US$)'}
                                        cols_numericas = ['Cantidad', 'Precio Unitario (S/)', 'Factor Ajuste', 'Monto (S/)', 'Monto (US$)', 'Horas']

                                        for j, extremo_data in enumerate(resultados_app['extremos']):
                                            st.markdown(f"##### Extremo {j + 1}: {extremo_data.get('tipo', 'N/A')}")

                                            if extremo_data.get('ce1_data'):
                                                st.markdown("###### Costo Evitado por Muestreo (CE1)")
                                                df_ce1 = pd.DataFrame(extremo_data['ce1_data']).rename(columns=columnas_map_ce1)
                                                cols_existentes = [col for col in cols_numericas if col in df_ce1.columns]
                                                st.dataframe(df_ce1.style.format("{:,.3f}", subset=cols_existentes, na_rep='').hide(axis="index"), use_container_width=True)

                                            if extremo_data.get('ce2_envio_data'):
                                                st.markdown("###### Costo Evitado por Envío de Muestras")
                                                df_ce2_envio = pd.DataFrame(extremo_data['ce2_envio_data']).rename(columns=columnas_map_ce2)
                                                cols_existentes = [col for col in cols_numericas if col in df_ce2_envio.columns]
                                                st.dataframe(df_ce2_envio.style.format("{:,.3f}", subset=cols_existentes, na_rep='').hide(axis="index"), use_container_width=True)

                                            if extremo_data.get('ce2_lab_data'):
                                                st.markdown("###### Costo Evitado por Análisis de Laboratorio")
                                                df_ce2_lab = pd.DataFrame(extremo_data['ce2_lab_data']).rename(columns=columnas_map_ce2)
                                                cols_existentes = [col for col in cols_numericas if col in df_ce2_lab.columns]
                                                st.dataframe(df_ce2_lab.style.format("{:,.3f}", subset=cols_existentes, na_rep='').hide(axis="index"), use_container_width=True)

                                            total_ce2_soles_extremo = extremo_data.get('ce2_envio_soles', 0) + extremo_data.get('ce2_lab_soles', 0)
                                            total_ce2_dolares_extremo = extremo_data.get('ce2_envio_dolares', 0) + extremo_data.get('ce2_lab_dolares', 0)

                                            if total_ce2_soles_extremo > 0:
                                                st.markdown("###### Resumen de Costo Evitado (CE2)")
                                                resumen_data_ce2 = []
                                                if extremo_data.get('ce2_envio_soles', 0) > 0:
                                                    resumen_data_ce2.append({'Componente': 'Subtotal Envío de Muestras', 'Monto (S/)': f"{extremo_data['ce2_envio_soles']:,.3f}", 'Monto (US$)': f"{extremo_data['ce2_envio_dolares']:,.3f}"})
                                                if extremo_data.get('ce2_lab_soles', 0) > 0:
                                                    resumen_data_ce2.append({'Componente': 'Subtotal Análisis de Laboratorio', 'Monto (S/)': f"{extremo_data['ce2_lab_soles']:,.3f}", 'Monto (US$)': f"{extremo_data['ce2_lab_dolares']:,.3f}"})
                                                resumen_data_ce2.append({'Componente': 'Total Costo Evitado (CE2)', 'Monto (S/)': f"{total_ce2_soles_extremo:,.3f}", 'Monto (US$)': f"{total_ce2_dolares_extremo:,.3f}"})
                                                st.dataframe(pd.DataFrame(resumen_data_ce2).style.hide(axis="index"), use_container_width=True)

                                            st.markdown("###### Resumen Total del Costo Evitado del Extremo")
                                            resumen_data_total = [
                                                {'Componente': 'Total Costo Evitado (CE1)', 'Monto (S/)': f"{extremo_data.get('ce1_soles', 0):,.3f}", 'Monto (US$)': f"{extremo_data.get('ce1_dolares', 0):,.3f}"},
                                                {'Componente': 'Total Costo Evitado (CE2)', 'Monto (S/)': f"{total_ce2_soles_extremo:,.3f}", 'Monto (US$)': f"{total_ce2_dolares_extremo:,.3f}"},
                                                {'Componente': 'Costo Evitado Total del Extremo', 'Monto (S/)': f"{extremo_data.get('total_soles_extremo', 0):,.3f}", 'Monto (US$)': f"{extremo_data.get('total_dolares_extremo', 0):,.3f}"}
                                            ]
                                            st.dataframe(pd.DataFrame(resumen_data_total).style.hide(axis="index"), use_container_width=True)

                                        totales_finales = resultados_app.get('totales', {})
                                        st.markdown("--- \n#### Totales Consolidados del Hecho")
                                        st.markdown("###### Beneficio Ilícito (BI)")
                                        df_bi_total = pd.DataFrame(totales_finales.get('bi_data_raw', []))
                                        if not df_bi_total.empty:
                                            # Usar descripcion_texto si existe (para superíndices), sino 'descripcion'
                                            desc_col = 'descripcion_texto' if 'descripcion_texto' in df_bi_total.columns else 'descripcion'
                                            df_bi_display = df_bi_total.rename(columns={desc_col: 'Descripción', 'monto': 'Monto'})
                                            cols_to_show = ['Descripción', 'Monto']
                                            if 'descripcion_superindice' in df_bi_total.columns:
                                                df_bi_display['Descripción'] = df_bi_display['Descripción'] + df_bi_total['descripcion_superindice'].fillna('')
                                            st.dataframe(df_bi_display[cols_to_show].style.hide(axis="index"), use_container_width=True)

                                # --- Lógica específica para INF004 ---
                                elif 'INF004' in id_infraccion_actual  or 'INF009' in id_infraccion_actual or 'INF010' in id_infraccion_actual:
                                    if 'extremos' in resultados_app and isinstance(resultados_app['extremos'], list):
                                        # --- Lógica para Múltiples Extremos (INF004) ---
                                        st.markdown("#### Desglose por Extremo")
                                        # Definir mapeo COMPLETO de columnas para INF004
                                        columnas_map_inf004 = {
                                            'descripcion': 'Descripción',
                                            'cantidad': 'Cantidad',
                                            'horas': 'Horas',
                                            'precio_soles': 'Precio asociado (S/)',
                                            'factor_ajuste': 'Factor Ajuste',
                                            'monto_soles': 'Monto (S/)',
                                            'monto_dolares': 'Monto (US$)'
                                        }
                                        # Columnas que SÍ deben tener 3 decimales fijos
                                        cols_formato_numerico = ['Precio asociado (S/)', 'Factor Ajuste', 'Monto (S/)', 'Monto (US$)']

                                        for j, extremo_data in enumerate(resultados_app['extremos']):
                                            st.markdown(f"##### Extremo {j + 1}: {extremo_data.get('tipo', 'N/A')}")
                                            st.markdown("###### Costo Evitado (CE)")
                                            datos_ce_crudos = extremo_data.get('ce_data', [])
                                            if datos_ce_crudos:
                                                df_ce_display = pd.DataFrame(datos_ce_crudos)

                                                # --- INICIO Req 2: Formato dinámico ---
                                                if 'cantidad' in df_ce_display.columns:
                                                    df_ce_display['cantidad'] = df_ce_display['cantidad'].apply(format_decimal_dinamico)
                                                if 'horas' in df_ce_display.columns:
                                                    df_ce_display['horas'] = df_ce_display['horas'].apply(format_decimal_dinamico)
                                                # --- FIN Req 2 ---
                                                
                                                # --- INICIO: AÑADIR FILA TOTAL ---
                                                total_monto_soles = sum(item.get('monto_soles', 0) for item in datos_ce_crudos)
                                                total_monto_dolares = sum(item.get('monto_dolares', 0) for item in datos_ce_crudos)
                                                total_df = pd.DataFrame([{'descripcion': 'Total', 'monto_soles': total_monto_soles, 'monto_dolares': total_monto_dolares}])
                                                df_ce_display = pd.concat([df_ce_display, total_df], ignore_index=True)
                                                # --- FIN: AÑADIR FILA TOTAL ---
                                                
                                                # --- INICIO Req 3: Arreglar tabla deformada ---
                                                columnas_a_mostrar = [col for col in columnas_map_inf004.keys() if col in df_ce_display.columns]
                                                df_ce_display = df_ce_display[columnas_a_mostrar].rename(columns=columnas_map_inf004)
                                                # Reindexar para asegurar todas las columnas y evitar deformación
                                                df_ce_display = df_ce_display.reindex(columns=columnas_map_inf004.values())
                                                # --- FIN Req 3 ---

                                                # Formatear columnas numéricas (moneda/factor)
                                                cols_numericas_existentes = [col for col in cols_formato_numerico if col in df_ce_display.columns]
                                                st.dataframe(
                                                    df_ce_display.style.format("{:,.3f}", subset=cols_numericas_existentes, na_rep='').hide(axis="index"), 
                                                    use_container_width=True
                                                )
                                            else:
                                                st.warning("No hay columnas válidas para mostrar en el Costo Evitado.")
                                            
                                            # Mostrar BI del extremo (si existe en los datos)
                                            st.markdown("###### Beneficio Ilícito (BI) del Extremo")
                                            df_bi_extremo = pd.DataFrame(extremo_data.get('bi_data', []))
                                            if not df_bi_extremo.empty:
                                                # Usar descripcion_texto si existe, sino 'descripcion'
                                                desc_col = 'descripcion_texto' if 'descripcion_texto' in df_bi_extremo.columns else 'descripcion'
                                                df_bi_display = df_bi_extremo.rename(columns={desc_col: 'Descripción', 'monto': 'Monto'})
                                                cols_to_show = ['Descripción', 'Monto']
                                                
                                                # --- INICIO DE LA CORRECCIÓN ---
                                                # 1. Verificar si 'descripcion_superindice' existe en el DF *renombrado*
                                                if 'descripcion_superindice' in df_bi_display.columns:
                                                    # 2. Usar las columnas del DF *renombrado* para la operación
                                                    df_bi_display['Descripción'] = df_bi_display['Descripción'] + df_bi_display['descripcion_superindice'].fillna('')
                                                
                                                # 3. Asegurarse de que las columnas existan ANTES de seleccionarlas
                                                cols_finales = [col for col in cols_to_show if col in df_bi_display.columns]
                                                st.dataframe(df_bi_display[cols_finales].style.hide(axis="index"), use_container_width=True)
                                                # --- FIN DE LA CORRECCIÓN ---
                                            else:
                                                st.warning("No hay datos de Beneficio Ilícito para este extremo.")

                                        totales_finales = resultados_app.get('totales', {}) # Obtener totales consolidados

                                    else:
                                    # --- Lógica para Hecho Simple (INF004) ---
                                        totales_finales = resultados_app 
                                        st.markdown("###### Costo Evitado (CE)")
                                        datos_ce_crudos = totales_finales.get('ce_data_raw', [])
                                        if datos_ce_crudos:
                                            df_ce_display = pd.DataFrame(datos_ce_crudos)
                                            # Definir mapeo COMPLETO de columnas para INF004
                                            columnas_map_inf004 = {
                                                'descripcion': 'Descripción',
                                                'cantidad': 'Cantidad',
                                                'horas': 'Horas',
                                                'precio_soles': 'Precio asociado (S/)',
                                                'factor_ajuste': 'Factor Ajuste',
                                                'monto_soles': 'Monto (S/)',
                                                'monto_dolares': 'Monto (US$)'
                                            }
                                            # Columnas que SÍ deben tener 3 decimales fijos
                                            cols_formato_numerico = ['Precio asociado (S/)', 'Factor Ajuste', 'Monto (S/)', 'Monto (US$)']

                                            # --- INICIO Req 2: Formato dinámico ---
                                            if 'cantidad' in df_ce_display.columns:
                                                df_ce_display['cantidad'] = df_ce_display['cantidad'].apply(format_decimal_dinamico)
                                            if 'horas' in df_ce_display.columns:
                                                df_ce_display['horas'] = df_ce_display['horas'].apply(format_decimal_dinamico)
                                            # --- FIN Req 2 ---
                                            
                                            # --- INICIO: AÑADIR FILA TOTAL ---
                                            total_monto_soles = sum(item.get('monto_soles', 0) for item in datos_ce_crudos)
                                            total_monto_dolares = sum(item.get('monto_dolares', 0) for item in datos_ce_crudos)
                                            total_df = pd.DataFrame([{'descripcion': 'Total', 'monto_soles': total_monto_soles, 'monto_dolares': total_monto_dolares}])
                                            df_ce_display = pd.concat([df_ce_display, total_df], ignore_index=True)
                                            # --- FIN: AÑADIR FILA TOTAL ---
                                            
                                            # --- INICIO Req 3: Arreglar tabla deformada ---
                                            columnas_a_mostrar = [col for col in columnas_map_inf004.keys() if col in df_ce_display.columns]
                                            df_ce_display = df_ce_display[columnas_a_mostrar].rename(columns=columnas_map_inf004)
                                            # Reindexar para asegurar todas las columnas y evitar deformación
                                            df_ce_display = df_ce_display.reindex(columns=columnas_map_inf004.values())
                                            # --- FIN Req 3 ---

                                            # Formatear columnas numéricas (moneda/factor)
                                            cols_numericas_existentes = [col for col in cols_formato_numerico if col in df_ce_display.columns]
                                            st.dataframe(
                                                df_ce_display.style.format("{:,.3f}", subset=cols_numericas_existentes, na_rep='').hide(axis="index"), 
                                                use_container_width=True
                                            )
                                        else:
                                            st.warning("No hay columnas válidas para mostrar en el Costo Evitado.")

                                        st.markdown("###### Beneficio Ilícito (BI)")
                                        df_bi_total = pd.DataFrame(totales_finales.get('bi_data_raw', []))
                                        if not df_bi_total.empty:
                                            # Usar descripcion_texto si existe, sino 'descripcion'
                                            desc_col = 'descripcion_texto' if 'descripcion_texto' in df_bi_total.columns else 'descripcion'
                                            df_bi_display = df_bi_total.rename(columns={desc_col: 'Descripción', 'monto': 'Monto'})
                                            cols_to_show = ['Descripción', 'Monto']

                                            # --- INICIO DE LA CORRECCIÓN ---
                                            # 1. Verificar si 'descripcion_superindice' existe en el DF *renombrado*
                                            if 'descripcion_superindice' in df_bi_display.columns:
                                                # 2. Usar las columnas del DF *renombrado* para la operación
                                                df_bi_display['Descripción'] = df_bi_display['Descripción'] + df_bi_display['descripcion_superindice'].fillna('')
                                            
                                            # 3. Asegurarse de que las columnas existan ANTES de seleccionarlas
                                            cols_finales = [col for col in cols_to_show if col in df_bi_display.columns]
                                            st.dataframe(df_bi_display[cols_finales].style.hide(axis="index"), use_container_width=True)
                                            # --- FIN DE LA CORRECCIÓN ---

                                # --- REVISED BLOCK FOR INF005 ---
                                elif any(inf in id_infraccion_actual for inf in ['INF001', 'INF005', 'INF007', 'INF008', 'INF009', 'INF011']):
                                    st.markdown("#### Desglose por Extremo")

                                    # Define column mappings and numeric columns
                                    # --- INICIO CORRECCIÓN 2: Columnas CE1 ---
                                    columnas_map_ce1 = {'descripcion': 'Descripción', 'cantidad': 'Cantidad', 'horas': 'Horas', 'precio_soles': 'Precio Base (S/)', 'factor_ajuste': 'Factor Ajuste', 'monto_soles': 'Monto (S/)', 'monto_dolares': 'Monto (US$)'}
                                    cols_num_ce1 = ['Cantidad', 'Horas', 'Precio Base (S/)', 'Factor Ajuste', 'Monto (S/)', 'Monto (US$)']
                                    # --- FIN CORRECCIÓN 2 ---
                                    
                                    columnas_map_ce2 = {'descripcion': 'Descripción', 'precio_soles': 'Precio Base (S/)','precio_dolares': 'Precio Base (US$)', 'factor_ajuste': 'Factor Ajuste', 'monto_soles': 'Monto (S/)','monto_dolares': 'Monto (US$)'}
                                    cols_num_ce2 = ['Precio Base (S/)', 'Precio Base (US$)', 'Factor Ajuste', 'Monto (S/)', 'Monto (US$)']

                                    # --- Data Handling for Simple vs Multiple Extremes ---
                                    if 'extremos' in resultados_app and isinstance(resultados_app['extremos'], list):
                                        extremos_a_mostrar = resultados_app['extremos']
                                        totales_finales = resultados_app.get('totales', {}) # Obtener totales consolidados
                                    else:
                                        extremos_a_mostrar = [resultados_app]
                                        if 'ce1_total_soles' in resultados_app: extremos_a_mostrar[0]['ce1_soles'] = resultados_app.get('ce1_total_soles', 0); extremos_a_mostrar[0]['ce1_dolares'] = resultados_app.get('ce1_total_dolares', 0)
                                        if 'ce2_total_soles_calculado' in resultados_app: extremos_a_mostrar[0]['ce2_soles_calculado'] = resultados_app.get('ce2_total_soles_calculado', 0); extremos_a_mostrar[0]['ce2_dolares_calculado'] = resultados_app.get('ce2_total_dolares_calculado', 0)
                                        if 'ce_total_soles_para_bi' in resultados_app: extremos_a_mostrar[0]['ce_soles_para_bi'] = resultados_app.get('ce_total_soles_para_bi', 0); extremos_a_mostrar[0]['ce_dolares_para_bi'] = resultados_app.get('ce_dolares_para_bi', 0)
                                        if 'aplicar_ce2_a_bi' not in extremos_a_mostrar[0]:
                                            totales_dict = resultados_app.get('totales', resultados_app) 
                                            extremos_a_mostrar[0]['aplicar_ce2_a_bi'] = totales_dict.get('aplicar_ce2_a_bi', False) 
                                        if 'ce1_data_raw' in resultados_app: extremos_a_mostrar[0]['ce1_data'] = resultados_app.get('ce1_data_raw', [])
                                        if 'ce2_data_raw' in resultados_app: extremos_a_mostrar[0]['ce2_data'] = resultados_app.get('ce2_data_raw', [])
                                        if 'bi_data_raw' in resultados_app:
                                            totales_dict = resultados_app.get('totales', resultados_app)
                                            extremos_a_mostrar[0]['bi_data'] = totales_dict.get('bi_data_raw', [])
                                        if 'tipo' not in extremos_a_mostrar[0]: extremos_a_mostrar[0]['tipo'] = "Incumplimiento Único"
                                        
                                        totales_finales = resultados_app.get('totales', resultados_app) # Totales para caso simple
                                    # --- End Data Handling ---

                                    # --- Loop through each extreme ---
                                    for j, extremo_data in enumerate(extremos_a_mostrar):
                                        st.markdown(f"##### Extremo {j + 1}: {extremo_data.get('tipo', 'N/A')}")
                                        aplicar_ce2_bi_extremo = extremo_data.get('aplicar_ce2_a_bi', False)
                                        datos_ce1_crudos = extremo_data.get('ce1_data', [])
                                        datos_ce2_crudos = extremo_data.get('ce2_data', [])

                                        st.markdown("###### CE1: Remisión")
                                        if datos_ce1_crudos:
                                            df_ce1_display = pd.DataFrame(datos_ce1_crudos)
                                            # --- INICIO CORRECCIÓN 2: Añadir Total CE1 ---
                                            total_ce1_s = df_ce1_display['monto_soles'].sum()
                                            total_ce1_d = df_ce1_display['monto_dolares'].sum()
                                            df_ce1_display = pd.concat([df_ce1_display, pd.DataFrame([{'descripcion': 'Total', 'monto_soles': total_ce1_s, 'monto_dolares': total_ce1_d}])], ignore_index=True)
                                            # --- FIN CORRECCIÓN 2 ---
                                            cols_exist_ce1 = [col for col in columnas_map_ce1.keys() if col in df_ce1_display.columns]
                                            if cols_exist_ce1:
                                                df_ce1_display = df_ce1_display[cols_exist_ce1].rename(columns=columnas_map_ce1)
                                                # --- INICIO CORRECCIÓN 2: Formato CE1 ---
                                                df_ce1_display = df_ce1_display.reindex(columns=columnas_map_ce1.values()) # Asegurar todas las columnas
                                                cols_num_exist_ce1 = [col for col in cols_num_ce1 if col in df_ce1_display.columns]
                                                st.dataframe(df_ce1_display.style.format("{:,.3f}", subset=cols_num_exist_ce1, na_rep='').hide(axis="index"), use_container_width=True)
                                                # --- FIN CORRECCIÓN 2 ---
                                            else: st.warning("No hay columnas válidas para mostrar en CE1.")
                                        else:
                                            st.warning("No hay datos de CE1 para este extremo.")

                                        if aplicar_ce2_bi_extremo and datos_ce2_crudos:
                                            st.markdown("###### CE2: Capacitación")
                                            df_ce2_display = pd.DataFrame(datos_ce2_crudos)
                                            # --- INICIO CORRECCIÓN 2: Añadir Total CE2 ---
                                            total_ce2_s = df_ce2_display['monto_soles'].sum()
                                            total_ce2_d = df_ce2_display['monto_dolares'].sum()
                                            df_ce2_display = pd.concat([df_ce2_display, pd.DataFrame([{'descripcion': 'Total', 'monto_soles': total_ce2_s, 'monto_dolares': total_ce2_d}])], ignore_index=True)
                                            # --- FIN CORRECCIÓN 2 ---
                                            cols_exist_ce2 = [col for col in columnas_map_ce2.keys() if col in df_ce2_display.columns]
                                            if cols_exist_ce2:
                                                df_ce2_display = df_ce2_display[cols_exist_ce2].rename(columns=columnas_map_ce2)
                                                # --- INICIO CORRECCIÓN 2: Formato CE2 ---
                                                df_ce2_display = df_ce2_display.reindex(columns=columnas_map_ce2.values()) # Asegurar todas las columnas
                                                cols_num_exist_ce2 = [col for col in cols_num_ce2 if col in df_ce2_display.columns]
                                                st.dataframe(df_ce2_display.style.format("{:,.3f}", subset=cols_num_exist_ce2, na_rep='').hide(axis="index"), use_container_width=True)
                                                # --- FIN CORRECCIÓN 2 ---
                                            else: st.warning("No hay columnas válidas para mostrar en CE2.")

                                        if aplicar_ce2_bi_extremo:
                                            st.markdown("###### Resumen CE del Extremo")
                                            resumen_ext_data = []
                                            ce1_s = extremo_data.get('ce1_soles', 0); ce1_d = extremo_data.get('ce1_dolares', 0)
                                            ce2_s_calc = extremo_data.get('ce2_soles_calculado', 0); ce2_d_calc = extremo_data.get('ce2_dolares_calculado', 0)
                                            total_ce_s = ce1_s + ce2_s_calc; total_ce_d = ce1_d + ce2_d_calc
                                            if ce1_s > 0 or ce1_d > 0: resumen_ext_data.append({'Componente': 'Subtotal Remisión (CE1)', 'Monto (S/)': f"{ce1_s:,.3f}", 'Monto (US$)': f"{ce1_d:,.3f}"})
                                            if ce2_s_calc > 0 or ce2_d_calc > 0: resumen_ext_data.append({'Componente': 'Subtotal Capacitación (CE2)', 'Monto (S/)': f"{ce2_s_calc:,.3f}", 'Monto (US$)': f"{ce2_d_calc:,.3f}"})
                                            resumen_ext_data.append({'Componente': 'Total CE Calculado (Extremo)', 'Monto (S/)': f"{total_ce_s:,.3f}", 'Monto (US$)': f"{total_ce_d:,.3f}"})

                                            if resumen_ext_data: 
                                                # --- INICIO CORRECCIÓN 2: Mostrar US$ en Resumen ---
                                                st.dataframe(pd.DataFrame(resumen_ext_data)[['Componente', 'Monto (S/)', 'Monto (US$)']].style.hide(axis="index"), use_container_width=True)
                                                # --- FIN CORRECCIÓN 2 ---

                                        st.markdown("###### Beneficio Ilícito (BI) del Extremo")
                                        df_bi_extremo = pd.DataFrame(extremo_data.get('bi_data', []))
                                        if not df_bi_extremo.empty:
                                            desc_col_bi = 'descripcion_texto' if 'descripcion_texto' in df_bi_extremo.columns else 'descripcion'
                                            df_bi_display_ext = df_bi_extremo.rename(columns={desc_col_bi: 'Descripción', 'monto': 'Monto'})
                                            cols_show_bi = ['Descripción', 'Monto']
                                            if 'descripcion_superindice' in df_bi_display_ext.columns: # <--- Usar df_bi_display_ext
                                                df_bi_display_ext['Descripción'] = df_bi_display_ext['Descripción'] + df_bi_display_ext['descripcion_superindice'].fillna('')
                                            st.dataframe(df_bi_display_ext[cols_show_bi].style.hide(axis="index"), use_container_width=True)
                                        else:
                                            st.warning("No hay datos de Beneficio Ilícito para mostrar para este extremo.")
                                        st.markdown("---") # Separator between extremes

                                # Lógica general para otras infracciones (sin desglose de extremos específico)
                                else:
                                    totales_finales = resultados_app # Asumir estructura simple
                                    st.markdown("###### Costo Evitado (CE)")
                                    datos_ce_crudos = totales_finales.get('ce_data_raw', [])
                                    if datos_ce_crudos:
                                        df_ce_display = pd.DataFrame(datos_ce_crudos)
                                        # Mapeo genérico (puede necesitar ajustes por infracción)
                                        columnas_map_generico = {
                                            'grupo': 'Grupo', 'subgrupo': 'Subgrupo', 'descripcion': 'Descripción',
                                            'unidad': 'Unidad', 'cantidad': 'Cantidad',
                                            'precio_unitario': 'Precio Unitario (S/)', # Asumiendo Soles
                                            'factor_ajuste': 'Factor Ajuste', 'monto_soles': 'Monto (S/)',
                                            'monto_dolares': 'Monto (US$)'
                                        }
                                        columnas_existentes = [col for col in columnas_map_generico.keys() if col in df_ce_display.columns]
                                        if columnas_existentes:
                                            df_ce_display = df_ce_display[columnas_existentes].rename(columns=columnas_map_generico)
                                            cols_numericas_generico = ['Cantidad', 'Precio Unitario (S/)', 'Factor Ajuste', 'Monto (S/)', 'Monto (US$)']
                                            cols_numericas_existentes = [col for col in cols_numericas_generico if col in df_ce_display.columns]
                                            st.dataframe(df_ce_display.style.format("{:,.3f}", subset=cols_numericas_existentes, na_rep='').hide(axis="index"), use_container_width=True)
                                        else:
                                            st.warning("No hay columnas válidas para mostrar en el Costo Evitado.")
                                    else:
                                        st.warning("No hay datos de Costo Evitado para mostrar.")

                                    st.markdown("###### Beneficio Ilícito (BI)")
                                    df_bi_total = pd.DataFrame(totales_finales.get('bi_data_raw', []))
                                    if not df_bi_total.empty:
                                        desc_col = 'descripcion_texto' if 'descripcion_texto' in df_bi_total.columns else 'descripcion'
                                        df_bi_display = df_bi_total.rename(columns={desc_col: 'Descripción', 'monto': 'Monto'})
                                        cols_to_show = ['Descripción', 'Monto']
                                        if 'descripcion_superindice' in df_bi_total.columns:
                                            df_bi_display['Descripción'] = df_bi_display['Descripción'] + df_bi_total['descripcion_superindice'].fillna('')
                                        st.dataframe(df_bi_display[cols_to_show].style.hide(axis="index"), use_container_width=True)

                                # --- SECCIÓN DE TOTALES Y MULTA (COMÚN Y SEGURA) ---
                                if totales_finales:
                                    st.markdown("---")
                                    st.markdown("#### Totales del Hecho Imputado")

                                    bi_uit_final = totales_finales.get('beneficio_ilicito_uit', 0)
                                    st.metric("Beneficio Ilícito Total (UIT)", f"{bi_uit_final:,.3f}")

                                    st.markdown("###### Multa Propuesta")
                                    df_multa = pd.DataFrame(totales_finales.get('multa_data_raw', []))
                                    if not df_multa.empty:
                                        df_multa_display = df_multa.rename(columns={'Componentes': 'Componentes', 'Monto': 'Monto'})
                                        st.dataframe(df_multa_display.style.hide(axis="index"), use_container_width=True)
                                    
                                    # --- INICIO: (REQ 1) MOSTRAR REDUCCIÓN EN UI (CORREGIDO) ---
                                    # Esta lógica ahora está fuera del 'if' de 'INF004' 
                                    # y se aplica a TODAS las infracciones.
                                    if totales_finales.get('aplica_reduccion') == 'Sí':
                                        porcentaje = totales_finales.get('porcentaje_reduccion', 'N/A')
                                        # Usamos la clave 'multa_con_reduccion_uit' que es la multa intermedia
                                        multa_con_reduccion = totales_finales.get('multa_con_reduccion_uit', 0) 
                                        st.info(f"Se aplica reducción del **{porcentaje}**.")
                                        st.metric("Multa con Reducción (UIT)", f"{multa_con_reduccion:,.3f}")
                                    # --- FIN: (REQ 1) ---

                # --- FIN DEL PASO 3 ---
                
                # --- INICIO: PASO 3.5: ANÁLISIS DE NO CONFISCATORIEDAD (CORREGIDO) ---
                all_steps_complete_check = all('resultados' in d for d in st.session_state.imputaciones_data)
                
                if all_steps_complete_check:
                    st.divider()
                    st.header("Análisis de no confiscatoriedad")
                    
                    if 'confiscatoriedad' not in st.session_state:
                        st.session_state.confiscatoriedad = {'aplica': 'No', 'datos_por_anio': {}}

                    aplica_conf = st.radio(
                        "¿El administrado presentó sus ingresos para el análisis de no confiscatoriedad?",
                        ["No", "Sí"],
                        key='confiscatoriedad_aplica',
                        index=0 if st.session_state.confiscatoriedad.get('aplica', 'No') == 'No' else 1
                    )
                    st.session_state.confiscatoriedad['aplica'] = aplica_conf

                    # --- INICIO CORRECCIÓN NameError (Paso 3.5) ---

                    if aplica_conf == 'Sí':
                        st.info("Sume los montos (Ventas Netas + Otros Ingresos Gravados + Otros Ingresos No Gravados) del año anterior a cada infracción.")
                        
                        # --- INICIO DE LA CORRECCIÓN (Req. 1) ---
                        # Pedir los datos del escrito UNA SOLA VEZ, fuera del bucle
                        
                        datos_conf_global = st.session_state.confiscatoriedad # Referencia al dict
                        
                        st.markdown("##### Documento de Acreditación (Único)")
                        col_e1, col_e2 = st.columns(2)
                        with col_e1:
                            # Guardar en el nivel superior de 'confiscatoriedad'
                            datos_conf_global['escrito_num_conf'] = st.text_input(
                                "N.° de Escrito (Ingresos)",
                                value=datos_conf_global.get('escrito_num_conf', ''),
                                key='conf_escrito_num_global_input'
                            )
                        with col_e2:
                            # Guardar en el nivel superior de 'confiscatoriedad'
                            datos_conf_global['escrito_fecha_conf'] = st.date_input(
                                "Fecha del Escrito (Ingresos)",
                                value=datos_conf_global.get('escrito_fecha_conf'),
                                key='conf_escrito_fecha_global_input',
                                format="DD/MM/YYYY"
                            )
                        # --- FIN DE LA CORRECCIÓN (Req. 1) ---

                        # 1. Identificar todos los años de incumplimiento
                        anios_incumplimiento = set()
                        
                        # --- INICIO DE LA CORRECCIÓN ---
                        # La lógica anterior solo buscaba 'fecha_incumplimiento' (usada por INF008).
                        # Ahora buscamos todas las claves posibles (ej: 'fecha_incumplimiento_extremo' de INF004).
                        
                        for i, datos_hecho in enumerate(st.session_state.imputaciones_data):
                            # Asegurarnos que el hecho tenga extremos antes de iterar
                            extremos_del_hecho = datos_hecho.get('extremos', [])
                            if not extremos_del_hecho:
                                st.warning(f"Hecho {i+1} no tiene extremos. Saltando para análisis de confiscatoriedad.")
                                continue
                                
                            for j, extremo in enumerate(extremos_del_hecho):
                                
                                # Clave 1: Usada por INF008 (y otros)
                                fecha_inc = extremo.get('fecha_incumplimiento') 
                                
                                # Clave 2: Usada por INF004
                                if not fecha_inc:
                                    fecha_inc = extremo.get('fecha_incumplimiento_extremo')
                                
                                # (Puedes añadir más 'elif not fecha_inc:' si otros módulos usan nombres diferentes)

                                if fecha_inc:
                                    try:
                                        # Asegurarnos que sea un objeto 'date' o 'datetime'
                                        anios_incumplimiento.add(fecha_inc.year)
                                    except AttributeError:
                                        st.warning(f"Hecho {i+1}, Extremo {j+1}: La fecha encontrada no es un objeto de fecha válido.")
                                else:
                                    # Si no encontramos NINGUNA clave de fecha
                                    st.warning(f"Hecho {i+1}, Extremo {j+1}: No se pudo encontrar una clave de fecha de incumplimiento ('fecha_incumplimiento' o 'fecha_incumplimiento_extremo').")

                        # --- FIN DE LA CORRECCIÓN ---
                        
                        anios_ordenados = sorted(list(anios_incumplimiento))
                        
                        if not anios_ordenados:
                            # Este error solo saltará si DE VERDAD no se encontró ninguna fecha en ningún hecho
                            st.error("Error: No se pudieron determinar los años de incumplimiento de ningún hecho.")
                        
                        # 2. Pedir ingresos para cada grupo de años
                        datos_por_anio_guardados = st.session_state.confiscatoriedad.get('datos_por_anio', {})
                        
                        for anio_incumplimiento in anios_ordenados:
                            anio_ingresos = anio_incumplimiento - 1
                            st.markdown(f"--- \n**Datos para incumplimientos del año {anio_incumplimiento} (se usarán ingresos de {anio_ingresos}):**")
                            
                            key_s = f"conf_soles_{anio_incumplimiento}"
                            key_a = f"conf_anio_{anio_incumplimiento}"
                            
                            datos_guardados_anio = datos_por_anio_guardados.get(anio_incumplimiento, {})
                            
                            # --- INICIO DE LA CORRECCIÓN (Req. 1) ---
                            # Ahora solo hay 2 columnas en el bucle
                            col_c1, col_c2 = st.columns(2) 
                            
                            with col_c1:
                                ingreso_total_soles = st.number_input(
                                    f"Ingreso Bruto Total {anio_ingresos} (S/)", 
                                    min_value=0.0, 
                                    value=datos_guardados_anio.get('ingreso_total_soles', 0.0),
                                    key=key_s,
                                    format="%.2f"
                                )
                            with col_c2:
                                anio_uit = st.number_input(
                                    f"Año de UIT (para ingresos {anio_ingresos})", 
                                    min_value=2000, 
                                    max_value=date.today().year, 
                                    value=datos_guardados_anio.get('anio_ingresos', anio_ingresos),
                                    key=key_a
                                )
                            
                            # Guardar los datos en el estado de sesión
                            if anio_incumplimiento not in st.session_state.confiscatoriedad['datos_por_anio']:
                                st.session_state.confiscatoriedad['datos_por_anio'][anio_incumplimiento] = {}
                            
                            st.session_state.confiscatoriedad['datos_por_anio'][anio_incumplimiento]['ingreso_total_soles'] = ingreso_total_soles
                            st.session_state.confiscatoriedad['datos_por_anio'][anio_incumplimiento]['anio_ingresos'] = anio_uit
                            # (Se elimina el guardado del N°/Fecha de escrito aquí)
                            # --- FIN DE LA CORRECCIÓN (Req. 1) ---

                            # --- SISTEMA DE GUARDADO UBICADO AL FINAL DE LOS INPUTS ---
                st.divider()
                col_save_title, col_save_btn = st.columns([2, 1])
                with col_save_title:
                    st.markdown("### 💾 Guardar Todo el Avance")
                    st.caption("Guarda hechos, graduación, reducciones y análisis de confiscatoriedad.")
                
                with col_save_btn:
                    if st.button("Guardar Caso en la Nube", type="secondary", use_container_width=True):
                        with st.spinner("Sincronizando con la base de datos..."):
                            # Lista maestra de claves a persistir
                            claves_sesion = [
                                'fecha_emision_informe', 'numero_rsd_base', 'fecha_rsd',
                                'confiscatoriedad', 'rubro_seleccionado', 'imputaciones_data',
                                'numero_ifi', 'fecha_ifi', 'num_informe_multa_ifi',
                                'monto_multa_ifi', 'num_imputaciones_ifi'
                            ]

                            # Limpiar y serializar
                            estado_sucio = {k: st.session_state[k] for k in claves_sesion if k in st.session_state}
                            estado_a_guardar = preparar_datos_para_json(estado_sucio)

                            expediente = st.session_state.num_expediente_formateado
                            producto = st.session_state.info_expediente.get('PRODUCTO', 'IFI')

                            exito, mensaje = guardar_datos_caso(cliente_gspread, expediente, producto, estado_a_guardar)

                        if exito:
                            st.success(f"✅ {mensaje}")
                        else:
                            st.error(f"❌ {mensaje}")
                                        
    # --- PASO 4: GENERAR INFORME FINAL ---
    all_steps_complete = False
    if 'imputaciones_data' in st.session_state and st.session_state.imputaciones_data:
        all_steps_complete = all('resultados' in d for d in st.session_state.imputaciones_data)

    if all_steps_complete:
            st.divider()
            st.header("Paso 4: Generar Informe Final")
            if st.button("🚀 Generar Informe", type="primary"):

                # --- CAMBIO PUNTUAL: Resetear Acrónimos ---
                if 'context_data' in st.session_state and 'acronyms' in st.session_state.context_data:
                    st.session_state.context_data['acronyms'].defined_acronyms = set()
                # ------------------------------------------

                with st.status("Generando informe con el nuevo motor...", expanded=True) as status:
                    try:
                        # 1. Recopilamos la información directamente del estado de la sesión
                        info_exp = st.session_state.info_expediente
                        ctx_data = st.session_state.context_data
                        imputaciones = st.session_state.imputaciones_data
                        confiscatoriedad = st.session_state.get('confiscatoriedad', {})
                        
                        status.update(label="Construyendo estructura del documento...")
                        
                        # 2. Llamamos a ifi.py, pasándole todos los datos pesados
                        nombre_archivo = generar_documento_ifi(info_exp, ctx_data, imputaciones, confiscatoriedad)
                        
                        status.update(label="¡Informe generado con éxito!", state="complete", expanded=False)

                        # 3. Botón de descarga leyendo el archivo guardado
                        nombre_descarga = f"IFI_{info_exp.get('ADMINISTRADO', 'Empresa')}.docx".replace(" ", "_")
                        
                        with open(nombre_archivo, "rb") as file:
                            st.download_button(
                                label="✅ Descargar Informe IFI Nuevo (.docx)",
                                data=file,
                                file_name=nombre_descarga,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                type="primary"
                            )

                    except Exception as e:
                        st.error(f"Ocurrió un error al generar el documento: {e}")
                        import traceback
                        st.exception(e)

if not cliente_gspread:
    st.error(
        "🔴 No se pudo establecer la conexión con Google Sheets. Revisa el archivo de credenciales y la conexión a internet.")