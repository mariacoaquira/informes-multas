import io
import gspread
import streamlit as st
import pandas as pd
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from num2words import num2words
import json

# ----------------------------------------------------------------------
#  CONFIGURACIÓN Y VARIABLES GLOBALES DE DATOS
# ----------------------------------------------------------------------
RUTA_CREDENCIALES_GCP = "credentials.json"
NOMBRE_GSHEET_MAESTRO = "Base de datos"
NOMBRE_GSHEET_ASIGNACIONES = "Base de asignaciones de multas"

# --------------------------------------------------------------------
#  FUNCIONES DE DATOS
# --------------------------------------------------------------------

# sheets.py

import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import json # <-- Asegúrate de que esta importación esté al principio

# ...

# REEMPLAZA TU FUNCIÓN conectar_gsheet CON ESTA VERSIÓN
@st.cache_resource(show_spinner="Conectando a Google Sheets...")
def conectar_gsheet():
    """Establece conexión con la API de Google Sheets usando los secretos de Streamlit."""
    try:
        # --- INICIO DE LA CORRECCIÓN ---
        # 1. Lee el secreto de Streamlit (que es un bloque de texto/string)
        creds_str = st.secrets["gcp_service_account"]
        
        # 2. Convierte ese texto en un diccionario de Python
        creds_dict = json.loads(creds_str)
        # --- FIN DE LA CORRECCIÓN ---

        scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        
        # 3. Ahora, la función recibe el diccionario que esperaba
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        client = gspread.authorize(creds)
        return client
        
    except json.JSONDecodeError:
        st.error("Error de formato en el secreto 'gcp_service_account'. Asegúrate de haber copiado y pegado el contenido completo de tu archivo credentials.json.")
        return None
    except Exception as e:
        st.error(f"Error al conectar con Google Sheets usando los secretos: {e}")
        return None


@st.cache_data(show_spinner="Cargando datos de la hoja...")
def cargar_hoja_a_df(_client, nombre_archivo, nombre_hoja):
    """Carga una hoja de Google Sheets en un DataFrame de Pandas."""
    if not _client:
        return None
    try:
        sheet = _client.open(nombre_archivo).worksheet(nombre_hoja)
        datos = sheet.get_all_records()
        df = pd.DataFrame(datos).replace(['', 'None'], None)
        return df
    except gspread.exceptions.WorksheetNotFound:
        return None
    except Exception as e:
        st.error(f"No se pudo cargar la hoja '{nombre_hoja}'. Error: {e}")
        return None

def convertir_porcentaje(valor_str):
    """Convierte un string de porcentaje (ej. '75%') a un float (0.75)."""
    if isinstance(valor_str, str) and '%' in valor_str:
        return float(valor_str.replace('%', '').strip()) / 100
    return float(valor_str or 0)

def descargar_archivo_drive(file_id, credentials_path):
    """Descarga un archivo de Google Drive y lo devuelve como un buffer en memoria."""
    try:
        drive_creds = Credentials.from_service_account_file(
            credentials_path,
            scopes=["https://www.googleapis.com/auth/drive.readonly"]
        )
        drive_service = build('drive', 'v3', credentials=drive_creds)
        request = drive_service.files().get_media(fileId=file_id)
        file_buffer = io.BytesIO()
        downloader = MediaIoBaseDownload(file_buffer, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
        file_buffer.seek(0)
        return file_buffer
    except Exception as e:
        st.error(f"Error al descargar el archivo '{file_id}' de Drive: {e}")
        return None

# --------------------------------------------------------------------
#  FUNCIONES DE CÁLCULO
# --------------------------------------------------------------------

def calcular_costo_evitado(datos_entrada):
    """
    Realiza el cálculo del Costo Evitado.
    Versión final, corregida y con todos los placeholders.
    [cite_start]"""
    try:
        # --- 1. Desempaquetar datos ---
        df_items_infracciones = datos_entrada['df_items_infracciones']
        df_costos_items = datos_entrada['df_costos_items']
        df_coti_general = datos_entrada['df_coti_general']
        df_salarios_general = datos_entrada['df_salarios_general']
        df_indices = datos_entrada['df_indices']
        id_infraccion = datos_entrada['id_infraccion']
        fecha_incumplimiento = datos_entrada['fecha_incumplimiento']
        id_rubro = datos_entrada['id_rubro']
        dias_habiles = datos_entrada.get('dias_habiles', 0)
        num_personal_capacitacion = datos_entrada.get('num_personal_capacitacion', 0)

        # --- 2. Preparación de DataFrames y datos base ---
        df_indices['Indice_Mes'] = pd.to_datetime(df_indices['Indice_Mes'], dayfirst=True, errors='coerce')
        df_indices.dropna(subset=['Indice_Mes'], inplace=True)
        df_coti_general['Fecha_Costeo'] = pd.to_datetime(df_coti_general['Fecha_Costeo'], dayfirst=True, errors='coerce')
        df_coti_general.dropna(subset=['Fecha_Costeo'], inplace=True)
        
        fecha_incumplimiento_dt = pd.to_datetime(fecha_incumplimiento)
        ipc_incumplimiento_row = df_indices[df_indices['Indice_Mes'].dt.to_period('M') == fecha_incumplimiento_dt.to_period('M')]
        ipc_incumplimiento = ipc_incumplimiento_row.iloc[0]['IPC_Mensual']
        tipo_cambio_incumplimiento = ipc_incumplimiento_row.iloc[0]['TC_Mensual']

        # --- 3. Inicialización de variables para placeholders ---
        sustentos_de_cotizaciones = []
        fuente_salario_final, pdf_salario_final, incluye_igv_final = '', '', ''
        salario_info_capturado, igv_info_capturado = False, False
        lineas_resumen_fuentes = []
        sustentos, ids_anexos, tabla_final_ce = [], [], []
        
        # --- 4. Bucle principal de cálculo ---
        receta_df = df_items_infracciones[df_items_infracciones['ID_Infraccion'] == id_infraccion]
        for _, item_receta in receta_df.iterrows():
            
            # --- Lógica de selección de ítem ---
            posibles_costos = pd.DataFrame()
            if id_infraccion == 'INF003':
                if num_personal_capacitacion == 1: codigo_item = 'ITEM0110'
                elif 2 <= num_personal_capacitacion <= 5: codigo_item = 'ITEM0111'
                elif 6 <= num_personal_capacitacion <= 10: codigo_item = 'ITEM0112'
                else: codigo_item = 'ITEM0113'
                posibles_costos = df_costos_items[df_costos_items['ID_Item'] == codigo_item].copy()
            else:
                id_item_infraccion_actual = item_receta['ID_Item_Infraccion']
                posibles_costos = df_costos_items[df_costos_items['ID_Item_Infraccion'] == id_item_infraccion_actual].copy()
            
            if posibles_costos.empty: continue

            tipo_item_receta = item_receta.get('Tipo_Item')
            df_candidatos = pd.DataFrame()
            if tipo_item_receta == 'Variable':
                df_candidatos = posibles_costos[posibles_costos['ID_Rubro'] == id_rubro].copy()
            elif tipo_item_receta == 'Fijo':
                df_candidatos = posibles_costos.copy()
            
            if df_candidatos.empty: continue

            fechas_fuente = []
            for _, candidato in df_candidatos.iterrows():
                id_general = candidato['ID_General']
                fecha_fuente = pd.NaT
                if pd.notna(id_general):
                    if 'SAL' in id_general:
                        fuente_salario = df_salarios_general[df_salarios_general['ID_Salario'] == id_general]
                        if not fuente_salario.empty:
                            anio = fuente_salario.iloc[0]['Costeo_Salario']
                            fecha_fuente = pd.to_datetime(f'{int(anio)}-06-30')
                    elif 'COT' in id_general:
                        fuente_cotizacion = df_coti_general[df_coti_general['ID_Cotizacion'] == id_general]
                        if not fuente_cotizacion.empty:
                            fecha_fuente = fuente_cotizacion.iloc[0]['Fecha_Costeo']
                fechas_fuente.append(fecha_fuente)
            
            df_candidatos['Fecha_Fuente'] = fechas_fuente
            df_candidatos.dropna(subset=['Fecha_Fuente'], inplace=True)
            if df_candidatos.empty: continue

            df_candidatos['Diferencia_Dias'] = (df_candidatos['Fecha_Fuente'] - fecha_incumplimiento_dt).dt.days.abs()
            fila_costo_final = df_candidatos.loc[df_candidatos['Diferencia_Dias'].idxmin()]

            if fila_costo_final is not None:
                # --- INICIO DE LA CORRECCIÓN ---
                # Recolectar sustentos y anexos para este ítem
                if pd.notna(fila_costo_final.get('Sustento_Item')):
                    sustentos.append(fila_costo_final['Sustento_Item'])
                if pd.notna(fila_costo_final.get('ID_Anexo_Drive')):
                    ids_anexos.append(fila_costo_final['ID_Anexo_Drive'])
                # --- FIN DE LA CORRECCIÓN ---
                # --- CAPTURA DE DATOS PARA PLACEHOLDERS ---
                id_general = fila_costo_final.get('ID_General')
                if id_general:
                    if 'COT' in id_general:
                        fuente_row = df_coti_general[df_coti_general['ID_Cotizacion'] == id_general]
                        if not fuente_row.empty:
                            sustento = fuente_row.iloc[0].get('Sustento_Cotizacion')
                            if sustento: sustentos_de_cotizaciones.append(sustento)
                            if not igv_info_capturado:
                                estado_igv = fuente_row.iloc[0].get('Incluye_IGV', '')
                                if estado_igv == 'SI': incluye_igv_final = "Precio incluye IGV"
                                elif estado_igv == 'NO': incluye_igv_final = "Precio no incluye IGV"
                                igv_info_capturado = True
                    elif 'SAL' in id_general and not salario_info_capturado:
                        fuente_row = df_salarios_general[df_salarios_general['ID_Salario'] == id_general]
                        if not fuente_row.empty:
                            fuente_salario_final = fuente_row.iloc[0].get('Fuente_Salario', '')
                            pdf_salario_final = fuente_row.iloc[0].get('PDF_Salario', '')
                            salario_info_capturado = True

                # --- CONSTRUCCIÓN DEL RESUMEN DE FUENTES ---
                descripcion_item = fila_costo_final.get('Descripcion_Item', 'Ítem no especificado')
                fecha_fuente_dt = fila_costo_final.get('Fecha_Fuente')
                ipc_costeo_row = df_indices[df_indices['Indice_Mes'].dt.to_period('M') == fecha_fuente_dt.to_period('M')]
                ipc_costeo_valor = ipc_costeo_row.iloc[0]['IPC_Mensual'] if not ipc_costeo_row.empty else 0
                fc_texto = f"promedio ({fecha_fuente_dt.year})" if (id_general and 'SAL' in id_general) else fecha_fuente_dt.strftime('%B %Y').lower()
                ipc_fc_texto = f"{ipc_costeo_valor:,.3f}"
                linea_resumen = f"{descripcion_item}: {fc_texto}, IPC={ipc_fc_texto}"
                lineas_resumen_fuentes.append(linea_resumen)

                # --- CÁLCULO DE COSTOS DEL ÍTEM (CORREGIDO) ---
                costo_original = float(fila_costo_final['Costo_Unitario_Item'])
                moneda_original = fila_costo_final['Moneda_Item']
                
                tc_en_fecha_costeo = 0
                tc_costeo_row = df_indices[df_indices['Indice_Mes'].dt.to_period('M') == fecha_fuente_dt.to_period('M')]
                if not tc_costeo_row.empty:
                    tc_en_fecha_costeo = tc_costeo_row.iloc[0]['TC_Mensual']
                
                precio_base_soles, precio_base_dolares = 0, 0
                if moneda_original == 'US$':
                    precio_base_dolares = costo_original
                    if tc_en_fecha_costeo > 0:
                        precio_base_soles = costo_original * tc_en_fecha_costeo
                else:
                    precio_base_soles = costo_original
                    if tc_en_fecha_costeo > 0:
                        precio_base_dolares = costo_original / tc_en_fecha_costeo
                
                ipc_costeo = ipc_costeo_valor
                if costo_original > 0 and ipc_costeo > 0:
                    precio_base_soles_con_igv = precio_base_soles * 1.18 if fila_costo_final['Incluye_IGV'] == 'NO' else precio_base_soles
                    factor_ajuste = round(ipc_incumplimiento / ipc_costeo, 3)
                    cantidad = float(item_receta.get('Cantidad_Recursos') or 1)
                    horas = (dias_habiles or 0) * 8 if id_infraccion == 'INF004' else float(item_receta.get('Cantidad_Horas') or 1)
                    monto_soles = cantidad * horas * precio_base_soles_con_igv * factor_ajuste
                    monto_dolares = monto_soles / tipo_cambio_incumplimiento
                    
                    tabla_final_ce.append({
                        "descripcion": descripcion_item,
                        "precio_soles": precio_base_soles,
                        "precio_dolares": precio_base_dolares,
                        "factor_ajuste": factor_ajuste,
                        "monto_soles": monto_soles,
                        "monto_dolares": monto_dolares,
                        "costo_original": costo_original,
                        "moneda_original": moneda_original,
                        "cantidad": cantidad,
                        "horas": horas,          
                    })
        
        # --- 5. Ensamblaje y retorno de resultados ---
        if not tabla_final_ce:
            return {'error': "No se pudieron encontrar costos aplicables para los ítems de la infracción con los criterios dados."}

        df_presentacion = pd.DataFrame(tabla_final_ce)
        fuente_coti_texto = "\n".join([f"- {s}" for s in list(dict.fromkeys(sustentos_de_cotizaciones))])
        resumen_final_texto = "\n".join(lineas_resumen_fuentes)
        fi_mes_texto = fecha_incumplimiento_dt.strftime('%B %Y').lower()

        return {
            "ce_data_raw": df_presentacion.to_dict('records'),
            "total_soles": df_presentacion['monto_soles'].sum(),
            "total_dolares": df_presentacion['monto_dolares'].sum(),
            "sustentos": list(set(sustentos)),
            "ids_anexos": list(set(ids_anexos)),
            "error": None,
            "fuente_coti": fuente_coti_texto,
            "fuente_salario": fuente_salario_final,
            "pdf_salario": pdf_salario_final,
            "fi_mes": fi_mes_texto,
            "fi_ipc": f"{ipc_incumplimiento:,.3f}",
            "fi_tc": f"{tipo_cambio_incumplimiento:,.3f}",
            "resumen_fuentes_costo": resumen_final_texto,
            "incluye_igv": incluye_igv_final
        }
    except Exception as e:
        import traceback
        print(traceback.format_exc())
        return {'error': f"Error crítico en el cálculo del CE: {e}"}

def calcular_beneficio_ilicito(datos_entrada):
    """
    Realiza el cálculo completo del Beneficio Ilícito (BI).
    Recibe un diccionario con los DataFrames y datos necesarios.
    Devuelve un diccionario con los resultados.
    """
    try:
        df_cos = datos_entrada['df_cos']
        df_uit = datos_entrada['df_uit']
        df_indices_bi = datos_entrada['df_indices_bi']
        rubro = datos_entrada['rubro']
        ce_soles = datos_entrada['ce_soles']
        ce_dolares = datos_entrada['ce_dolares']
        fecha_incumplimiento_calc = datos_entrada['fecha_incumplimiento']
        cos_info = df_cos[df_cos['Sector_Rubro'] == rubro]
        if cos_info.empty:
            return {'error': f"No se encontró información de COS para el rubro '{rubro}'."}
        fuente_cos = cos_info.iloc[0]['Fuente_COS']
        moneda_cos = str(cos_info.iloc[0]['Moneda_COS']).strip()
        cos_anual = convertir_porcentaje(cos_info.iloc[0]['COS_Anual'])
        cos_mensual = convertir_porcentaje(cos_info.iloc[0]['COS_Mensual'])
        ce = ce_soles if moneda_cos == 'S/' else ce_dolares
        fecha_calculo = date.today()
        diff = relativedelta(fecha_calculo, fecha_incumplimiento_calc)
        meses_completos = diff.years * 12 + diff.months
        dias_sobrantes = diff.days
        t_meses_decimal = meses_completos + (dias_sobrantes / 30.0)
        ce_ajustado = ce * ((1 + cos_mensual) ** t_meses_decimal)
        df_indices_bi['Indice_Mes'] = pd.to_datetime(df_indices_bi['Indice_Mes'], dayfirst=True, errors='coerce')
        df_indices_bi.dropna(subset=['Indice_Mes'], inplace=True)
        df_indices_bi['TC_Mensual'] = pd.to_numeric(df_indices_bi['TC_Mensual'], errors='coerce')
        end_date_tc = pd.to_datetime(fecha_calculo)
        start_date_tc = end_date_tc - relativedelta(months=12)
        tc_promedio_df = df_indices_bi[(df_indices_bi['Indice_Mes'] > start_date_tc) & (df_indices_bi['Indice_Mes'] <= end_date_tc)]
        tc_promedio_12m = tc_promedio_df['TC_Mensual'].mean()
        beneficio_ilicito_soles = ce_ajustado if moneda_cos == 'S/' else ce_ajustado * tc_promedio_12m
        ano_actual = fecha_calculo.year
        uit_info = df_uit[df_uit['Año_UIT'] == ano_actual]
        valor_uit = float(uit_info.iloc[0]['Valor_UIT']) if not uit_info.empty else 0
        beneficio_ilicito_uit = beneficio_ilicito_soles / valor_uit if valor_uit > 0 else 0
        tabla_resumen_data = {
            "Descripción": ["CE para el hecho imputado", "COS (anual)", "COSm (mensual)", "T: meses transcurridos", "Costo evitado ajustado", "Tipo de cambio promedio", "Beneficio ilícito (S/)", f"UIT al año {ano_actual}", "Beneficio Ilícito (UIT)"],
            "Monto": [f"{'S/' if moneda_cos == 'S/' else 'US$'} {ce:,.3f}", f"{cos_anual:,.3%}", f"{cos_mensual:,.3%}", f"{t_meses_decimal:,.3f}", f"{'S/' if moneda_cos == 'S/' else 'US$'} {ce_ajustado:,.3f}", f"{tc_promedio_12m:,.3f}", f"S/ {beneficio_ilicito_soles:,.3f}", f"S/ {valor_uit:,.2f}", f"{beneficio_ilicito_uit:,.3f}"]
        }
        df_resumen = pd.DataFrame(tabla_resumen_data)
        return {
            "beneficio_ilicito_uit": beneficio_ilicito_uit,
            "bi_data_raw": df_resumen.to_dict('records'),
            "fuente_cos": fuente_cos,
            "error": None
        }
    except Exception as e:
        return {'error': f"Error en el cálculo del BI: {e}"}

def calcular_multa(datos_entrada):
    """
    Realiza el cálculo final de la multa para un hecho imputado.
    Recibe un diccionario con el DataFrame de tipificación y los datos necesarios.
    Devuelve un diccionario con los resultados.
    """
    try:
        df_tipificacion = datos_entrada['df_tipificacion']
        id_infraccion = datos_entrada['id_infraccion']
        b = datos_entrada['beneficio_ilicito']
        infraccion_info = df_tipificacion[df_tipificacion['ID_Infraccion'] == id_infraccion]
        p = 0
        if not infraccion_info.empty:
            p_str = infraccion_info.iloc[0]['Prob_Deteccion']
            p = convertir_porcentaje(p_str)
        f = 1.0
        multa_uit = (b / p) * f if p > 0 else 0
        tabla_multa_data = {
            "Componentes": ["Beneficio Ilícito (B)", "Probabilidad de detección (p)", "Factores para la graduación de sanciones F=(1+f1+f2+f3+f4+f5+f6+f7)", "Multa en UIT (B/p)*(F)"],
            "Monto": [f"{b:,.3f} UIT", f"{p:,.3f}", f"{f:.0%}", f"{multa_uit:,.3f} UIT"]
        }
        df_multa = pd.DataFrame(tabla_multa_data)
        return {
            "multa_final_uit": multa_uit,
            "multa_data_raw": df_multa.to_dict('records'),
            "error": None if p > 0 else f"No se encontró la probabilidad de detección para la infracción {id_infraccion}."
        }
    except Exception as e:
        return {'error': f"Error en el cálculo de la multa: {e}"}

# --------------------------------------------------------------------
#  INFORME FINAL
# --------------------------------------------------------------------

def get_person_details_by_base_name(nombre_base, df_analistas_func):
    """Busca y devuelve los detalles de una persona en el DataFrame de analistas."""
    if nombre_base and df_analistas_func is not None:
        info = df_analistas_func[df_analistas_func['Nombre_Base_Analista'] == str(nombre_base)]
        if not info.empty:
            details = {
                'titulo': info.iloc[0].get('Titulo_Analista'),
                'nombre': info.iloc[0].get('Nombre_Analista'),
                'cargo': info.iloc[0].get('Cargo_Analista'),
                'colegiatura': info.iloc[0].get('Colegiatura_Analista')
            }
            return {k: (v if v is not None else '') for k, v in details.items()}
    return {'titulo': '', 'nombre': '', 'cargo': '', 'colegiatura': ''}