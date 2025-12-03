import io
import gspread
import streamlit as st
import pandas as pd
from babel.dates import format_date
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from num2words import num2words
# Cerca de la línea 13 de sheets.py
from docxtpl import DocxTemplate, RichText
from funciones import create_main_table_subdoc, create_table_subdoc, texto_con_numero, create_footnotes_subdoc, format_decimal_dinamico, redondeo_excel

# ----------------------------------------------------------------------
#  CONFIGURACIÓN Y VARIABLES GLOBALES DE DATOS
# ----------------------------------------------------------------------
RUTA_CREDENCIALES_GCP = "credentials.json"
NOMBRE_GSHEET_MAESTRO = "Base de datos"
NOMBRE_GSHEET_ASIGNACIONES = "Base de asignaciones de multas"

# --------------------------------------------------------------------
#  FUNCIONES DE DATOS
# --------------------------------------------------------------------

import streamlit as st

def actualizar_hoja_con_df(cliente, nombre_archivo, nombre_hoja, df_nuevos_datos):
    """
    Abre una hoja de Google Sheets y añade las filas de un DataFrame al final.
    """
    try:
        sheet = cliente.open(nombre_archivo).worksheet(nombre_hoja)
        
        # Convertir DataFrame a lista de listas para gspread
        # Asegurarse de que las fechas tengan el formato correcto de la hoja
        df_copia = df_nuevos_datos.copy()
        df_copia['Indice_Mes'] = df_copia['Indice_Mes'].dt.strftime('%d/%m/%Y')
        
        # Obtener solo las columnas que están en la hoja (en el orden correcto)
        # Asumiendo que las 3 primeras columnas son Indice_Mes, IPC_Mensual, TC_Mensual
        columnas_finales = ['Indice_Mes', 'IPC_Mensual', 'TC_Mensual']
        lista_de_valores = df_copia[columnas_finales].values.tolist()
        
        if not lista_de_valores:
            return 0 # No hay nada que añadir
            
        # Añadir las nuevas filas
        sheet.append_rows(lista_de_valores, value_input_option='USER_ENTERED')
        
        return len(lista_de_valores)
    except Exception as e:
        st.error(f"Error al actualizar la hoja '{nombre_hoja}': {e}")
        import traceback
        traceback.print_exc()
        return -1


@st.cache_resource(show_spinner="Conectando a Google Sheets...")
def conectar_gsheet():
    """Establece conexión con la API de Google Sheets usando los secretos de Streamlit."""
    try:
        # Streamlit ahora nos entrega un diccionario directamente
        creds_dict = st.secrets["gcp_service_account"]
        
        scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        client = gspread.authorize(creds)
        return client
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

def descargar_archivo_drive(file_id): # Ya no necesita 'credentials_path'
    """
    Descarga un archivo de Google Drive usando los secretos de Streamlit.
    """
    try:
        # Lee las credenciales desde los secretos de Streamlit
        creds_dict = st.secrets["gcp_service_account"]
        
        # Crea las credenciales desde el diccionario, no desde un archivo
        drive_creds = Credentials.from_service_account_info(
            creds_dict,
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

def calcular_beneficio_ilicito_extemporaneo(datos_entrada):
    """
    Calcula el BI para casos tardíos. La tabla de resultados ahora es dinámica
    y se simplifica si la moneda del COS es en Soles.
    """
    try:
        # 1. Obtenemos el texto del hecho que pasamos en el paso anterior
        texto_hecho = datos_entrada.get('texto_del_hecho', 'CE para el hecho imputado')
        df_indices = datos_entrada['df_indices']
        df_uit = datos_entrada['df_uit']
        fecha_incumplimiento_calc = datos_entrada['fecha_incumplimiento']
        fecha_cumplimiento_extemporaneo = datos_entrada['fecha_cumplimiento_extemporaneo']
        ce_soles = datos_entrada['ce_soles']
        ce_dolares = datos_entrada['ce_dolares']
        cos_anual = datos_entrada.get('cos_anual', 0)
        cos_mensual = datos_entrada.get('cos_mensual', 0)
        moneda_cos = datos_entrada.get('moneda_cos', 'S/')
        fuente_cos = datos_entrada.get('fuente_cos', '')
        ce = ce_soles if moneda_cos == 'S/' else ce_dolares
        fecha_hoy = datos_entrada.get('fecha_emision_informe', date.today()) # <-- CAMBIO
        diff_cap = relativedelta(fecha_cumplimiento_extemporaneo, fecha_incumplimiento_calc)
        t_cap = (diff_cap.years * 12 + diff_cap.months) + redondeo_excel(diff_cap.days / 30.0, 3)
        ce_ajustado_cap = redondeo_excel(ce * ((1 + cos_mensual) ** t_cap), 3)
        
        end_date_tc = pd.to_datetime(fecha_cumplimiento_extemporaneo)
        start_date_tc = end_date_tc - relativedelta(months=12)
        tc_promedio_df = df_indices[(df_indices['Indice_Mes'] > start_date_tc) & (df_indices['Indice_Mes'] <= end_date_tc)]
        tc_promedio_12m = redondeo_excel(tc_promedio_df['TC_Mensual'].mean() if not tc_promedio_df.empty else 0, 3)        
        bi_cap_soles = ce_ajustado_cap if moneda_cos == 'S/' else ce_ajustado_cap * tc_promedio_12m
        df_indices_sorted = df_indices.dropna(subset=['Indice_Mes']).sort_values(by='Indice_Mes', ascending=False)
        ipc_hoy = float(df_indices_sorted.iloc[0]['IPC_Mensual']) if not df_indices_sorted.empty and df_indices_sorted.iloc[0]['IPC_Mensual'] is not None else 0
        ipc_ext_row = df_indices[df_indices['Indice_Mes'].dt.to_period('M') == pd.to_datetime(fecha_cumplimiento_extemporaneo).to_period('M')]
        ipc_ext = float(ipc_ext_row.iloc[0]['IPC_Mensual']) if not ipc_ext_row.empty and ipc_ext_row.iloc[0]['IPC_Mensual'] is not None else 0
        ajuste_inflacionario = redondeo_excel(ipc_hoy / ipc_ext if ipc_ext > 0 else 1, 3)
        bi_final_soles = bi_cap_soles * ajuste_inflacionario
        valor_uit_row = df_uit[df_uit['Año_UIT'] == fecha_hoy.year]
        valor_uit = float(valor_uit_row.iloc[0]['Valor_UIT']) if not valor_uit_row.empty else 0
        beneficio_ilicito_uit = redondeo_excel(bi_final_soles / valor_uit if valor_uit > 0 else 0, 3)
# --- INICIO DE LA MODIFICACIÓN ---
        # 2. Crea el mapa que conecta cada letra con el texto de la fuente.
        #    (SE MUEVE ANTES DE USARLO)
        footnote_mapping = {
            'a': 'ce_anexo',
            'b': 'cok',
            'c': 'periodo_bi_ext',
            'd': 'ce_ext',
            'e': 'bcrp',
            'f': 'ajuste_inflacionario_detalle',
            'g': 'ipc_fecha',
            'h': 'sunat'
        }

# 3. Construcción dinámica de la tabla de resultados
        tabla_resumen_filas = [
            {"descripcion_texto": f"CE: {texto_hecho}", "descripcion_superindice": "", "monto": f"{'S/' if moneda_cos == 'S/' else 'US$'} {ce:,.3f}", "ref": "a"},
            {"descripcion_texto": "COS (anual)", "descripcion_superindice": "", "monto": f"{cos_anual:,.3%}", "ref": "b"},
            {"descripcion_texto": "COSm (mensual)", "descripcion_superindice": "", "monto": f"{cos_mensual:,.3%}", "ref": None},
            {"descripcion_texto": f"T: Meses transcurridos desde el periodo de incumplimiento hasta la fecha de cumplimiento extemporáneo", "descripcion_superindice": "", "monto": f"{t_cap:,.3f}", "ref": "c"},
        ]

        # Lógica condicional para las filas que cambian
        if moneda_cos == 'S/':
            # Si el COS es en Soles, se muestra la fila simplificada
            tabla_resumen_filas.append(
                {"descripcion_texto": "Costo evitado ajustado a la fecha de cumplimiento extemporáneo de la conducta: CE*(1+COSm)", "descripcion_superindice": "T", "monto": f"S/ {bi_cap_soles:,.3f}", "ref": "d"}
            )
        else:
            # Si el COS es en Dólares, se muestran los pasos de conversión
            tabla_resumen_filas.extend([
                {"descripcion_texto": "Costo evitado ajustado a la fecha de cumplimiento extemporáneo de la conducta: CE*(1+COSm)", "descripcion_superindice": "T", "monto": f"US$ {ce_ajustado_cap:,.3f}", "ref": "d"},
                {"descripcion_texto": "Tipo de cambio promedio de los últimos 12 meses a fecha de cumplimiento extemporáneo", "descripcion_superindice": "", "monto": f"{tc_promedio_12m:,.3f}", "ref": "e"},
                {"descripcion_texto": f"Beneficio ilícito a la fecha de cumplimiento extemporáneo (S/)", "descripcion_superindice": "", "monto": f"S/ {bi_cap_soles:,.3f}", "ref": None},
            ])

        # Se añaden las filas finales que son comunes a ambos casos
        tabla_resumen_filas.extend([
            {"descripcion_texto": "Ajuste inflacionario desde la fecha de cumplimiento extemporáneo hasta la fecha de emisión del presente informe", "descripcion_superindice": "", "monto": f"{ajuste_inflacionario:,.3f}", "ref": "f"},
            {"descripcion_texto": "Beneficio ilícito a la fecha de emisión del informe (S/)", "descripcion_superindice": "", "monto": f"S/ {bi_final_soles:,.3f}", "ref": "g"},
            {"descripcion_texto": f"Unidad Impositiva Tributaria al año {fecha_hoy.year} - UIT {fecha_hoy.year}", "descripcion_superindice": "", "monto": f"S/ {valor_uit:,.2f}", "ref": "h"},
            {"descripcion_texto": "Beneficio Ilícito (UIT)", "descripcion_superindice": "", "monto": f"{beneficio_ilicito_uit:,.3f} UIT", "ref": None}
        ])

        # 4. Recolecta los datos necesarios para formatear las plantillas.
        datos_para_fuentes = {
            'rubro': datos_entrada.get('rubro', ''),
            'fuente_cos': fuente_cos,
            'fecha_incumplimiento_texto': format_date(fecha_incumplimiento_calc, "d 'de' MMMM 'de' yyyy", locale='es'),
            'fecha_extemporanea_texto': format_date(fecha_cumplimiento_extemporaneo, "d 'de' MMMM 'de' yyyy", locale='es'),
            'mes_actual_texto': format_date(fecha_hoy, "MMMM 'de' yyyy", locale='es'),
            'ultima_fecha_ipc_texto': format_date(df_indices.dropna(subset=['Indice_Mes']).sort_values(by='Indice_Mes', ascending=False).iloc[0]['Indice_Mes'], 'MMMM yyyy', locale='es'),
            'fecha_hoy_texto': format_date(fecha_hoy, "d 'de' MMMM 'de' yyyy", locale='es'),
            'mes_ipc_hoy_texto': format_date(df_indices_sorted.iloc[0]['Indice_Mes'], "MMMM 'de' yyyy", locale='es'),
            'mes_ipc_ext_texto': format_date(fecha_cumplimiento_extemporaneo, "MMMM 'de' yyyy", locale='es'),
            'valor_ipc_hoy': ipc_hoy,
            'valor_ipc_ext': ipc_ext
        }

        # 5. Devolver la estructura de datos completa y estandarizada.
        return {
            "table_rows": tabla_resumen_filas,
            "footnote_mapping": footnote_mapping,
            "footnote_data": datos_para_fuentes,
            "beneficio_ilicito_uit": beneficio_ilicito_uit,
            "error": None
        }

    except Exception as e:
        import traceback
        traceback.print_exc()
        return {'error': f"Error en cálculo de BI extemporáneo: {e}"}

def calcular_beneficio_ilicito(datos_entrada):
    """
    Realiza el cálculo del BI. La tabla de resultados ahora es dinámica
    y se simplifica si la moneda del COS es en Soles.
    """
    try:

        # 1. Obtenemos el texto del hecho que pasamos en el paso anterior
        texto_hecho = datos_entrada.get('texto_del_hecho', 'CE para el hecho imputado')

        df_cos = datos_entrada['df_cos']
        df_uit = datos_entrada['df_uit']
        df_indices = datos_entrada['df_indices']
        rubro = datos_entrada['rubro']
        ce_soles = datos_entrada['ce_soles']
        ce_dolares = datos_entrada['ce_dolares']
        fecha_incumplimiento_calc = datos_entrada['fecha_incumplimiento']
        fecha_calculo = datos_entrada.get('fecha_emision_informe', date.today()) # <-- CAMBIO

        cos_info = df_cos[df_cos['Sector_Rubro'] == rubro]
        if cos_info.empty:
            return {'error': f"No se encontró información de COS para el rubro '{rubro}'."}
        
        fuente_cos = cos_info.iloc[0]['Fuente_COS']
        moneda_cos = str(cos_info.iloc[0]['Moneda_COS']).strip()
        cos_anual = convertir_porcentaje(cos_info.iloc[0]['COS_Anual'])
        cos_mensual = convertir_porcentaje(cos_info.iloc[0]['COS_Mensual'])
        ce = ce_soles if moneda_cos == 'S/' else ce_dolares
        
        diff = relativedelta(fecha_calculo, fecha_incumplimiento_calc)
        t_meses_decimal = (diff.years * 12 + diff.months) + redondeo_excel(diff.days / 30.0, 3)
        
        ce_ajustado = redondeo_excel(ce * ((1 + cos_mensual) ** t_meses_decimal), 3)

        # --- INICIO: CORRECCIÓN TC PROMEDIO 12 MESES ---
        
        # 1. Definir la fecha de emisión del informe
        fecha_emision_dt = pd.to_datetime(fecha_calculo)

        # 2. Encontrar el último mes CON DATOS de TC disponible en o antes de la fecha de emisión
        df_indices_disponibles = df_indices[
            (df_indices['Indice_Mes'] <= fecha_emision_dt)
        ].dropna(subset=['TC_Mensual']).sort_values(by='Indice_Mes', ascending=False)

        if df_indices_disponibles.empty:
            return {'error': f"No se encontraron datos de TC en 'Indices_BCRP' en o antes de {fecha_emision_dt.strftime('%Y-%m-%d')}"}

        # 3. Esta es la *verdadera* fecha final para el cálculo (ej. 2025-09-01)
        end_date_tc = df_indices_disponibles.iloc[0]['Indice_Mes']
        
        # 4. Calcular la fecha de inicio (12 meses atrás desde la última fecha CON DATOS)
        start_date_tc = end_date_tc - relativedelta(months=12) # ej. 2024-09-01
        
        # 5. Filtrar el DataFrame (El filtro > start_date e <= end_date asegura los 12 meses)
        tc_promedio_df = df_indices[
            (df_indices['Indice_Mes'] > start_date_tc) & 
            (df_indices['Indice_Mes'] <= end_date_tc)
        ].dropna(subset=['TC_Mensual'])
        
        tc_promedio_12m = redondeo_excel(tc_promedio_df['TC_Mensual'].mean() if not tc_promedio_df.empty else 0, 3)

        beneficio_ilicito_soles = ce_ajustado if moneda_cos == 'S/' else ce_ajustado * tc_promedio_12m
        # --- FIN: CORRECCIÓN ---
        
        uit_info = df_uit[df_uit['Año_UIT'] == fecha_calculo.year]
        valor_uit = float(uit_info.iloc[0]['Valor_UIT']) if not uit_info.empty else 0
        beneficio_ilicito_uit = redondeo_excel(beneficio_ilicito_soles / valor_uit if valor_uit > 0 else 0, 3)
        # 2. Crea un mapa que le dice al sistema qué texto corresponde a cada letra
        footnote_mapping = {
            'a': 'ce_anexo',
            'b': 'cok',
            'c': 'periodo_bi',
            'd': 'bcrp',
            'e': 'ipc_fecha',  
            'f': 'sunat'       
        }
        
# 3. Construcción dinámica de la tabla de resultados
        tabla_resumen_filas = [
            {"descripcion_texto": f"CE: {texto_hecho}", "descripcion_superindice": "", "monto": f"{'S/' if moneda_cos == 'S/' else 'US$'} {ce:,.3f}", "ref": "a"},
            {"descripcion_texto": "COS (anual)", "descripcion_superindice": "", "monto": f"{cos_anual:,.3%}", "ref": "b"},
            {"descripcion_texto": "COSm (mensual)", "descripcion_superindice": "", "monto": f"{cos_mensual:,.3%}", "ref": None},
            {"descripcion_texto": "T: meses transcurridos durante el periodo de incumplimiento", "descripcion_superindice": "", "monto": f"{t_meses_decimal:,.3f}", "ref": "c"},
        ]

        # Lógica condicional para las filas que cambian
        if moneda_cos == 'S/':
            tabla_resumen_filas.append(
                {"descripcion_texto": "Costo evitado ajustado a la fecha del cálculo de la multa: CE*(1+COSm)", "descripcion_superindice": "T", "monto": f"S/ {beneficio_ilicito_soles:,.3f}", "ref": None}
            )
        else:
            tabla_resumen_filas.extend([
                {"descripcion_texto": "Costo evitado ajustado a la fecha del cálculo de la multa: CE*(1+COSm)", "descripcion_superindice": "T", "monto": f"US$ {ce_ajustado:,.3f}", "ref": None},
                {"descripcion_texto": "Tipo de cambio promedio de los últimos 12 meses", "descripcion_superindice": "", "monto": f"{tc_promedio_12m:,.3f}", "ref": "d"},
                {"descripcion_texto": "Beneficio ilícito a la fecha del cálculo de la multa (S/)", "descripcion_superindice": "", "monto": f"S/ {beneficio_ilicito_soles:,.3f}", "ref": "e"},
            ])
        
        # Se añaden las filas finales que son comunes a ambos casos
        tabla_resumen_filas.extend([
            {"descripcion_texto": f"Unidad Impositiva Tributaria al año {fecha_calculo.year} - UIT {fecha_calculo.year}", "descripcion_superindice": "", "monto": f"S/ {valor_uit:,.2f}", "ref": "f"},
            {"descripcion_texto": "Beneficio Ilícito (UIT)", "descripcion_superindice": "", "monto": f"{beneficio_ilicito_uit:,.3f} UIT", "ref": None}
        ])

        # 3. Recolectar datos para formatear las fuentes
        datos_para_fuentes = {
            'rubro': rubro,
            'fuente_cos': fuente_cos,
            # --- LÍNEAS MODIFICADAS ---
            'fecha_incumplimiento_texto': format_date(fecha_incumplimiento_calc, "d 'de' MMMM 'de' yyyy", locale='es'),
            'fecha_hoy_texto': format_date(fecha_calculo, "d 'de' MMMM 'de' yyyy", locale='es'),
            'mes_actual_texto': format_date(fecha_calculo, "MMMM 'de' yyyy", locale='es'),
            'ultima_fecha_ipc_texto': format_date(df_indices.dropna(subset=['Indice_Mes']).sort_values(by='Indice_Mes', ascending=False).iloc[0]['Indice_Mes'], 'MMMM yyyy', locale='es'),
        }

        # 4. Devolver la nueva estructura de datos
        return {
            "table_rows": tabla_resumen_filas,
            "footnote_mapping": footnote_mapping,
            "footnote_data": datos_para_fuentes,
            "beneficio_ilicito_uit": beneficio_ilicito_uit,
            "fuente_cos": fuente_cos, # Mantenemos estos para compatibilidad si es necesario
            "cos_anual": cos_anual,
            "cos_mensual": cos_mensual,
            "moneda_cos": moneda_cos,
            "error": None
        }

    except Exception as e:
        import traceback
        traceback.print_exc()
        return {'error': f"Error en el cálculo del BI: {e}"}

def calcular_multa(datos_entrada):
    """
    Realiza el cálculo final de la multa para un hecho imputado.
    Recibe un diccionario con el DataFrame de tipificación y los datos necesarios.
    Ahora acepta 'factor_f' para la graduación.
    """
    try:
        df_tipificacion = datos_entrada['df_tipificacion']
        id_infraccion = datos_entrada['id_infraccion']
        b = datos_entrada['beneficio_ilicito']
        
        # --- INICIO CAMBIO: Recibir Factor F ---
        # Si no se pasa (ej. scripts antiguos), se asume 1.0 (100%)
        f_factor = datos_entrada.get('factor_f', 1.0)
        # --- FIN CAMBIO ---

        infraccion_info = df_tipificacion[df_tipificacion['ID_Infraccion'] == id_infraccion]
        p = 0
        if not infraccion_info.empty:
            p_str = infraccion_info.iloc[0]['Prob_Deteccion']
            p = convertir_porcentaje(p_str)
        
        # --- INICIO CAMBIO: Aplicar F en la fórmula ---
        # Multa = (Beneficio / Probabilidad) * Factor_Graduación
        multa_uit = redondeo_excel((b / p) * f_factor, 3) if p > 0 else 0
        # --- FIN CAMBIO ---

        tabla_multa_data = {
            "Componentes": [
                "Beneficio Ilícito (B)", 
                "Probabilidad de detección (p)", 
                "Factores para la graduación de sanciones F=(1+f1+f2+f3+f4+f5+f6+f7)", 
                "Multa en UIT (B/p)*(F)"
            ],
            "Monto": [
                f"{b:,.3f} UIT", 
                f"{p:,.3f}", 
                f"{f_factor:.2%}", # Mostrar el porcentaje real (ej. 110.00%)
                f"{multa_uit:,.3f} UIT"
            ]
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


# --- INICIO: NUEVAS FUNCIONES DE MEMORIA DE CASOS ---

import json
from datetime import datetime, date

def json_serializador_fecha(obj):
    """Convierte objetos de fecha/datetime a string ISO para JSON."""
    if isinstance(obj, (date, datetime)):
        return obj.isoformat()
    raise TypeError(f"Tipo {type(obj)} no es serializable en JSON")

def guardar_datos_caso(cliente, expediente, producto, datos_python):
    """
    Guarda el diccionario de datos de un caso en GSheet 'Memoria_Casos' como JSON.
    """
    try:
        sheet_memoria = cliente.open("Base de datos").worksheet("Memoria_Casos")
        
        # 1. Convertir el diccionario de Python a un string JSON
        # Usamos 'default' para manejar las fechas (date, datetime)
        datos_en_json = json.dumps(datos_python, default=json_serializador_fecha)
        
        # 2. Preparar la fila
        fecha_actual_str = datetime.now().isoformat()
        
        # 3. Buscar si el expediente ya existe
        try:
            cell = sheet_memoria.find(expediente)
        except gspread.exceptions.CellNotFound:
            cell = None
            
        if cell:
            # YA EXISTE: Actualizar la fila existente
            fila_para_actualizar = [expediente, producto, fecha_actual_str, datos_en_json]
            sheet_memoria.update(f"A{cell.row}:D{cell.row}", [fila_para_actualizar])
            return True, f"Avance actualizado para {expediente}."
        else:
            # NO EXISTE: Añadir una fila nueva
            fila_nueva = [expediente, producto, fecha_actual_str, datos_en_json]
            sheet_memoria.append_row(fila_nueva)
            return True, f"Avance guardado por primera vez para {expediente}."

    except Exception as e:
        import traceback
        traceback.print_exc()
        return False, f"Error al guardar en GSheet: {e}"

def cargar_datos_caso(cliente, expediente):
    """
    Carga el string JSON desde GSheet 'Memoria_Casos' y lo convierte a diccionario Python.
    """
    try:
        sheet_memoria = cliente.open("Base de datos").worksheet("Memoria_Casos")
        
        # 1. Buscar el expediente
        try:
            cell = sheet_memoria.find(expediente)
        except gspread.exceptions.CellNotFound:
            return None, f"No se encontró un avance guardado para {expediente}."
            
        # 2. Obtener el string JSON de la columna 4 (datos_json)
        json_guardado = sheet_memoria.cell(cell.row, 4).value
        if not json_guardado:
            return None, "La celda de datos estaba vacía."
            
        # 3. Convertir de JSON a diccionario Python
        datos_cargados = json.loads(json_guardado)
        
        # 4. ¡CRÍTICO! Re-convertir los strings de fecha a objetos 'date'
        
        # Primero las fechas de 'imputaciones_data' (las más complejas)
        for hecho in datos_cargados.get('imputaciones_data', []):
            # Fechas de reducción (de app.py)
            if hecho.get('memo_fecha'):
                hecho['memo_fecha'] = datetime.fromisoformat(hecho['memo_fecha']).date()
            if hecho.get('escrito_fecha'):
                hecho['escrito_fecha'] = datetime.fromisoformat(hecho['escrito_fecha']).date()
                
            # Fechas de extremos (de INF007, INF008, etc.)
            for extremo in hecho.get('extremos', []):
                if extremo.get('fecha_incumplimiento'):
                    extremo['fecha_incumplimiento'] = datetime.fromisoformat(extremo['fecha_incumplimiento']).date()
                if extremo.get('fecha_extemporanea'):
                    extremo['fecha_extemporanea'] = datetime.fromisoformat(extremo['fecha_extemporanea']).date()

        # Ahora las fechas del 'Paso 2' y 'Paso 3.5'
        if datos_cargados.get('fecha_emision_informe'):
            datos_cargados['fecha_emision_informe'] = datetime.fromisoformat(datos_cargados['fecha_emision_informe']).date()
        if datos_cargados.get('fecha_rsd'):
            datos_cargados['fecha_rsd'] = datetime.fromisoformat(datos_cargados['fecha_rsd']).date()
        
        for anio_data in datos_cargados.get('confiscatoriedad', {}).get('datos_por_anio', {}).values():
            if anio_data.get('escrito_fecha_conf'):
                anio_data['escrito_fecha_conf'] = datetime.fromisoformat(anio_data['escrito_fecha_conf']).date()

        return datos_cargados, "Datos cargados exitosamente."

    except Exception as e:
        import traceback
        traceback.print_exc()
        return None, f"Error al cargar/procesar datos: {e}"

# --- FIN: NUEVAS FUNCIONES DE MEMORIA DE CASOS ---