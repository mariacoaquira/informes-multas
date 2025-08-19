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
    Calcula el Beneficio Ilícito para casos de cumplimiento tardío y construye
    la tabla de resultados detallada.
    """
    try:
        # 1. Desempaquetar datos
        # ... (esta parte se mantiene igual) ...
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
        fecha_hoy = date.today()

        # 2. Cálculos numéricos
        # ... (toda tu lógica de cálculo para t_cap, bi_cap_soles, ajuste_inflacionario, etc. se mantiene igual) ...
        diff_cap = relativedelta(fecha_cumplimiento_extemporaneo, fecha_incumplimiento_calc)
        t_cap = (diff_cap.years * 12 + diff_cap.months) + (diff_cap.days / 30.0)
        ce_ajustado_cap = ce * ((1 + cos_mensual) ** t_cap)
        end_date_tc = pd.to_datetime(fecha_cumplimiento_extemporaneo)
        start_date_tc = end_date_tc - relativedelta(months=12)
        tc_promedio_df = df_indices[(df_indices['Indice_Mes'] > start_date_tc) & (df_indices['Indice_Mes'] <= end_date_tc)]
        tc_promedio_12m = tc_promedio_df['TC_Mensual'].mean()
        bi_cap_soles = ce_ajustado_cap if moneda_cos == 'S/' else ce_ajustado_cap * tc_promedio_12m
        df_indices_sorted = df_indices.dropna(subset=['Indice_Mes']).sort_values(by='Indice_Mes', ascending=False)
        ipc_hoy = 0
        if not df_indices_sorted.empty:
            valor_ipc_hoy = df_indices_sorted.iloc[0]['IPC_Mensual']
            if valor_ipc_hoy is not None:
                ipc_hoy = float(valor_ipc_hoy)
        ipc_ext_row = df_indices[df_indices['Indice_Mes'].dt.to_period('M') == pd.to_datetime(fecha_cumplimiento_extemporaneo).to_period('M')]
        ipc_ext = 0
        if not ipc_ext_row.empty:
            valor_ipc_ext = ipc_ext_row.iloc[0]['IPC_Mensual']
            if valor_ipc_ext is not None:
                ipc_ext = float(valor_ipc_ext)
        ajuste_inflacionario = ipc_hoy / ipc_ext if ipc_ext > 0 else 1
        bi_final_soles = bi_cap_soles * ajuste_inflacionario
        valor_uit_row = df_uit[df_uit['Año_UIT'] == fecha_hoy.year]
        valor_uit = float(valor_uit_row.iloc[0]['Valor_UIT']) if not valor_uit_row.empty else 0
        beneficio_ilicito_uit = bi_final_soles / valor_uit if valor_uit > 0 else 0

        # ---- INICIO DE LA SECCIÓN MODIFICADA ----
        
        # 5. Definir las filas de la tabla con sus claves de referencia para las fuentes
        tabla_resumen_filas = [
            {"descripcion": "CE para el hecho imputado", "monto": f"{'S/' if moneda_cos == 'S/' else 'US$'} {ce:,.3f}", "ref": "ce_anexo"},
            {"descripcion": "COS (anual)", "monto": f"{cos_anual:,.3%}", "ref": "cok"},
            {"descripcion": "COSm (mensual)", "monto": f"{cos_mensual:,.3%}", "ref": "cok"},
            {"descripcion": f"T: Meses hasta cumplimiento extemporáneo ({fecha_cumplimiento_extemporaneo.strftime('%d/%m/%Y')})", "monto": f"{t_cap:,.3f}", "ref": "periodo_bi_ext"},
            {"descripcion": "Costo evitado ajustado", "monto": f"{'S/' if moneda_cos == 'S/' else 'US$'} {ce_ajustado_cap:,.3f}", "ref": None},
            {"descripcion": "Tipo de cambio promedio", "monto": f"{tc_promedio_12m:,.3f}", "ref": "bcrp"},
            {"descripcion": f"Beneficio ilícito al {fecha_cumplimiento_extemporaneo.strftime('%d/%m/%Y')}", "monto": f"S/ {bi_cap_soles:,.3f}", "ref": None},
            {"descripcion": "Ajuste inflacionario", "monto": f"{ajuste_inflacionario:,.3f}", "ref": "ipc_fecha"},
            {"descripcion": "Beneficio ilícito a la fecha de emisión del informe", "monto": f"S/ {bi_final_soles:,.3f}", "ref": None},
            {"descripcion": f"UIT al año {fecha_hoy.year}", "monto": f"S/ {valor_uit:,.2f}", "ref": None},
            {"descripcion": "Beneficio Ilícito (UIT)", "monto": f"{beneficio_ilicito_uit:,.3f}", "ref": None}
        ]

        # 6. Recolectar todos los datos necesarios para formatear las plantillas de fuentes
        datos_para_fuentes = {
            'rubro': datos_entrada.get('rubro', ''),
            'fuente_cos': fuente_cos,
            # Formateamos las fechas a texto con el nombre de tu placeholder
            'fecha_incumplimiento_texto': fecha_incumplimiento_calc.strftime('%d de %B de %Y'),
            'fecha_extemporanea_texto': fecha_cumplimiento_extemporaneo.strftime('%d de %B de %Y'),
            'mes_actual_texto': fecha_hoy.strftime('%B de %Y'),
            'ultima_fecha_ipc_texto': df_indices.dropna(subset=['Indice_Mes']).sort_values(by='Indice_Mes', ascending=False).iloc[0]['Indice_Mes'].strftime('%B %Y')
        }

        # 7. Devolver la nueva estructura de datos
        return {
            "table_rows": tabla_resumen_filas,
            "footnote_data": datos_para_fuentes,
            "beneficio_ilicito_uit": beneficio_ilicito_uit, # Mantenemos este valor clave
            "error": None
        }
        # ---- FIN DE LA SECCIÓN MODIFICADA ----

    except Exception as e:
        import traceback
        traceback.print_exc()
        return {'error': f"Error en cálculo de BI extemporáneo: {e}"}
    

def calcular_beneficio_ilicito(datos_entrada):
    """
    Realiza el cálculo del BI y devuelve los datos en el nuevo formato estándar
    con claves para fuentes dinámicas.
    """
    try:
        # 1. Desempaquetado y cálculos (la mayoría de esto ya lo tenías)
        df_cos = datos_entrada['df_cos']
        df_uit = datos_entrada['df_uit']
        df_indices = datos_entrada['df_indices']
        rubro = datos_entrada['rubro']
        ce_soles = datos_entrada['ce_soles']
        ce_dolares = datos_entrada['ce_dolares']
        fecha_incumplimiento_calc = datos_entrada['fecha_incumplimiento']
        fecha_calculo = date.today()

        cos_info = df_cos[df_cos['Sector_Rubro'] == rubro]
        if cos_info.empty:
            return {'error': f"No se encontró información de COS para el rubro '{rubro}'."}
        
        fuente_cos = cos_info.iloc[0]['Fuente_COS']
        moneda_cos = str(cos_info.iloc[0]['Moneda_COS']).strip()
        cos_anual = convertir_porcentaje(cos_info.iloc[0]['COS_Anual'])
        cos_mensual = convertir_porcentaje(cos_info.iloc[0]['COS_Mensual'])
        ce = ce_soles if moneda_cos == 'S/' else ce_dolares
        
        diff = relativedelta(fecha_calculo, fecha_incumplimiento_calc)
        t_meses_decimal = (diff.years * 12 + diff.months) + (diff.days / 30.0)
        ce_ajustado = ce * ((1 + cos_mensual) ** t_meses_decimal)
        
        end_date_tc = pd.to_datetime(fecha_calculo)
        start_date_tc = end_date_tc - relativedelta(months=12)
        tc_promedio_df = df_indices[(df_indices['Indice_Mes'] > start_date_tc) & (df_indices['Indice_Mes'] <= end_date_tc)]
        tc_promedio_12m = tc_promedio_df['TC_Mensual'].mean() if not tc_promedio_df.empty else 0
        
        beneficio_ilicito_soles = ce_ajustado if moneda_cos == 'S/' else ce_ajustado * tc_promedio_12m
        
        uit_info = df_uit[df_uit['Año_UIT'] == fecha_calculo.year]
        valor_uit = float(uit_info.iloc[0]['Valor_UIT']) if not uit_info.empty else 0
        beneficio_ilicito_uit = beneficio_ilicito_soles / valor_uit if valor_uit > 0 else 0

        # --- INICIO DE LA NUEVA SECCIÓN ---

        # 2. Definir las filas de la tabla con sus claves de referencia
        tabla_resumen_filas = [
            {"descripcion": "CE para el hecho imputado", "monto": f"{'S/' if moneda_cos == 'S/' else 'US$'} {ce:,.3f}", "ref": "ce_anexo"},
            {"descripcion": "COS (anual)", "monto": f"{cos_anual:,.3%}", "ref": "cok"},
            {"descripcion": "COSm (mensual)", "monto": f"{cos_mensual:,.3%}", "ref": "cok"},
            {"descripcion": "T: meses transcurridos", "monto": f"{t_meses_decimal:,.3f}", "ref": "periodo_bi"},
            {"descripcion": "Costo evitado ajustado", "monto": f"{'S/' if moneda_cos == 'S/' else 'US$'} {ce_ajustado:,.3f}", "ref": None},
            {"descripcion": "Tipo de cambio promedio", "monto": f"{tc_promedio_12m:,.3f}", "ref": "bcrp"},
            {"descripcion": "Beneficio ilícito (S/)", "monto": f"S/ {beneficio_ilicito_soles:,.3f}", "ref": None},
            {"descripcion": f"UIT al año {fecha_calculo.year}", "monto": f"S/ {valor_uit:,.2f}", "ref": None},
            {"descripcion": "Beneficio Ilícito (UIT)", "monto": f"{beneficio_ilicito_uit:,.3f}", "ref": None}
        ]

        # 3. Recolectar datos para formatear las fuentes
        datos_para_fuentes = {
            'rubro': rubro,
            'fuente_cos': fuente_cos,
            # Formateamos las fechas a texto con el nombre de tu placeholder
            'fecha_incumplimiento_texto': fecha_incumplimiento_calc.strftime('%d de %B de %Y'),
            'fecha_hoy_texto': fecha_calculo.strftime('%d de %B de %Y'),
        }

        # 4. Devolver la nueva estructura de datos
        return {
            "table_rows": tabla_resumen_filas,
            "footnote_data": datos_para_fuentes,
            "beneficio_ilicito_uit": beneficio_ilicito_uit,
            "fuente_cos": fuente_cos, # Mantenemos estos para compatibilidad si es necesario
            "cos_anual": cos_anual,
            "cos_mensual": cos_mensual,
            "moneda_cos": moneda_cos,
            "error": None
        }
        # --- FIN DE LA NUEVA SECCIÓN ---

    except Exception as e:
        import traceback
        traceback.print_exc()
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
