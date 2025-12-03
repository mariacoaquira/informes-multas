# --- Archivo: producto_coercitiva.py ---

# --- BIBLIOTECAS ---
import streamlit as st
import pandas as pd
from datetime import date
from babel.dates import format_date
from docxtpl import DocxTemplate, RichText
from sheets import cargar_hoja_a_df, descargar_archivo_drive
from funciones import create_main_table_coercitiva, create_summary_table_subdoc, texto_con_numero
import io
import math

# ---------------------------------------------------------------------
# FUNCIÓN 1: RENDERIZAR INTERFAZ DE USUARIO (CORREGIDO CON RD1 y RD2)
# ---------------------------------------------------------------------
def renderizar_inputs_coercitiva(datos_informe):
    """ Renderiza la interfaz para multa coercitiva, centrada en medidas y resoluciones específicas. """
    st.header("Paso 2: Datos para Multa Coercitiva")
    
    # --- SECCIÓN A: RESOLUCIÓN DIRECTORAL DE SANCIÓN (RD1) ---
    st.subheader("A. Resolución Directoral de Sanción (RD1)")
    st.caption("Resolución donde se dictó la medida correctiva.")
    
    col_rd1, col_rd2 = st.columns(2)
    with col_rd1: 
        datos_informe['numero_rd1'] = st.text_input("N.° Resolución Directoral (RD1)", value=datos_informe.get('numero_rd1', ''))
    with col_rd2: 
        datos_informe['fecha_rd1'] = st.date_input("Fecha de notificación de RD1", value=datos_informe.get('fecha_rd1'), format="DD/MM/YYYY")
    
    col_rd3, col_rd4, col_rd5 = st.columns(3)
    with col_rd3:
        datos_informe['num_hechos_rd1'] = st.number_input("N.° Total de Hechos en RD1", min_value=1, step=1, value=datos_informe.get('num_hechos_rd1', 1))
    with col_rd4:
        datos_informe['multa_total_rd1'] = st.number_input("Multa Total impuesta en RD1 (UIT)", min_value=0.0, step=0.001, format="%.3f", value=datos_informe.get('multa_total_rd1', 0.0))
    with col_rd5:
        datos_informe['num_medidas_rd1'] = st.number_input("N.° Total de Medidas en RD1", min_value=1, step=1, value=datos_informe.get('num_medidas_rd1', 1))
    st.divider()

    # --- SECCIÓN B: RESOLUCIÓN DE INCUMPLIMIENTO (RD2) ---
    st.subheader("B. Resolución que declara el Incumplimiento (RD2)")
    st.caption("Resolución donde se verificó el incumplimiento de la medida.")
    
    col_rd2_1, col_rd2_2 = st.columns(2)
    with col_rd2_1: 
        datos_informe['numero_rd2'] = st.text_input("N.° Resolución (RD2)", value=datos_informe.get('numero_rd2', ''))
    with col_rd2_2: 
        datos_informe['fecha_rd2'] = st.date_input("Fecha de notificación de RD2", value=datos_informe.get('fecha_rd2'), format="DD/MM/YYYY")
    st.divider()

    # --- SECCIÓN C: MEDIDAS CORRECTIVAS INCUMPLIDAS ---
    st.subheader("C. Medidas Correctivas Incumplidas")
    st.caption("Añade cada medida correctiva incumplida y, dentro de ella, las conductas infractoras (hechos) de la RD que la originaron.")
    
    if 'medidas_incumplidas' not in datos_informe:
        datos_informe['medidas_incumplidas'] = [
            {'num_medida': '1', 'desc_medida': '', 'obligacion': '', 'hechos_asociados': [
                {'num_hecho': '1', 'desc_hecho': '', 'multa_uit_rd': 0.0}
            ]}
        ]

    if st.button("➕ Añadir Medida Correctiva"):
        next_num = len(datos_informe['medidas_incumplidas']) + 1
        datos_informe['medidas_incumplidas'].append(
            {'num_medida': str(next_num), 'desc_medida': '', 'obligacion': '', 'hechos_asociados': [
                {'num_hecho': '1', 'desc_hecho': '', 'multa_uit_rd': 0.0}
            ]}
        )
        st.rerun()

    for i, medida in enumerate(datos_informe['medidas_incumplidas']):
        with st.container(border=True):
            st.markdown(f"#### Medida Correctiva N° {i+1}")
            
            # --- INICIO CAMBIO PUNTUAL ---
            col_m1, col_m2 = st.columns([1, 4])
            with col_m1: 
                medida['num_medida'] = st.text_input("N°/ID", value=medida.get('num_medida', str(i+1)), key=f"med_num_{i}")
            with col_m2: 
                # Usamos 'desc_medida' como el campo principal para toda la descripción/obligación
                medida['desc_medida'] = st.text_area(
                    "Medida Correctiva (Descripción y Obligación)", 
                    value=medida.get('desc_medida', ''), 
                    key=f"med_desc_{i}", 
                    height=100, 
                    placeholder="Ingrese la descripción completa de la medida..."
                )
                # Eliminamos el input separado de 'obligacion'
            # --- FIN CAMBIO PUNTUAL ---

            st.markdown("**Conductas Infractoras (Hechos) Vinculadas:**")
            
            for j, hecho in enumerate(medida['hechos_asociados']):
                with st.container():
                    col_h1, col_h2, col_h3 = st.columns([1, 4, 1.5])
                    with col_h1: 
                        hecho['num_hecho'] = st.text_input("N° Hecho (RD)", value=hecho.get('num_hecho', str(j+1)), key=f"m{i}_h{j}_num")
                    with col_h2: 
                        hecho['desc_hecho'] = st.text_area("Descripción Hecho (RD)", value=hecho.get('desc_hecho', ''), key=f"m{i}_h{j}_desc", height=68)
                    with col_h3: 
                        hecho['multa_uit_rd'] = st.number_input("Multa Base (UIT)", min_value=0.0, value=hecho.get('multa_uit_rd', 0.0), format="%.3f", key=f"m{i}_h{j}_uit")
                    
                    if len(medida['hechos_asociados']) > 1:
                        if st.button(f"➖ Quitar Hecho", key=f"med_{i}_del_hecho_{j}"):
                            medida['hechos_asociados'].pop(j); st.rerun()
                    st.divider()

            if st.button(f"➕ Añadir Hecho a Medida {i+1}", key=f"med_{i}_add_h"):
                medida['hechos_asociados'].append({'num_hecho': str(len(medida['hechos_asociados'])+1), 'desc_hecho': '', 'multa_uit_rd': 0.0})
                st.rerun()

            if len(datos_informe['medidas_incumplidas']) > 1:
                if st.button(f"❌ Quitar Medida {i+1}", key=f"del_med_{i}", type="secondary"):
                    datos_informe['medidas_incumplidas'].pop(i); st.rerun()
    st.divider()

    # --- SECCIÓN D: CONFIGURACIÓN DEL CÁLCULO ---
    st.subheader("D. Configuración del Cálculo")
    
    col_conf1, col_conf2 = st.columns(2)
    with col_conf1:
        num_coercitiva = st.number_input("Número de Multa Coercitiva a imponer", min_value=1, step=1, value=datos_informe.get('num_coercitiva', 1))
        datos_informe['num_coercitiva'] = num_coercitiva
    
    with col_conf2:
        if num_coercitiva == 1:
            metodologias = ["Nueva", "Antigua"]
            datos_informe['metodologia'] = st.radio("Metodología:", metodologias, index=0 if datos_informe.get('metodologia') == "Nueva" else 1, horizontal=True)
        else:
            st.info("Para multas sucesivas (>1), se requiere el monto anterior.")
            datos_informe['metodologia'] = None

    # --- Historial de Coercitivas Anteriores (Solo si > 1) ---
    if num_coercitiva > 1:
        st.markdown("###### Historial de Multas Coercitivas Previas")
        if 'historial_previo' not in datos_informe:
            datos_informe['historial_previo'] = []
        
        # Asegurar filas para 1 a N-1
        num_previas = num_coercitiva - 1
        while len(datos_informe['historial_previo']) < num_previas:
            datos_informe['historial_previo'].append({})
            
        for k in range(num_previas):
            st.caption(f"Datos de la {k+1}ª Multa Coercitiva (Global):")
            c_hist1, c_hist2, c_hist3 = st.columns(3)
            prev = datos_informe['historial_previo'][k]
            with c_hist1: prev['num_res'] = st.text_input(f"N° Resolución ({k+1}ª)", value=prev.get('num_res', ''), key=f"hist_num_{k}")
            with c_hist2: prev['fecha'] = st.date_input(f"Fecha Emisión ({k+1}ª)", value=prev.get('fecha'), key=f"hist_fec_{k}", format="DD/MM/YYYY")
            with c_hist3: prev['monto'] = st.number_input(f"Monto Impuesto ({k+1}ª) UIT", value=prev.get('monto', 0.0), format="%.3f", key=f"hist_mnt_{k}")
    
    # --- E. CUADRO RESUMEN EN LA APP ---
    st.divider()
    st.subheader("Resumen: Medidas Correctivas y Conductas infractoras Vinculadas")
    resumen_data = []
    for medida in datos_informe.get('medidas_incumplidas', []):
        if medida.get('num_medida'):
            medida_key = f"N° {medida['num_medida']}: {medida.get('desc_medida', '(Sin descripción)')}"
            hechos_str_list = []
            multa_base_sumada = 0.0
            for hecho in medida.get('hechos_asociados', []):
                hechos_str_list.append(f"- Conducta infractora N° {hecho.get('num_hecho', '?')}: {hecho.get('desc_hecho', 'N/A')}")
                multa_base_sumada += pd.to_numeric(hecho.get('multa_uit_rd', 0.0), errors='coerce')
            resumen_data.append({"Medida Correctiva Incumplida": f"**{medida_key}**", "Hechos Originales Vinculados": "\n".join(hechos_str_list), "Multa Base para Coercitiva (UIT)": f"{multa_base_sumada:,.3f}"})

    if resumen_data: st.dataframe(pd.DataFrame(resumen_data), use_container_width=True, hide_index=True)
    else: st.info("Aún no se han añadido medidas incumplidas.")
            
    return datos_informe

# ---------------------------------------------------------------------
# FUNCIÓN 2: VALIDACIÓN (CORREGIDA CON RD1 Y RD2)
# ---------------------------------------------------------------------
def validar_inputs_coercitiva(datos_informe):
    # 1. Validar RD1
    if not all([datos_informe.get('numero_rd1'), datos_informe.get('fecha_rd1')]): return False
    
    # 2. Validar RD2
    if not all([datos_informe.get('numero_rd2'), datos_informe.get('fecha_rd2')]): return False
    
    # 3. Validar Medidas
    medidas = datos_informe.get('medidas_incumplidas', [])
    if not medidas: return False
    for m in medidas:
        if not m.get('num_medida'): return False
        for h in m.get('hechos_asociados', []):
            if not h.get('num_hecho') or pd.to_numeric(h.get('multa_uit_rd'), errors='coerce') <= 0:
                return False
                
    # 4. Validar Historial (si aplica)
    if datos_informe.get('num_coercitiva', 1) > 1:
        historial = datos_informe.get('historial_previo', [])
        for item in historial:
            if not item.get('monto') or item.get('monto') <= 0:
                return False
    return True

# ---------------------------------------------------------------------
# FUNCIÓN 3: HELPER BÚSQUEDA
# ---------------------------------------------------------------------
def _buscar_en_cuadro(multa_base_uit, cuadro_df):
    if cuadro_df is None or cuadro_df.empty: raise ValueError("Cuadro vacío.")
    req = ['Rango_Min_UIT', 'Rango_Max_UIT', 'Coercitiva_Primera_UIT']
    if not all(c in cuadro_df.columns for c in req): raise ValueError("Faltan columnas en cuadro.")
    for c in req: cuadro_df[c] = pd.to_numeric(cuadro_df[c], errors='coerce')
    cuadro_df = cuadro_df.dropna(subset=['Rango_Min_UIT'])
    for _, row in cuadro_df.iterrows():
        min_r, max_r = row['Rango_Min_UIT'], row['Rango_Max_UIT']
        if pd.isna(max_range := row['Rango_Max_UIT']): 
            if multa_base_uit >= min_r: return row['Coercitiva_Primera_UIT']
        elif min_r <= multa_base_uit < max_r: return row['Coercitiva_Primera_UIT']
    raise ValueError(f"Multa base ({multa_base_uit}) fuera de rango.")

# ---------------------------------------------------------------------
# FUNCIÓN AUXILIAR GLOBAL: FORMATO DE HECHOS
# ---------------------------------------------------------------------
def fmt_hechos(lista_numeros):
    """Helper para formatear listas de hechos."""
    if not lista_numeros: return ""
    lista_fmt = [f"n.° {n}" for n in lista_numeros]
    if len(lista_fmt) == 1: return lista_fmt[0]
    if len(lista_fmt) == 2: return f"{lista_fmt[0]} y {lista_fmt[1]}"
    return ", ".join(lista_fmt[:-1]) + " y " + lista_fmt[-1]

# ---------------------------------------------------------------------
# FUNCIÓN AUXILIAR: PLACEHOLDERS GRAMATICALES
# ---------------------------------------------------------------------
def _generar_placeholders_gramaticales(num_coercitiva, medidas_incumplidas, todos_hechos_unicos):
    num_medidas = len(medidas_incumplidas)
    num_hechos_total = len(todos_hechos_unicos)
    placeholders = {}

    # 1. Primera multa
    if num_medidas == 1:
        ordinales_sg = {1: "primera", 2: "segunda", 3: "tercera", 4: "cuarta", 5: "quinta"}
        ord_txt = ordinales_sg.get(num_coercitiva, f"{num_coercitiva}a")
        placeholders['ph_primera_multa'] = f"la {ord_txt} multa coercitiva"
    else:
        ordinales_pl = {1: "primeras", 2: "segundas", 3: "terceras", 4: "cuartas", 5: "quintas"}
        ord_txt = ordinales_pl.get(num_coercitiva, f"{num_coercitiva}as")
        placeholders['ph_primera_multa'] = f"las {ord_txt} multas coercitivas"

    # 2. Detalle medidas hechos
    lista_descripciones = []
    for medida in medidas_incumplidas:
        hechos_medida = [h.get('num_hecho') for h in medida.get('hechos_asociados', [])]
        texto_hechos = fmt_hechos(hechos_medida)
        prefix_hecho = "la conducta infractora" if len(hechos_medida) == 1 else "las conductas infractoras"
        desc = f"una (1) medida correctiva ordenada a {prefix_hecho} {texto_hechos}"
        lista_descripciones.append(desc)
    
    if len(lista_descripciones) == 1: placeholders['ph_detalle_medidas_hechos'] = lista_descripciones[0]
    elif len(lista_descripciones) == 2: placeholders['ph_detalle_medidas_hechos'] = f"{lista_descripciones[0]} y {lista_descripciones[1]}"
    else: placeholders['ph_detalle_medidas_hechos'] = "; ".join(lista_descripciones[:-1]) + " y " + lista_descripciones[-1]

    # Resto de placeholders gramaticales...
    placeholders['ph_persistencia'] = "la persistencia en " if num_coercitiva > 1 else ""
    placeholders['ph_medida_ordenada_qty'] = "la medida correctiva ordenada" if num_medidas == 1 else "las medidas correctivas ordenadas"
    placeholders['ph_conducta_infractora_qty'] = "la conducta infractora" if num_hechos_total == 1 else "las conductas infractoras"
    placeholders['ph_medida_correctiva_qty'] = "la medida correctiva" if num_medidas == 1 else "las medidas correctivas"
    
    todos_hechos_str = fmt_hechos(list(todos_hechos_unicos))
    prefix_hecho_global = "la conducta infractora" if num_hechos_total == 1 else "las conductas infractoras"
    if num_medidas == 1: placeholders['ph_unica_medida_asociada'] = f"la única medida correctiva asociada a {prefix_hecho_global} {todos_hechos_str}"
    else: placeholders['ph_unica_medida_asociada'] = f"las medidas correctivas asociadas a {prefix_hecho_global} {todos_hechos_str}"

    qty_txt_multa = texto_con_numero(num_medidas, genero='f')
    suffix_multa_qty = "multa coercitiva" if num_medidas == 1 else "multas coercitivas"
    placeholders['ph_cantidad_multas_coercitivas'] = f"{qty_txt_multa} {suffix_multa_qty}"
    placeholders['ph_multa_coercitiva_qty'] = "la multa coercitiva" if num_medidas == 1 else "las multas coercitivas"

    if num_medidas == 1: placeholders['ph_se_le_impondra_multas'] = "se le impondrá una (1) multa coercitiva"
    else: placeholders['ph_se_le_impondra_multas'] = f"se le impondrán {qty_txt_multa} multas coercitivas"
        
    placeholders['ph_unica_medida_correctiva'] = "la única medida correctiva" if num_medidas == 1 else "las medidas correctivas"
    placeholders['ph_conducta_infractora_asociada_describe'] = "La conducta infractora asociada se describe" if num_hechos_total == 1 else "Las conductas infractoras asociadas se describen"
    placeholders['ph_medida_correctiva_mencionada'] = "la medida correctiva mencionada" if num_medidas == 1 else "las medidas correctivas mencionadas"

    return placeholders

# ---------------------------------------------------------------------
# FUNCIÓN 4: PROCESAMIENTO Y TABLAS
# ---------------------------------------------------------------------
def procesar_coercitiva(datos_comunes, datos_informe):
    try:
        cliente_gs = datos_comunes.get('cliente_gspread'); nombre_gsheet = "Base de datos"
        if not cliente_gs: return {'error': 'No hay conexión a Google Sheets.'}
        df_criterios_total = cargar_hoja_a_df(cliente_gs, nombre_gsheet, "Criterios_Coercitiva")
        if df_criterios_total is None: return {'error': "No se pudo cargar la hoja 'Criterios_Coercitiva'."}

        medidas_agrupadas = {}
        hechos_rd_flat_list = []
        multa_total_rd_uit = 0.0
        todos_hechos_unicos = set()
        
        for medida_input in datos_informe.get('medidas_incumplidas', []):
            num_med = medida_input.get('num_medida')
            desc_med = medida_input.get('desc_medida', '')
            if not num_med: continue
            
            # --- CAMBIO: La obligación ahora es la misma descripción ---
            # Al haber unificado el input, usamos desc_med como la obligación para la Tabla 6
            obligacion = desc_med 
            # -----------------------------------------------------------

            multa_base_para_esta_medida = 0.0
            hechos_originales_para_medida = []

            for hecho_input in medida_input.get('hechos_asociados', []):
                multa_rd = pd.to_numeric(hecho_input.get('multa_uit_rd', 0.0), errors='coerce') or 0.0
                num_hecho = hecho_input.get('num_hecho')
                if num_hecho: todos_hechos_unicos.add(num_hecho)
                
                multa_base_para_esta_medida += multa_rd
                multa_total_rd_uit += multa_rd

                hechos_rd_flat_list.append({
                    'num_hecho': num_hecho,
                    'desc_hecho': hecho_input.get('desc_hecho'),
                    'multa_uit_rd': multa_rd,
                    'num_medida': num_med,
                    'desc_medida': desc_med
                })
                hechos_originales_para_medida.append(hecho_input)

            medidas_agrupadas[num_med] = {
                'hechos': hechos_originales_para_medida, 
                'multa_base_total_uit': multa_base_para_esta_medida,
                'desc_medida': desc_med,
                'obligacion': obligacion
            }
        
        # 3. Calcular Multa Coercitiva
        resultados_medidas = []; multa_coercitiva_total_uit = 0.0
        MAX_COERCITIVA_UIT = 100
        num_coercitiva_global = datos_informe.get('num_coercitiva', 1)
        
        num_ordinal = {1: "Primera", 2: "Segunda", 3: "Tercera", 4: "Cuarta", 5: "Quinta"}
        num_coercitiva_texto = num_ordinal.get(num_coercitiva_global, f"{num_coercitiva_global}a")
        titulo_coercitiva = f"{num_coercitiva_texto} Multa Coercitiva"
        
        total_medidas_a_procesar = len(medidas_agrupadas)

        for num_med, data in medidas_agrupadas.items():
            multa_base = data['multa_base_total_uit']; multa_coercitiva = 0.0
            
            if num_coercitiva_global == 1:
                metodologia = datos_informe.get('metodologia')
                cuadro_sel = df_criterios_total[df_criterios_total['Metodologia'].str.strip().str.lower() == metodologia.lower()].copy()
                try: multa_coercitiva = _buscar_en_cuadro(multa_base, cuadro_sel)
                except ValueError as e: st.error(f"Error Medida '{num_med}': {e}")
            else:
                uit_ant = datos_informe.get('uit_anterior_por_medida', {}).get(num_med, 0.0)
                if uit_ant > 0: multa_coercitiva = uit_ant * 2
                else: st.error(f"Monto anterior no encontrado o inválido para Medida '{num_med}'.")
            
            multa_coercitiva_final = min(multa_coercitiva, MAX_COERCITIVA_UIT)
            
            if total_medidas_a_procesar == 1: label_medida_dinamica = "Única medida correctiva"
            else: label_medida_dinamica = f"Medida correctiva n.° {num_med}"

            resultados_medidas.append({
                'num_medida': num_med, 
                'desc_medida': data.get('desc_medida', ''), 
                'label_medida': label_medida_dinamica,
                'obligacion': data.get('obligacion', ''),
                'hechos': data.get('hechos', []), 
                'multa_base_uit': multa_base, 
                'num_coercitiva_texto': num_coercitiva_texto, 
                'multa_coercitiva_final_uit': multa_coercitiva_final
            })
            multa_coercitiva_total_uit += multa_coercitiva_final

        # 4. Cargar Plantilla
        df_productos = datos_comunes.get('df_productos'); id_plantilla = df_productos[df_productos['Producto'] == 'COERCITIVA'].iloc[0].get('ID_Plantilla_Inicio')
        if not id_plantilla: return {'error': 'No se encontró ID de plantilla para COERCITIVA.'}
        buffer_plantilla = descargar_archivo_drive(id_plantilla);
        if not buffer_plantilla: return {'error': 'No se pudo descargar la plantilla Coercitiva.'}
        doc_tpl = DocxTemplate(buffer_plantilla)

        # 5. Preparar Tablas
        contexto_modificado = datos_comunes.get('context_data', {}).copy()
        df_subdirecciones = cargar_hoja_a_df(cliente_gs, nombre_gsheet, "Subdirecciones")
        if df_subdirecciones is not None:
            dfai_row = df_subdirecciones[df_subdirecciones['ID_Subdireccion'] == 'DFAI']
            if not dfai_row.empty:
                contexto_modificado['nombre_encargado_sub1'] = dfai_row.iloc[0].get('Encargado_Sub', '')
                contexto_modificado['cargo_encargado_sub1'] = dfai_row.iloc[0].get('Cargo_Encargado_Sub', '')
                contexto_modificado['titulo_encargado_sub1'] = dfai_row.iloc[0].get('Titulo_Encargado_Sub', '')
                contexto_modificado['subdireccion'] = dfai_row.iloc[0].get('Subdireccion', 'Dirección de Fiscalización y Aplicación de Incentivos')
        
        # Tabla 4 (Req): N°, Conducta, Medida
        datos_t4 = [{'n': x['num_hecho'], 'hecho': x['desc_hecho'], 'medida': f"Medida N° {x['num_medida']}: {x['desc_medida']}"} for x in hechos_rd_flat_list]
        t4_subdoc = create_main_table_coercitiva(doc_tpl, ["N°", "Conducta Infractora", "Medida Correctiva"], datos_t4, ['n', 'hecho', 'medida'], column_widths=(0.5, 3, 3))
        
        # Tabla 5 (Req): Conducta, Multa Final (Total)
        datos_t5 = [{'hecho': f"Hecho N° {x['num_hecho']}: {x['desc_hecho']}", 'multa': f"{x['multa_uit_rd']:,.3f} UIT"} for x in hechos_rd_flat_list]
        datos_t5.append({'hecho': 'Total', 'multa': f"{multa_total_rd_uit:,.3f} UIT"})
        t5_subdoc = create_main_table_coercitiva(doc_tpl, ["Conducta Infractora", "Multa Final"], datos_t5, ['hecho', 'multa'], column_widths=(5, 1.5))
        
        # Tabla 6 (Req): N° Conducta, Obligación
        datos_t6 = []
        for num_med, data in medidas_agrupadas.items():
            # AQUÍ USAMOS LA NUEVA VARIABLE 'obligacion' QUE CONTIENE LA DESCRIPCIÓN
            oblig = data.get('obligacion', '')
            for h in data.get('hechos', []):
                texto_obl = f"{oblig}\n\n(Multa Base asociada: {float(h.get('multa_uit_rd', 0)):,.3f} UIT)"
                datos_t6.append({'n': h['num_hecho'], 'obl': texto_obl})
        t6_subdoc = create_main_table_coercitiva(doc_tpl, ["N° Conducta", "Obligación de la Medida"], datos_t6, ['n', 'obl'], column_widths=(1, 5.5))
        
        # Tabla 7 (Req): N°, Medida, Monto Coercitiva (Total - Usando Summary Table)
        datos_t7 = [{'Numeral': x['num_medida'], 'Medida': x['desc_medida'], 'Monto': f"{x['multa_coercitiva_final_uit']:,.3f} UIT"} for x in resultados_medidas]
        datos_t7.append({'Numeral': '', 'Medida': 'Total', 'Monto': f"{multa_coercitiva_total_uit:,.3f} UIT"})
        t7_subdoc = create_summary_table_subdoc(doc_tpl, ["N°", "Medida Correctiva", "Monto Multa Coercitiva"], datos_t7, ['Numeral', 'Medida', 'Monto'], column_widths=(0.5, 4, 2))

        # 6. Contexto final
        placeholders_gramaticales = _generar_placeholders_gramaticales(datos_informe.get('num_coercitiva', 1), datos_informe.get('medidas_incumplidas', []), todos_hechos_unicos)
        
        contexto_final = {
            **contexto_modificado, 
            **placeholders_gramaticales, 
            'numero_coercitiva_titulo': titulo_coercitiva,
            'numero_rd1': datos_informe.get('numero_rd1'),
            'fecha_rd1': format_date(datos_informe.get('fecha_rd1'), "d 'de' MMMM 'de' yyyy", locale='es') if datos_informe.get('fecha_rd1') else '',
            'num_hechos_rd1': datos_informe.get('num_hechos_rd1'),
            'multa_total_rd1': f"{datos_informe.get('multa_total_rd1'):,.3f}",
            'num_medidas_rd1': datos_informe.get('num_medidas_rd1'),
            'numero_rd2': datos_informe.get('numero_rd2'),
            'fecha_rd2': format_date(datos_informe.get('fecha_rd2'), "d 'de' MMMM 'de' yyyy", locale='es') if datos_informe.get('fecha_rd2') else '',
            
            'multa_base_calculada': f"{multa_total_rd_uit:,.3f}", 
            'multa_base_calculada_uit': f"{multa_total_rd_uit:,.3f} UIT",
            'multa_coercitiva_total': f"{multa_coercitiva_total_uit:,.3f}",
            'multa_coercitiva_total_uit': f"{multa_coercitiva_total_uit:,.3f} UIT",
            
            'es_metodologia_nueva': (datos_informe.get('num_coercitiva') == 1 and datos_informe.get('metodologia') == 'Nueva'),
            
            'lista_medidas_calculadas': resultados_medidas,
            
            'tabla_hecho_medida': t4_subdoc,
            'tabla_hecho_multa': t5_subdoc,
            'tabla_hecho_obligacion': t6_subdoc,
            'tabla_medida_coercitiva': t7_subdoc,
        }

        doc_tpl.render(contexto_final, autoescape=True)
        buffer_final_word = io.BytesIO()
        doc_tpl.save(buffer_final_word)
        
        return {'doc_pre_compuesto': buffer_final_word, 'resultados_para_app': {'tabla_resumen_coercitivas': resultados_medidas, 'total_uit': multa_coercitiva_total_uit}}

    except Exception as e:
        import traceback; traceback.print_exc(); return {'error': f"Error fatal al procesar Multa Coercitiva: {e}"}