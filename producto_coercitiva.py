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
# FUNCIÓN 1: RENDERIZAR INTERFAZ DE USUARIO
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
    
    # --- INICIO CORRECCIÓN: Ambos artículos en RD1 ---
    col_art1, col_art2 = st.columns(2)
    with col_art1:
        datos_informe['articulo_resp_rd1'] = st.text_input(
            "Artículo que declara la responsabilidad/incumplimiento", 
            value=datos_informe.get('articulo_resp_rd1', ''), 
            help="Ej: Artículo 1° (donde se dice que el administrado es responsable).",
            placeholder="Ej: Artículo 1°"
        )
    with col_art2:
        datos_informe['articulo_medida_rd1'] = st.text_input(
            "Artículo que ordena la medida correctiva", 
            value=datos_informe.get('articulo_medida_rd1', ''), 
            help="Ej: Artículo 2° (donde se dicta la medida).",
            placeholder="Ej: Artículo 2°"
        )
    # --- FIN CORRECCIÓN ---
    
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
    
    # --- NUEVO: Artículo de Incumplimiento ---
    datos_informe['articulo_rd2'] = st.text_input("Artículo que declara el incumplimiento (ej: Artículo 1°)", value=datos_informe.get('articulo_rd2', 'Artículo X°'), help="Se usará para indicar dónde se declaró el incumplimiento.")
    st.divider()

    # --- SECCIÓN C: MEDIDAS CORRECTIVAS INCUMPLIDAS ---
    st.subheader("C. Medidas Correctivas Incumplidas")
    st.caption("Añade cada medida correctiva incumplida y, dentro de ella, las conductas infractoras (hechos) de la RD que la originaron.")
    
    if 'medidas_incumplidas' not in datos_informe:
        datos_informe['medidas_incumplidas'] = [
            {'num_medida': '1', 'desc_medida': '', 'hechos_asociados': [
                {'num_hecho': '1', 'desc_hecho': '', 'multa_uit_rd': 0.0}
            ]}
        ]

    if st.button("➕ Añadir Medida Correctiva"):
        next_num = len(datos_informe['medidas_incumplidas']) + 1
        datos_informe['medidas_incumplidas'].append(
            {'num_medida': str(next_num), 'desc_medida': '', 'hechos_asociados': [
                {'num_hecho': '1', 'desc_hecho': '', 'multa_uit_rd': 0.0}
            ]}
        )
        st.rerun()

    for i, medida in enumerate(datos_informe['medidas_incumplidas']):
        with st.container(border=True):
            st.markdown(f"#### Medida Correctiva N° {i+1}")
            
            col_m1, col_m2 = st.columns([1, 4])
            with col_m1: 
                medida['num_medida'] = st.text_input("N°/ID", value=medida.get('num_medida', str(i+1)), key=f"med_num_{i}")
            with col_m2: 
                medida['desc_medida'] = st.text_area(
                    "Medida Correctiva (Descripción y Obligación)", 
                    value=medida.get('desc_medida', ''), 
                    key=f"med_desc_{i}", 
                    height=100, 
                    placeholder="Ingrese la descripción completa de la medida..."
                )

            st.caption("Conductas Infractoras (Hechos) Vinculadas:")
            
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
# FUNCIÓN 2: VALIDACIÓN
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
    """Helper para formatear listas de hechos (n.° 1, n.° 2...)."""
    if not lista_numeros: return ""
    lista_fmt = [f"n.° {n}" for n in lista_numeros]
    if len(lista_fmt) == 1: return lista_fmt[0]
    if len(lista_fmt) == 2: return f"{lista_fmt[0]} y {lista_fmt[1]}"
    return ", ".join(lista_fmt[:-1]) + " y " + lista_fmt[-1]

# ---------------------------------------------------------------------
# FUNCIÓN AUXILIAR: PLACEHOLDERS GRAMATICALES (GLOBALES)
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

    # 2. Detalle medidas hechos (Complejo)
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
# FUNCIÓN 4: PROCESAMIENTO Y TABLAS (Solo 3 Tablas Solicitadas)
# ---------------------------------------------------------------------
def procesar_coercitiva(datos_comunes, datos_informe):
    try:
        cliente_gs = datos_comunes.get('cliente_gspread')
        # Cargar criterios (para búsqueda en tabla de rangos si es 1ra coercitiva)
        df_criterios_total = cargar_hoja_a_df(cliente_gs, "Base de datos", "Criterios_Coercitiva")
        if df_criterios_total is None: return {'error': "Error carga Criterios."}

        # --- 1. PROCESAMIENTO DE DATOS ---
        medidas_agrupadas = {}
        hechos_rd_flat_list = []
        multa_total_rd_uit = 0.0
        todos_hechos_unicos = set()
        
        # Aplanar estructura para tablas y cálculos
        for medida_input in datos_informe.get('medidas_incumplidas', []):
            num_med = medida_input.get('num_medida')
            desc_med = medida_input.get('desc_medida', '')
            if not num_med: continue
            
            obligacion = desc_med # La obligación es la misma descripción
            multa_base_para_esta_medida = 0.0
            hechos_originales_para_medida = []

            for h_in in medida_input.get('hechos_asociados', []):
                m_rd = pd.to_numeric(h_in.get('multa_uit_rd', 0.0), errors='coerce') or 0.0
                num_hecho = h_in.get('num_hecho')
                if num_hecho: todos_hechos_unicos.add(num_hecho)
                
                multa_base_para_esta_medida += m_rd
                multa_total_rd_uit += m_rd

                # Lista plana para Tabla 1 y 2
                hechos_rd_flat_list.append({
                    'num_hecho': num_hecho,
                    'desc_hecho': h_in.get('desc_hecho'),
                    'multa_uit_rd': m_rd,
                    'num_medida': num_med,
                    'desc_medida': desc_med
                })
                hechos_originales_para_medida.append(h_in)

            medidas_agrupadas[num_med] = {
                'hechos': hechos_originales_para_medida, 
                'multa_base_total_uit': multa_base_para_esta_medida,
                'desc_medida': desc_med,
                'obligacion': obligacion
            }
        
        # --- 2. CÁLCULO DE MULTA COERCITIVA ---
        resultados_medidas = []; multa_coercitiva_total_uit = 0.0
        MAX_UIT = 100
        num_coer = datos_informe.get('num_coercitiva', 1)
        ords = {1: "Primera", 2: "Segunda", 3: "Tercera", 4: "Cuarta", 5: "Quinta"}
        num_coer_txt = ords.get(num_coer, f"{num_coer}a")
        titulo_coercitiva = f"{num_coer_txt} Multa Coercitiva"
        total_medidas_count = len(medidas_agrupadas)

        for num_med, data in medidas_agrupadas.items():
            m_base = data['multa_base_total_uit']; m_coer = 0.0
            
            # Lógica de cálculo (Nueva/Antigua/Historial)
            if num_coer == 1:
                metodo = datos_informe.get('metodologia')
                df_sel = df_criterios_total[df_criterios_total['Metodologia'].str.strip().str.lower() == metodo.lower()].copy()
                try: m_coer = _buscar_en_cuadro(m_base, df_sel)
                except ValueError as e: st.error(f"Error Medida '{num_med}': {e}")
            else:
                m_ant = datos_informe.get('uit_anterior_por_medida', {}).get(num_med, 0.0)
                m_coer = min(m_ant * 2, 100.0)
            
            m_coer_final = min(m_coer, MAX_UIT)
            label_dyn = "Única medida correctiva" if total_medidas_count == 1 else f"Medida correctiva n.° {num_med}"
            
            # Placeholders gramaticales específicos del bucle
            nums_h = [h.get('num_hecho') for h in data.get('hechos', [])]
            h_loop_str = f"{'conducta infractora' if len(nums_h)==1 else 'conductas infractoras'} {fmt_hechos(nums_h)}"
            label_cond = "La conducta infractora" if len(nums_h) == 1 else "Las conductas infractoras"

            resultados_medidas.append({
                'num_medida': num_med, 
                'desc_medida': data['desc_medida'], 
                'label_medida': label_dyn, 
                'obligacion': data['obligacion'],
                'hechos_lista_str': h_loop_str, 
                'multa_base_formato': f"{m_base:,.3f} UIT", 
                'ph_conducta_label': label_cond,
                'multa_base_uit': m_base, 
                'num_coercitiva_texto': num_coer_txt, 
                'multa_coercitiva_final_uit': m_coer_final
            })
            multa_coercitiva_total_uit += m_coer_final

        # --- 3. GENERACIÓN DE DOCUMENTO Y TABLAS ---
        id_tpl = datos_comunes['df_productos'][datos_comunes['df_productos']['Producto'] == 'COERCITIVA'].iloc[0].get('ID_Plantilla_Inicio')
        doc_tpl = DocxTemplate(descargar_archivo_drive(id_tpl))
        
        total_conductas = len(hechos_rd_flat_list)

        # === TABLA 1: MULTA IMPUESTA (Conducta | Multa Final) ===
        datos_tabla_1 = []
        for x in hechos_rd_flat_list:
            # Formato Conducta: Negrita/Subrayado para título + Texto Normal
            rt_cond = RichText()
            titulo_cond = "Única conducta infractora:" if total_conductas == 1 else f"Conducta infractora n.° {x['num_hecho']}:"
            rt_cond.add(titulo_cond, bold=True, underline=True)
            rt_cond.add(f"\n{x['desc_hecho']}")
            
            datos_tabla_1.append({
                'conducta': rt_cond,
                'multa': f"{x['multa_uit_rd']:,.3f} UIT"
            })
        
        datos_tabla_1.append({'conducta': 'Total', 'multa': f"{multa_total_rd_uit:,.3f} UIT"})
        
        # Generar subdoc Tabla 1
        t1_subdoc = create_main_table_coercitiva(
            doc_tpl, 
            ["Conducta Infractora", "Multa Final"], 
            datos_tabla_1, 
            ['conducta', 'multa'], 
            column_widths=(5, 1.5)
        )

        # === TABLA 2: MEDIDA CORRECTIVA (Ítem | Conducta | Medida) ===
        datos_tabla_2 = []
        for i, x in enumerate(hechos_rd_flat_list):
            # Formato Conducta (igual a tabla 1)
            rt_cond = RichText()
            titulo_cond = "Única conducta infractora:" if total_conductas == 1 else f"Conducta infractora n.° {x['num_hecho']}:"
            rt_cond.add(titulo_cond, bold=True, underline=True)
            rt_cond.add(f"\n{x['desc_hecho']}")

            # Formato Medida: Negrita/Subrayado para título + Texto Normal
            rt_med = RichText()
            titulo_med = "Única medida correctiva:" if total_medidas_count == 1 else f"Medida correctiva n.° {x['num_medida']}:"
            rt_med.add(titulo_med, bold=True, underline=True)
            rt_med.add(f"\n{x['desc_medida']}")

            datos_tabla_2.append({
                'item': str(i + 1), # Ítem 1, 2, 3...
                'conducta': rt_cond,
                'medida': rt_med
            })

        # Generar subdoc Tabla 2
        t2_subdoc = create_main_table_coercitiva(
            doc_tpl, 
            ["Ítem", "Conducta Infractora", "Medida Correctiva"], 
            datos_tabla_2, 
            ['item', 'conducta', 'medida'], 
            column_widths=(0.5, 3, 3)
        )

        # === TABLA 3: MULTA COERCITIVA (Ítem | Medida | Multa Coercitiva) ===
        datos_tabla_3 = []
        for x in resultados_medidas:
            datos_tabla_3.append({
                'item': x['num_medida'], # Usamos el ID de la medida (ej. "1", "2")
                'medida': x['desc_medida'],
                'monto': f"{x['multa_coercitiva_final_uit']:,.3f} UIT"
            })
        datos_tabla_3.append({'item': '', 'medida': 'Total', 'monto': f"{multa_coercitiva_total_uit:,.3f} UIT"})

        # Generar subdoc Tabla 3 (Usamos summary table para estilo)
        t3_subdoc = create_summary_table_subdoc(
            doc_tpl, 
            ["Ítem", "Medida Correctiva", "Monto Multa Coercitiva"], 
            datos_tabla_3, 
            ['item', 'medida', 'monto'], 
            column_widths=(0.5, 4, 2)
        )

        # --- 4. CONTEXTO FINAL ---
        ph_gram = _generar_placeholders_gramaticales(num_coer, datos_informe.get('medidas_incumplidas', []), todos_hechos_unicos)
        
        # Datos globales de la primera medida (para uso fuera del bucle)
        d_med_glob = {}
        if resultados_medidas:
            m1 = resultados_medidas[0]
            d_med_glob = {
                'hechos_lista_str_global': m1['hechos_lista_str'], 
                'multa_base_formato_global': m1['multa_base_formato'], 
                'ph_conducta_label_global': m1['ph_conducta_label'], 
                'desc_medida_global': m1['desc_medida'], 
                'obligacion_global': m1['obligacion']
            }

        num_hechos_rd1 = datos_informe.get('num_hechos_rd1', 0)
        
        ctx = {
            **datos_comunes['context_data'], **ph_gram, **d_med_glob,
            'numero_coercitiva_titulo': titulo_coercitiva,
            
            # Datos RD1
            'numero_rd1': datos_informe.get('numero_rd1'),
            'fecha_rd1': format_date(datos_informe.get('fecha_rd1'), "d 'de' MMMM 'de' yyyy", locale='es') if datos_informe.get('fecha_rd1') else '',
            'articulo_declaracion_incumplimiento': datos_informe.get('articulo_resp_rd1', ''),
            'articulo_imposicion_medida': datos_informe.get('articulo_medida_rd1', ''),
            'num_hechos_rd1': num_hechos_rd1,
            'ph_cantidad_infracciones_total': f"{texto_con_numero(num_hechos_rd1, 'f')} ({num_hechos_rd1}) {'infracción' if num_hechos_rd1==1 else 'infracciones'}",
            'multa_total_rd1': f"{datos_informe.get('multa_total_rd1'):,.3f}",
            'num_medidas_rd1': datos_informe.get('num_medidas_rd1'),
            
            # Datos RD2
            'numero_rd2': datos_informe.get('numero_rd2'),
            'fecha_rd2': format_date(datos_informe.get('fecha_rd2'), "d 'de' MMMM 'de' yyyy", locale='es') if datos_informe.get('fecha_rd2') else '',
            'articulo_rd2': datos_informe.get('articulo_rd2', ''),

            # Totales y lógica
            'multa_base_calculada': f"{multa_total_rd_uit:,.3f}", 
            'multa_base_calculada_uit': f"{multa_total_rd_uit:,.3f} UIT",
            'multa_coercitiva_total': f"{multa_coercitiva_total_uit:,.3f}", 
            'multa_coercitiva_total_uit': f"{multa_coercitiva_total_uit:,.3f} UIT",
            'es_metodologia_nueva': (num_coer == 1 and datos_informe.get('metodologia') == 'Nueva'),
            'lista_medidas_calculadas': resultados_medidas,
            
            # --- TABLAS SOLICITADAS ---
            'tabla_multa_impuesta': t1_subdoc,      # Tabla 1
            'tabla_medida_correctiva': t2_subdoc,   # Tabla 2
            'tabla_calculo_coercitiva': t3_subdoc   # Tabla 3
        }

        doc_tpl.render(ctx, autoescape=True); buf = io.BytesIO(); doc_tpl.save(buf)
        return {'doc_pre_compuesto': buf, 'resultados_para_app': {'tabla_resumen_coercitivas': resultados_medidas, 'total_uit': multa_coercitiva_total_uit}}

    except Exception as e:
        import traceback; traceback.print_exc(); return {'error': f"Error fatal Coercitiva: {e}"}