from datetime import date
from docxtpl import RichText
from num2words import num2words
from funciones import create_table_subdoc, create_main_table_subdoc

def preparar_contexto_especifico(doc_tpl, datos_hecho, datos_generales):
    """
    Prepara y devuelve el diccionario de contexto COMPLETO y específico
    para la infracción tipo INF003.
    """
    print("-> Cargando lógica y contexto específico de INF003...")

    # --- 1. CREACIÓN DE LAS TABLAS ---
    
    # Tabla Costo Evitado (CE)
    tabla_ce_subdoc = None
    ce_data_cruda = datos_hecho.get('ce_data_raw', [])
    if ce_data_cruda:
        ce_table_formatted = []
        for item in ce_data_cruda:
            ce_table_formatted.append({
                'descripcion': item.get('descripcion', ''),
                'precio_usd': f"US$ {item.get('precio_dolares', 0):,.3f}",
                'precio_soles': f"S/ {item.get('precio_soles', 0):,.3f}",
                'factor_ajuste': f"{item.get('factor_ajuste', 0):,.3f}",
                'monto_soles': f"S/ {item.get('monto_soles', 0):,.3f}",
                'monto_usd': f"US$ {item.get('monto_dolares', 0):,.3f}"
            })
        
        total_s = datos_hecho.get('ce_total_soles', 0)
        total_d = datos_hecho.get('ce_total_dolares', 0)
        ce_table_formatted.append({
            'descripcion': 'Total', 'precio_usd': '', 'precio_soles': '', 'factor_ajuste': '',
            'monto_soles': f"S/ {total_s:,.3f}",
            'monto_usd': f"US$ {total_d:,.3f}"
        })

        tabla_ce_subdoc = create_table_subdoc(
            doc_tpl,
            ["Descripción", "Precio asociado (US$)", "Precio asociado (S/)", 
             "Factor de ajuste", "Monto (S/)", "Monto (US$)"],
            ce_table_formatted,
            ['descripcion', 'precio_usd', 'precio_soles', 'factor_ajuste', 
             'monto_soles', 'monto_usd']
        )

    # Tabla Beneficio Ilícito (BI)
    tabla_bi_subdoc = None
    bi_data_cruda = datos_hecho.get('bi_data_raw', [])
    if bi_data_cruda:
        bi_table_formatted = [{'descripcion': item.get('Descripción', ''), 'monto': item.get('Monto', '')} for item in bi_data_cruda]
        tabla_bi_subdoc = create_main_table_subdoc(
            doc_tpl, ["Descripción", "Monto"], bi_table_formatted, ['descripcion', 'monto']
        )

    # Tabla Multa
    tabla_multa_subdoc = None
    multa_data_cruda = datos_hecho.get('multa_data_raw', [])
    if multa_data_cruda:
        multa_table_formatted = [{'componente': item.get('Componentes', ''), 'monto': item.get('Monto', '')} for item in multa_data_cruda]
        tabla_multa_subdoc = create_main_table_subdoc(
            doc_tpl, ["Componentes", "Monto"], multa_table_formatted, ['componente', 'monto']
        )

    # --- 2. PREPARAR DATOS ADICIONALES ---
    
    # Lógica para texto_sustentos
    lista_sustentos = datos_hecho.get('sustentos', [])
    texto_sustentos_rt = RichText()
    for sustento in lista_sustentos:
        texto_sustentos_rt.add(f'- {sustento}\n')

    # Lógica para 'persona_cap'
    num_personas = datos_hecho.get('num_personal_capacitacion', 0)
    texto_personas = "una" if num_personas == 1 else num2words(num_personas, lang='es')
    persona_cap = f"{texto_personas} ({num_personas}) persona{'s' if num_personas != 1 else ''}"

    # Lógica para 'precio_dol'
    precio_dol_texto = "No aplica"
    if ce_data_cruda:
        primer_item = ce_data_cruda[0]
        if primer_item.get('moneda_original') == 'US$':
            costo_orig = primer_item.get('costo_original', 0)
            precio_dol_texto = f"US$ {costo_orig:,.3f}"

    # Lógica para formatear la fecha de incumplimiento
    fecha_inc_obj = datos_hecho.get('fecha_incumplimiento')
    if fecha_inc_obj:
        fecha_incumplimiento_texto = fecha_inc_obj.strftime('%d de %B de %Y').lower()
    else:
        fecha_incumplimiento_texto = "No aplica"

    # --- 3. CONSTRUCCIÓN DEL DICCIONARIO DE CONTEXTO ---
    
    # Primero, creamos el diccionario interno 'hecho'
    datos_para_hecho = {
        'numero_imputado': datos_generales['numero_hecho_actual'],
        'descripcion': RichText(datos_hecho.get('texto_hecho', '')),
        'tabla_ce': tabla_ce_subdoc,
        'tabla_bi': tabla_bi_subdoc,
        'tabla_multa': tabla_multa_subdoc,
        'texto_sustentos': texto_sustentos_rt,
        'persona_cap': persona_cap,
    }

    # Ahora, creamos el contexto final
    contexto_final = {
        **datos_generales['context_data'],
        **datos_hecho,
        'hecho': datos_para_hecho,
        'precio_dol': precio_dol_texto,
        'fecha_incumplimiento_texto': fecha_incumplimiento_texto,
        'mes_hoy': date.today().strftime('%B %Y').lower()
    }
    
    return contexto_final