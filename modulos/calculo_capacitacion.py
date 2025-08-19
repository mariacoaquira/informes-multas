import pandas as pd

def calcular_costo_capacitacion(num_personal, df_costos_items, df_indices, fecha_incumplimiento_dt, ipc_incumplimiento, tipo_cambio_incumplimiento):
    """
    Calcula de forma independiente el costo asociado a la capacitación.
    Devuelve un diccionario con los items, sustentos y anexos calculados.
    """
    if num_personal == 1: codigo_item = 'ITEM0110'
    elif 2 <= num_personal <= 5: codigo_item = 'ITEM0111'
    elif 6 <= num_personal <= 10: codigo_item = 'ITEM0112'
    else: codigo_item = 'ITEM0113'
    
    fila_costo_final = df_costos_items[df_costos_items['ID_Item'] == codigo_item].iloc[0]

    # Lógica de cálculo extraída de la función principal
    costo_original = float(fila_costo_final['Costo_Unitario_Item'])
    moneda_original = fila_costo_final['Moneda_Item']

    # Asumimos que la fecha de costeo de la capacitación es fija (ej. '2020-06-30')
    # Este dato podría venir de la fila_costo_final si fuera necesario
    fecha_fuente_dt = pd.to_datetime('2020-06-30')
    
    ipc_costeo_row = df_indices[df_indices['Indice_Mes'].dt.to_period('M') == fecha_fuente_dt.to_period('M')]
    ipc_costeo = ipc_costeo_row.iloc[0]['IPC_Mensual'] if not ipc_costeo_row.empty else 0
    
    tc_costeo_row = df_indices[df_indices['Indice_Mes'].dt.to_period('M') == fecha_fuente_dt.to_period('M')]
    tc_en_fecha_costeo = tc_costeo_row.iloc[0]['TC_Mensual'] if not tc_costeo_row.empty else 0

    precio_base_soles, precio_base_dolares = 0, 0
    if moneda_original == 'US$':
        precio_base_dolares = costo_original
        if tc_en_fecha_costeo > 0:
            precio_base_soles = costo_original * tc_en_fecha_costeo
    else:
        precio_base_soles = costo_original
        if tc_en_fecha_costeo > 0:
            precio_base_dolares = costo_original / tc_en_fecha_costeo

    factor_ajuste = round(ipc_incumplimiento / ipc_costeo, 3) if ipc_costeo > 0 else 0
    precio_base_soles_con_igv = precio_base_soles * 1.18 if fila_costo_final['Incluye_IGV'] == 'NO' else precio_base_soles
    
    monto_soles = precio_base_soles_con_igv * factor_ajuste
    monto_dolares = monto_soles / tipo_cambio_incumplimiento if tipo_cambio_incumplimiento > 0 else 0

    item_calculado = {
        "descripcion": fila_costo_final['Descripcion_Item'],
        "cantidad": 1, "horas": 1,
        "precio_soles": precio_base_soles, "precio_dolares": precio_base_dolares,
        "factor_ajuste": factor_ajuste,
        "monto_soles": monto_soles, "monto_dolares": monto_dolares,
        "costo_original": costo_original, "moneda_original": moneda_original
    }
    
    sustentos = [fila_costo_final.get('Sustento_Item')] if fila_costo_final.get('Sustento_Item') else []
    ids_anexos = [fila_costo_final.get('ID_Anexo_Drive')] if fila_costo_final.get('ID_Anexo_Drive') else []

    return {
        "items_calculados": [item_calculado],
        "sustentos": sustentos,
        "ids_anexos": ids_anexos
    }