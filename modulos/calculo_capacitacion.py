import pandas as pd
from babel.dates import format_date
from funciones import redondeo_excel
from decimal import Decimal, ROUND_HALF_UP

def calcular_costo_capacitacion(num_personal, datos_comunes):
    """
    Calcula el costo de capacitación y ahora devuelve tanto el cálculo final
    como los componentes base para poder realizar el prorrateo.
    """
    try:
        df_costos_items = datos_comunes['df_costos_items']
        df_coti_general = datos_comunes['df_coti_general']
        df_salarios_general = datos_comunes['df_salarios_general']
        df_indices = datos_comunes['df_indices']
        fecha_incumplimiento_dt = datos_comunes['fecha_incumplimiento']
        
        ipc_row_inc = df_indices[df_indices['Indice_Mes'].dt.to_period('M') == pd.to_datetime(fecha_incumplimiento_dt).to_period('M')]
        if ipc_row_inc.empty:
            return {'error': f"No se encontró data de Índices para la fecha {fecha_incumplimiento_dt.strftime('%m-%Y')}"}
        
        ipc_incumplimiento = ipc_row_inc.iloc[0]['IPC_Mensual']
        tipo_cambio_incumplimiento = ipc_row_inc.iloc[0]['TC_Mensual']
        
        if num_personal == 1: codigo_item = 'ITEM0032'
        elif 2 <= num_personal <= 5: codigo_item = 'ITEM0033'
        elif 6 <= num_personal <= 10: codigo_item = 'ITEM0034'
        else: codigo_item = 'ITEM0035'
        
        fila_costo_final = df_costos_items[df_costos_items['ID_Item'] == codigo_item].iloc[0]
        id_general = fila_costo_final['ID_General']
# (Línea 39 aprox.)
        # Inicializar como Decimals
        ipc_costeo, tc_en_fecha_costeo = Decimal('0'), Decimal('0')
        fuente_salario_final, pdf_salario_final, fuente_coti_final = '', '', ''

        if pd.notna(id_general):
            if 'SAL' in id_general:
                fuente_row = df_salarios_general[df_salarios_general['ID_Salario'] == id_general]
                if not fuente_row.empty:
                    anio = int(fuente_row.iloc[0]['Costeo_Salario'])
                    indices_del_anio = df_indices[df_indices['Indice_Mes'].dt.year == anio]
                    if not indices_del_anio.empty:
                        
                        # --- INICIO CORRECCIÓN: Cálculo de media con Decimal ---
                        # 1. Convertir CADA número a Decimal (vía string) antes de promediar
                        ipc_list = [Decimal(str(ipc)) for ipc in indices_del_anio['IPC_Mensual'].dropna()]
                        tc_list = [Decimal(str(tc)) for tc in indices_del_anio['TC_Mensual'].dropna()]
                        
                        if not ipc_list or not tc_list:
                            return {'error': f"No se encontraron datos de IPC/TC para el año {anio}"}

                        # 2. Calcular el promedio usando Decimal (suma / contador)
                        ipc_costeo = sum(ipc_list) / Decimal(len(ipc_list))
                        tc_en_fecha_costeo = sum(tc_list) / Decimal(len(tc_list))
                        # --- FIN CORRECCIÓN ---
                        
                    fuente_salario_final, pdf_salario_final = fuente_row.iloc[0].get('Fuente_Salario', ''), fuente_row.iloc[0].get('PDF_Salario', '')
            
            elif 'COT' in id_general:
                fuente_row = df_coti_general[df_coti_general['ID_Cotizacion'] == id_general]
                if not fuente_row.empty:
                    fecha_fuente_dt = fuente_row.iloc[0]['Fecha_Costeo']
                    if pd.notna(fecha_fuente_dt):
                        ipc_row = df_indices[df_indices['Indice_Mes'].dt.to_period('M') == fecha_fuente_dt.to_period('M')]
                        if not ipc_row.empty:
                            # --- INICIO CORRECCIÓN: Convertir a Decimal vía string ---
                            ipc_costeo = Decimal(str(ipc_row.iloc[0]['IPC_Mensual']))
                            tc_en_fecha_costeo = Decimal(str(ipc_row.iloc[0]['TC_Mensual']))
                            # --- FIN CORRECCIÓN ---
                    fuente_coti_final = fuente_row.iloc[0].get('Fuente_Cotizacion', '')
        
        if ipc_costeo == 0 or tc_en_fecha_costeo == 0:
            return {'error': f"No se encontraron índices para la fecha de costeo del ítem {codigo_item}."}

        # --- INICIO CORRECCIÓN: Usar Decimal para todos los cálculos ---
        
        # 1. Convertir costo original e IGV a Decimal
        costo_original = Decimal(str(fila_costo_final['Costo_Unitario_Item']))
        moneda_original = fila_costo_final['Moneda_Item']
        igv_factor = Decimal('1.18')

        # 2. Calcular precios base (la multiplicación ahora es precisa)
        if moneda_original != 'US$':
            precio_base_soles = costo_original
            precio_base_dolares = costo_original / tc_en_fecha_costeo if tc_en_fecha_costeo != 0 else Decimal('0')
        else:
            # Esta es la multiplicación que causaba el error
            precio_base_soles = costo_original * tc_en_fecha_costeo 
            precio_base_dolares = costo_original

        # 3. Redondear precios base (usando la función de funciones.py)
        precio_base_soles = redondeo_excel(precio_base_soles, 3)
        precio_base_dolares = redondeo_excel(precio_base_dolares, 3)
        
        # 4. Volver a convertir a Decimal para seguir calculando
        precio_base_soles = Decimal(str(precio_base_soles))
        precio_base_dolares = Decimal(str(precio_base_dolares))

        precio_base_soles_con_igv = precio_base_soles * igv_factor if fila_costo_final['Incluye_IGV'] == 'NO' else precio_base_soles
        precio_base_dolares_con_igv = precio_base_dolares * igv_factor if fila_costo_final['Incluye_IGV'] == 'NO' else precio_base_dolares
        
        precio_base_soles_con_igv = redondeo_excel(precio_base_soles_con_igv, 3)
        precio_base_dolares_con_igv = redondeo_excel(precio_base_dolares_con_igv, 3)

        # 5. Volver a convertir a Decimal para el cálculo final
        precio_base_soles_con_igv = Decimal(str(precio_base_soles_con_igv))
        
        # 6. Convertir los factores de incumplimiento a Decimal
        ipc_incumplimiento = Decimal(str(ipc_incumplimiento))
        tipo_cambio_incumplimiento = Decimal(str(tipo_cambio_incumplimiento))
        
        factor_ajuste = ipc_incumplimiento / ipc_costeo if ipc_costeo > 0 else Decimal('0')
        factor_ajuste = redondeo_excel(factor_ajuste, 3) # redondeo_excel devuelve float
        factor_ajuste = Decimal(str(factor_ajuste)) # Convertir de nuevo a Decimal

        # 7. Cálculo final con Decimal
        monto_soles = precio_base_soles_con_igv * factor_ajuste
        monto_dolares = monto_soles / tipo_cambio_incumplimiento if tipo_cambio_incumplimiento > 0 else Decimal('0')
        
        # 8. Redondeo final (usando la función de funciones.py)
        monto_soles = redondeo_excel(monto_soles, 3)
        monto_dolares = redondeo_excel(monto_dolares, 3)
        # --- FIN CORRECCIÓN ---

        item_calculado = {
            "descripcion": fila_costo_final['Descripcion_Item'],
            # 9. Convertir todo de vuelta a float para el diccionario
            "precio_soles": float(precio_base_soles),
            "precio_dolares": float(precio_base_dolares),
            "factor_ajuste": float(factor_ajuste),
            "monto_soles": float(monto_soles),
            "monto_dolares": float(monto_dolares),
        }
        
        ids_anexos = [fila_costo_final.get('ID_Anexo_Drive')] if fila_costo_final.get('ID_Anexo_Drive') else []

        return {
            "items_calculados": [item_calculado],
            "precio_base_soles_con_igv": float(precio_base_soles_con_igv),
            "precio_base_dolares_con_igv": float(precio_base_dolares_con_igv),
            "ipc_costeo": float(ipc_costeo),
            "descripcion": fila_costo_final['Descripcion_Item'],
            "fuente_salario": fuente_salario_final,
            "pdf_salario": pdf_salario_final,
            "fuente_coti": fuente_coti_final,
            "ids_anexos": ids_anexos,
            "fi_mes": format_date(fecha_incumplimiento_dt, "MMMM 'de' yyyy", locale='es'),
            "fi_ipc": float(ipc_incumplimiento),
            "fi_tc": float(tipo_cambio_incumplimiento),
            "precio_dolares": float(precio_base_dolares), # Asegúrate que precio_dolares también sea float
            "error": None
        }

    except Exception as e:
        import traceback
        traceback.print_exc()
        return {'error': f"Error en el módulo de cálculo de capacitación: {e}"}