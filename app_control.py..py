import streamlit as st
import pandas as pd
import io

# Configuración de la página
st.set_page_config(page_title="Control Gerencial", page_icon="📊", layout="wide")
st.title("📊 Sistema de Control Gerencial: Apropiado vs Pagado")
st.markdown("Automatización de cruces de información por NIT para el control de pagos y apropiaciones.")

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Cruce Consolidado')
    return output.getvalue()

# FUNCIÓN MÁGICA PARA EL CONSOLIDADO NACIONAL (PESTAÑA 1)
def to_excel_tab1(df):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='CONSOLIDADO NACIONAL')
    
    workbook = writer.book
    worksheet = writer.sheets['CONSOLIDADO NACIONAL']
    
    # 1. Crear el formato de encabezado rojo con letra blanca y salto de línea
    header_format = workbook.add_format({
        'bg_color': '#C00000', # Color rojo oscuro
        'font_color': 'white',
        'bold': True,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True # Permite que el texto se divida en dos líneas
    })
    
    worksheet.set_row(0, 35) # Hace más alto el encabezado para que quepa el texto
    
    # 2. Aplicar el formato
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)
        
    # 3. Inyectar las fórmulas reales fila por fila
    for row_num in range(1, len(df) + 1):
        excel_row = row_num + 1 
        
        # F: TOTAL APROPIACION = D + E (NOMINA + PRESTACIONES APROP)
        worksheet.write_formula(row_num, 5, f'=D{excel_row}+E{excel_row}')
        
        # J: DIFERENCIA = F - I (con condicional para validar contra Prestaciones Aprop "Columna E")
        formula_dif = f'=IF(F{excel_row}-I{excel_row}<0, IF(ABS(F{excel_row}-I{excel_row})<=E{excel_row}, 0, F{excel_row}-I{excel_row}), F{excel_row}-I{excel_row})'
        worksheet.write_formula(row_num, 9, formula_dif)
        
    # 4. Ajustar el ancho de las columnas
    worksheet.set_column('A:A', 15)
    worksheet.set_column('B:B', 35)
    worksheet.set_column('C:C', 18)
    worksheet.set_column('D:I', 18)
    worksheet.set_column('J:J', 15)
    
    writer.close()
    return output.getvalue()

def get_direct_excel_link(url):
    if not url or url.strip() == "":
        return None
    if "drive.google.com/file/d/" in url:
        file_id = url.split("/d/")[1].split("/")[0]
        return f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
    elif "docs.google.com/spreadsheets/d/" in url:
        file_id = url.split("/d/")[1].split("/")[0]
        return f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
    return url

def safe_fillna(df):
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            df[col] = df[col].fillna(0)
        else:
            df[col] = df[col].fillna("")
    return df

st.info("⚠️ Importante: Asegúrate de que los enlaces de Google Drive tengan el permiso configurado como 'Cualquier persona con el enlace puede leer'.")

tab1, tab2, tab3 = st.tabs(["1️⃣ Apropiación vs Pago", "2️⃣ Seguridad Social", "3️⃣ FIC"])

# ==========================================
# GESTIÓN 1: APROPIACIÓN VS PAGO (ACTUALIZADA)
# ==========================================
with tab1:
    st.header("Cruce: Apropiación vs Pago")
    col1, col2 = st.columns(2)
    
    with col1:
        url_aprop_1 = st.text_input("Pega la URL de Drive - Apropiación:", key="url_aprop1")
        ciudad = st.selectbox("Selecciona la región del pago:", ["Bogotá (NOMCONBOG)", "Eje Cafetero (NOMCONEJE)"])
    with col2:
        url_pago_1 = st.text_input("Pega la URL de Drive - Aprobación de Pago:", key="url_pago1")
    
    if st.button("Generar Cruce Apropiación vs Pago", type="primary"):
        if url_aprop_1 and url_pago_1:
            try:
                link_aprop = get_direct_excel_link(url_aprop_1)
                link_pago = get_direct_excel_link(url_pago_1)
                
                df_aprop = pd.read_excel(link_aprop, sheet_name="EMPRESA")
                # Usa NOMCONBOG o NOMCONEJE dependiendo del selectbox
                hoja_pago = "NOMCONBOG" if "Bogotá" in ciudad else "NOMCONEJE"
                df_pago = pd.read_excel(link_pago, sheet_name=hoja_pago)
                
                df_aprop.columns = df_aprop.columns.str.strip()
                df_pago.columns = df_pago.columns.str.strip()
                
                # 1. FILTRAR Y AGRUPAR APROPIACIÓN
                aprop_cols = ['NIT', 'EMPRESA', 'CANT EMPLEADOS', 'NOMINA', 'PRESTACIONES']
                for c in aprop_cols:
                    if c not in df_aprop.columns:
                        df_aprop[c] = 0 if c != 'EMPRESA' else ""
                        
                agg_aprop = {'EMPRESA': 'first', 'CANT EMPLEADOS': 'sum', 'NOMINA': 'sum', 'PRESTACIONES': 'sum'}
                df_aprop_agrupado = df_aprop.groupby('NIT').agg(agg_aprop).reset_index()
                
                # 2. FILTRAR Y AGRUPAR PAGO
                pago_cols = ['NIT', 'EMPRESA', 'CANT EMPLEADOS', 'NOMINA', 'PRESTACIONES', 'TOTAL']
                for c in pago_cols:
                    if c not in df_pago.columns:
                        df_pago[c] = 0 if c != 'EMPRESA' else ""
                        
                agg_pago = {'EMPRESA': 'first', 'CANT EMPLEADOS': 'sum', 'NOMINA': 'sum', 'PRESTACIONES': 'sum', 'TOTAL': 'sum'}
                df_pago_agrupado = df_pago.groupby('NIT').agg(agg_pago).reset_index()
                
                # 3. CRUCE CONSOLIDADO
                cruce_1 = pd.merge(df_aprop_agrupado, df_pago_agrupado, on="NIT", how="outer", suffixes=('_Apropiado', '_Pagado'))
                cruce_1 = safe_fillna(cruce_1)
                
                # 4. PROMEDIO DE EMPLEADOS
                cant_promedio = []
                for a, p in zip(cruce_1['CANT EMPLEADOS_Apropiado'], cruce_1['CANT EMPLEADOS_Pagado']):
                    a_val = pd.to_numeric(a, errors='coerce')
                    p_val = pd.to_numeric(p, errors='coerce')
                    a_val = a_val if not pd.isna(a_val) else 0
                    p_val = p_val if not pd.isna(p_val) else 0
                    
                    if a_val > 0 and p_val > 0:
                        cant_promedio.append(int(round((a_val + p_val) / 2)))
                    elif a_val > 0:
                        cant_promedio.append(int(a_val))
                    else:
                        cant_promedio.append(int(p_val))
                        
                # 5. CONSTRUIR LA ESTRUCTURA DEL CONSOLIDADO NACIONAL
                df_export = pd.DataFrame()
                df_export['NIT'] = cruce_1['NIT']
                
                # Nombre de la empresa de cualquiera de los lados
                df_export['EMPRESA'] = cruce_1['EMPRESA_Apropiado'].where(cruce_1['EMPRESA_Apropiado'] != "", cruce_1['EMPRESA_Pagado'])
                
                df_export['CANT EMPLEADOS'] = cant_promedio
                df_export['NOMINA\nAPROPIACION'] = cruce_1['NOMINA_Apropiado']
                df_export['PRESTACIONES\nAPROPIACION'] = cruce_1['PRESTACIONES_Apropiado']
                
                # Cálculo para la vista previa de la miniapp (en Excel irá la fórmula real)
                df_export['TOTAL APROPIACION'] = df_export['NOMINA\nAPROPIACION'] + df_export['PRESTACIONES\nAPROPIACION']
                
                df_export['NOMINA\nAPROBACION'] = cruce_1['NOMINA_Pagado']
                df_export['PRESTACIONES\nAPROBACION'] = cruce_1['PRESTACIONES_Pagado']
                df_export['TOTAL APROBACION'] = cruce_1['TOTAL']
                
                # Cálculo temporal para la vista previa de DIFERENCIA en la miniapp
                dif = df_export['TOTAL APROPIACION'] - df_export['TOTAL APROBACION']
                prestaciones_aprop = df_export['PRESTACIONES\nAPROPIACION']
                df_export['DIFERENCIA'] = dif.where(~((dif < 0) & (abs(dif) <= prestaciones_aprop)), 0)
                
                st.success("¡Cruce Consolidado Nacional generado con éxito!")
                st.dataframe(df_export.head())
                
                # Botón de Descarga
                st.download_button(
                    label="📥 Descargar CONSOLIDADO NACIONAL", 
                    data=to_excel_tab1(df_export), 
                    file_name="CONSOLIDADO NACIONAL.xlsx", 
                    mime="application/vnd.ms-excel"
                )
            except Exception as e:
                st.error(f"Error procesando los archivos. Detalle: {e}")
        else:
            st.warning("Por favor, pega ambas URLs para continuar.")

# ==========================================
# GESTIÓN 2: SEGURIDAD SOCIAL
# ==========================================
with tab2:
    st.header("Cruce: Seguridad Social")
    st.markdown("Pega de 1 a 3 enlaces de apropiación.")
    
    col_ap1, col_ap2, col_ap3 = st.columns(3)
    with col_ap1: url_ss_ap1 = st.text_input("URL Apropiación 1:", key="url_ss_ap1")
    with col_ap2: url_ss_ap2 = st.text_input("URL Apropiación 2 (Opcional):", key="url_ss_ap2")
    with col_ap3: url_ss_ap3 = st.text_input("URL Apropiación 3 (Opcional):", key="url_ss_ap3")
        
    url_ss_mensual = st.text_input("URL Reporte Mensual de Seguridad Social:", key="url_ss_mensual")
    
    if st.button("Generar Cruce Seguridad Social", type="primary"):
        urls_validas = [u for u in [url_ss_ap1, url_ss_ap2, url_ss_ap3] if u.strip() != ""]
        
        if urls_validas and url_ss_mensual:
            try:
                lista_df_aprop = []
                for url in urls_validas:
                    link_directo = get_direct_excel_link(url)
                    df_temp = pd.read_excel(link_directo, sheet_name="EMPRESA")
                    lista_df_aprop.append(df_temp[["NIT", "SEGURIDAD SOCIAL"]])
                
                df_aprop_consol = pd.concat(lista_df_aprop)
                df_aprop_agrupado = df_aprop_consol.groupby("NIT")["SEGURIDAD SOCIAL"].sum().reset_index()
                df_aprop_agrupado.rename(columns={"SEGURIDAD SOCIAL": "SEGURIDAD_SOCIAL_APROPIADA"}, inplace=True)
                
                link_mensual = get_direct_excel_link(url_ss_mensual)
                df_ss_mensual = pd.read_excel(link_mensual, sheet_name="SSGCON")
                
                cruce_2 = pd.merge(df_aprop_agrupado, df_ss_mensual, on="NIT", how="outer")
                cruce_2 = safe_fillna(cruce_2)
                
                st.success("¡Cruce de Seguridad Social consolidado!")
                st.dataframe(cruce_2.head())
                st.download_button(label="📥 Descargar Reporte Seguridad Social", data=to_excel(cruce_2), file_name="Cruce_Seguridad_Social.xlsx")
            except Exception as e:
                st.error(f"Error procesando. Detalle: {e}")
        else:
            st.warning("Pega al menos 1 URL de apropiación y el reporte mensual.")

# ==========================================
# GESTIÓN 3: FIC
# ==========================================
with tab3:
    st.header("Cruce: FIC")
    st.markdown("Pega de 1 a 3 enlaces de apropiación.")
    
    col_fic1, col_fic2, col_fic3 = st.columns(3)
    with col_fic1: url_fic_ap1 = st.text_input("URL Apropiación 1:", key="url_fic_ap1")
    with col_fic2: url_fic_ap2 = st.text_input("URL Apropiación 2 (Opcional):", key="url_fic_ap2")
    with col_fic3: url_fic_ap3 = st.text_input("URL Apropiación 3 (Opcional):", key="url_fic_ap3")
        
    url_fic_mensual = st.text_input("URL Reporte Mensual de FIC:", key="url_fic_mensual")
    
    if st.button("Generar Cruce FIC", type="primary"):
        urls_validas = [u for u in [url_fic_ap1, url_fic_ap2, url_fic_ap3] if u.strip() != ""]
        
        if urls_validas and url_fic_mensual:
            try:
                lista_df_fic = []
                for url in urls_validas:
                    link_directo = get_direct_excel_link(url)
                    df_temp = pd.read_excel(link_directo, sheet_name="EMPRESA")
                    lista_df_fic.append(df_temp[["NIT", "FIC"]])
                
                df_fic_consol = pd.concat(lista_df_fic)
                df_fic_agrupado = df_fic_consol.groupby("NIT")["FIC"].sum().reset_index()
                df_fic_agrupado.rename(columns={"FIC": "FIC_APROPIADO"}, inplace=True)
                
                link_mensual = get_direct_excel_link(url_fic_mensual)
                df_fic_mensual = pd.read_excel(link_mensual, sheet_name="EMPRESA")
                
                cruce_3 = pd.merge(df_fic_agrupado, df_fic_mensual, on="NIT", how="outer")
                cruce_3 = safe_fillna(cruce_3)
                
                st.success("¡Cruce de FIC consolidado!")
                st.dataframe(cruce_3.head())
                st.download_button(label="📥 Descargar Reporte FIC", data=to_excel(cruce_3), file_name="Cruce_FIC.xlsx")
            except Exception as e:
                st.error(f"Error procesando. Detalle: {e}")
        else:
            st.warning("Pega al menos 1 URL de apropiación y el reporte mensual.")
