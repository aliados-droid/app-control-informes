import streamlit as st
import pandas as pd
import io

# Configuración de la página
st.set_page_config(page_title="Control Gerencial", page_icon="📊", layout="wide")
st.title("📊 Sistema de Control Gerencial: Apropiado vs Pagado")
st.markdown("Automatización de cruces de información por NIT para el control de pagos y apropiaciones.")

# Función original para Seguridad Social y FIC
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Cruce Consolidado')
    return output.getvalue()

# NUEVA FUNCIÓN MÁGICA: Genera Excel con FÓRMULAS y ENCABEZADOS ROJOS para la Gestión 1
def to_excel_tab1(df):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Cruce Consolidado')
    
    workbook = writer.book
    worksheet = writer.sheets['Cruce Consolidado']
    
    # 1. Crear el formato de encabezado rojo con letra blanca
    header_format = workbook.add_format({
        'bg_color': '#C00000', # Color rojo oscuro idéntico a tu imagen
        'font_color': 'white',
        'bold': True,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })
    
    # 2. Crear formato para el porcentaje (sin decimales, ej. 98%)
    pct_format = workbook.add_format({'num_format': '0%', 'align': 'center'})
    
    # 3. Aplicar el fondo rojo a todos los encabezados
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)
        
    # 4. Inyectar las fórmulas reales de Excel fila por fila
    for row_num in range(1, len(df) + 1):
        excel_row = row_num + 1 # Fila en Excel (1-based + 1 por el encabezado)
        
        # Fórmula DIFERENCIA (Col H) = NOMINA (Col D) - EMPRESA_Pagado (Col G)
        worksheet.write_formula(row_num, 7, f'=D{excel_row}-G{excel_row}')
        
        # Fórmula PORCENTAJE (Col I) = EMPRESA_Pagado (Col G) / NOMINA (Col D)
        worksheet.write_formula(row_num, 8, f'=IFERROR(G{excel_row}/D{excel_row}, 0)', pct_format)
        
    # Ajustar el ancho de las columnas para que se vea ordenado
    worksheet.set_column('A:A', 15)
    worksheet.set_column('B:B', 35)
    worksheet.set_column('C:C', 18)
    worksheet.set_column('D:G', 16)
    worksheet.set_column('H:I', 15)
    
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
# GESTIÓN 1: APROPIACIÓN VS PAGO
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
                hoja_pago = "NOMCONBOG" if "Bogotá" in ciudad else "NOMCONEJE"
                df_pago = pd.read_excel(link_pago, sheet_name=hoja_pago)
                
                df_aprop.columns = df_aprop.columns.str.strip()
                df_pago.columns = df_pago.columns.str.strip()
                
                # AGRUPAR APROPIACIÓN
                aprop_num_cols = df_aprop.select_dtypes(include='number').columns.tolist()
                if 'NIT' in aprop_num_cols: aprop_num_cols.remove('NIT')
                agg_aprop = {col: 'sum' for col in aprop_num_cols}
                if 'EMPRESA' in df_aprop.columns: agg_aprop['EMPRESA'] = 'first'
                if 'TOTAL ACTA' in df_aprop.columns and 'TOTAL ACTA' not in agg_aprop: agg_aprop['TOTAL ACTA'] = 'sum'
                df_aprop_agrupado = df_aprop.groupby('NIT').agg(agg_aprop).reset_index()
                
                # AGRUPAR PAGO
                pago_num_cols = df_pago.select_dtypes(include='number').columns.tolist()
                if 'NIT' in pago_num_cols: pago_num_cols.remove('NIT')
                agg_pago = {col: 'sum' for col in pago_num_cols}
                if 'EMPRESA' in df_pago.columns: agg_pago['EMPRESA'] = 'first'
                if 'TOTAL' in df_pago.columns and 'TOTAL' not in agg_pago: agg_pago['TOTAL'] = 'sum'
                df_pago_agrupado = df_pago.groupby('NIT').agg(agg_pago).reset_index()
                
                # CRUCE
                cruce_1 = pd.merge(df_aprop_agrupado, df_pago_agrupado, on="NIT", how="outer", suffixes=('_Apropiado', '_Pagado'))
                cruce_1 = safe_fillna(cruce_1)
                
                # CONSTRUIR LA ESTRUCTURA EXACTA QUE PEDISTE
                df_export = pd.DataFrame()
                df_export['NIT'] = cruce_1['NIT']
                df_export['EMPRESA_Apropiado'] = cruce_1['EMPRESA_Apropiado']
                
                col_cant = 'CANT EMPLEADOS_Apropiado' if 'CANT EMPLEADOS_Apropiado' in cruce_1.columns else 'CANT EMPLEADOS'
                df_export['CANT EMPLEADOS'] = cruce_1.get(col_cant, 0)
                
                df_export['NOMINA'] = cruce_1.get('NOMINA_Apropiado', 0)
                df_export['PRESTACIONES'] = cruce_1.get('PRESTACIONES_Apropiado', 0)
                
                col_total_aprop = 'TOTAL ACTA' if 'TOTAL ACTA' in cruce_1.columns else 'TOTAL_Apropiado'
                df_export['TOTAL'] = cruce_1.get(col_total_aprop, 0)
                
                # Se utiliza NOMINA_Pagado bajo el nombre de columna EMPRESA_Pagado como en tu imagen
                df_export['EMPRESA_Pagado'] = cruce_1.get('NOMINA_Pagado', 0)
                
                # Cálculos temporales solo para que se vean en la pantalla de la web
                df_export['DIFERENCIA'] = df_export['NOMINA'] - df_export['EMPRESA_Pagado']
                nomina_segura = df_export['NOMINA'].replace(0, 1) # Evitar dividir por cero en la web
                df_export['PORCENTAJE'] = (df_export['EMPRESA_Pagado'] / nomina_segura).where(df_export['NOMINA'] != 0, 0)
                
                st.success("¡Reporte Gerencial Listo! El archivo Excel incluye encabezados rojos y fórmulas reales.")
                st.dataframe(df_export.head())
                
                # Descarga utilizando la nueva función con formato y fórmulas
                st.download_button(
                    label="📥 Descargar Reporte en Excel", 
                    data=to_excel_tab1(df_export), 
                    file_name="Reporte_Gerencial_Apropiacion_vs_Pago.xlsx", 
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
