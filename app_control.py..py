import streamlit as st
import pandas as pd
import io
import json
import gspread

# Configuración de la página
st.set_page_config(page_title="Control Gerencial", page_icon="📊", layout="wide")
st.title("📊 Sistema de Control Gerencial: Apropiado vs Pagado")
st.markdown("Automatización de cruces de información por NIT para el control de pagos y apropiaciones.")

def to_excel_tab1(df):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='CONSOLIDADO NACIONAL')
    workbook = writer.book
    worksheet = writer.sheets['CONSOLIDADO NACIONAL']
    
    header_format = workbook.add_format({'bg_color': '#C00000', 'font_color': 'white', 'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    worksheet.set_row(0, 35) 
    
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)
        
    for row_num in range(1, len(df) + 1):
        excel_row = row_num + 1 
        worksheet.write_formula(row_num, 6, f'=E{excel_row}+F{excel_row}')
        formula_dif = f'=IF(G{excel_row}-J{excel_row}<0, IF(ABS(G{excel_row}-J{excel_row})<=F{excel_row}, 0, G{excel_row}-J{excel_row}), G{excel_row}-J{excel_row})'
        worksheet.write_formula(row_num, 10, formula_dif)
        
    worksheet.set_column('A:A', 15)
    worksheet.set_column('B:B', 35)
    worksheet.set_column('C:C', 16)
    worksheet.set_column('D:D', 20)
    worksheet.set_column('E:J', 18)
    worksheet.set_column('K:K', 15)
    worksheet.set_column('L:L', 20)
    writer.close()
    return output.getvalue()

def to_excel_ss(df):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='SEGURIDAD SOCIAL')
    workbook = writer.book
    worksheet = writer.sheets['SEGURIDAD SOCIAL']
    
    header_format = workbook.add_format({'bg_color': '#C00000', 'font_color': 'white', 'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    worksheet.set_row(0, 35) 
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)
        
    for row_num in range(1, len(df) + 1):
        excel_row = row_num + 1 
        # H: TOTAL SEG SOCIAL APROP = E + F + G
        worksheet.write_formula(row_num, 7, f'=E{excel_row}+F{excel_row}+G{excel_row}')
        # J: DIFERENCIA = H - I
        worksheet.write_formula(row_num, 9, f'=H{excel_row}-I{excel_row}')
        
    worksheet.set_column('A:A', 15)
    worksheet.set_column('B:B', 35)
    worksheet.set_column('C:C', 16)
    worksheet.set_column('D:D', 20)
    worksheet.set_column('E:J', 20)
    worksheet.set_column('K:K', 15)
    writer.close()
    return output.getvalue()

def to_excel_fic(df):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='FIC')
    workbook = writer.book
    worksheet = writer.sheets['FIC']
    
    header_format = workbook.add_format({'bg_color': '#C00000', 'font_color': 'white', 'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    worksheet.set_row(0, 35) 
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)
        
    for row_num in range(1, len(df) + 1):
        excel_row = row_num + 1 
        # H: TOTAL FIC = E + F + G
        worksheet.write_formula(row_num, 7, f'=E{excel_row}+F{excel_row}+G{excel_row}')
        # J: DIFERENCIA = H - I
        worksheet.write_formula(row_num, 9, f'=H{excel_row}-I{excel_row}')
        
    worksheet.set_column('A:A', 15)
    worksheet.set_column('B:B', 35)
    worksheet.set_column('C:C', 16)
    worksheet.set_column('D:D', 20)
    worksheet.set_column('E:J', 20)
    worksheet.set_column('K:K', 15)
    writer.close()
    return output.getvalue()

def get_direct_excel_link(url):
    if not url or url.strip() == "": return None
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

def guardar_en_drive(dataframe, hoja):
    try:
        cred_dict = json.loads(st.secrets["google_credentials"])
        gc = gspread.service_account_from_dict(cred_dict)
        # ID de tu archivo consolidado nacional
        sh = gc.open_by_key("1W0zekg36sXn3z_n4Vpa976yS59lLhLab4T9BwzU3G_s")
        worksheet = sh.worksheet(hoja)
        
        df_clean = safe_fillna(dataframe.copy())
        datos_para_enviar = df_clean.values.tolist()
        worksheet.append_rows(datos_para_enviar)
        return True, f"¡Datos guardados exitosamente en la hoja {hoja} de Google Drive!"
    except Exception as e:
        return False, f"Error al guardar en Drive: {e}"

def extraer_datos_multi(url, col_valor):
    if not url or url.strip() == "": return None
    link = get_direct_excel_link(url)
    df = pd.read_excel(link, sheet_name="EMPRESA")
    df.columns = df.columns.str.strip()
    for c in ['NIT', 'EMPRESA', 'PERIODO', 'CANT EMPLEADOS', col_valor, 'TIPO']:
        if c not in df.columns:
            df[c] = 0 if c not in ['EMPRESA', 'PERIODO', 'TIPO'] else ""
    
    df = df.dropna(subset=['NIT'])
    df = df[df['NIT'].astype(str).str.strip() != ""]
    agg_dict = {'EMPRESA': 'first', 'PERIODO': 'first', 'CANT EMPLEADOS': 'sum', col_valor: 'sum', 'TIPO': 'first'}
    df_grouped = df.groupby('NIT').agg(agg_dict).reset_index()
    return df_grouped.set_index('NIT')

def compilar_cruce_multiple(url_ap1, url_ap2, url_ap3, url_mensual, col_valor, sheet_mensual):
    dict_aprop = {}
    all_nits = set()
    
    if url_ap1.strip():
        df1 = extraer_datos_multi(url_ap1, col_valor)
        if df1 is not None: dict_aprop[1] = df1; all_nits.update(df1.index)
    if url_ap2.strip():
        df2 = extraer_datos_multi(url_ap2, col_valor)
        if df2 is not None: dict_aprop[2] = df2; all_nits.update(df2.index)
    if url_ap3.strip():
        df3 = extraer_datos_multi(url_ap3, col_valor)
        if df3 is not None: dict_aprop[3] = df3; all_nits.update(df3.index)
            
    dict_mensual = {}
    if url_mensual.strip():
        link_m = get_direct_excel_link(url_mensual)
        try:
            df_m = pd.read_excel(link_m, sheet_name=sheet_mensual)
        except:
            df_m = pd.read_excel(link_m, sheet_name="EMPRESA")
            
        df_m.columns = df_m.columns.str.strip()
        if col_valor not in df_m.columns: df_m[col_valor] = 0
        df_m = df_m.dropna(subset=['NIT'])
        df_m = df_m[df_m['NIT'].astype(str).str.strip() != ""]
        
        df_m_grouped = df_m.groupby('NIT').agg({col_valor: 'sum'}).reset_index().set_index('NIT')
        dict_mensual = df_m_grouped
        all_nits.update(df_m_grouped.index)
        
    rows = []
    for nit in list(all_nits):
        empresa, periodo, tipo = "", "", ""
        cant_emps = []
        val_1, val_2, val_3, val_men = 0, 0, 0, 0
        
        for i in range(1, 4):
            if i in dict_aprop and nit in dict_aprop[i].index:
                row = dict_aprop[i].loc[nit]
                if not empresa and row['EMPRESA']: empresa = row['EMPRESA']
                if not periodo and row['PERIODO']: periodo = row['PERIODO']
                if not tipo and row['TIPO']: tipo = row['TIPO']
                emp_val = pd.to_numeric(row['CANT EMPLEADOS'], errors='coerce')
                if not pd.isna(emp_val) and emp_val > 0: cant_emps.append(emp_val)
                
                if i == 1: val_1 = row[col_valor]
                elif i == 2: val_2 = row[col_valor]
                elif i == 3: val_3 = row[col_valor]
        
        if len(dict_mensual) > 0 and nit in dict_mensual.index:
            val_men = dict_mensual.loc[nit][col_valor]
            
        avg_emp = int(round(sum(cant_emps)/len(cant_emps))) if cant_emps else 0
        
        if col_valor == 'SEGURIDAD SOCIAL':
            rows.append({
                'NIT': nit, 'EMPRESA': empresa, 'CANT EMPLEADOS': avg_emp, 'PERIODO': periodo,
                'SEGURIDAD SOCIAL APROPIACION 1': val_1, 'SEGURIDAD SOCIAL APROPIACION 2': val_2, 'SEGURIDAD SOCIAL APROPIACION 3': val_3,
                'TOTAL SEGURIDAD SOCIAL APROPIADO': val_1 + val_2 + val_3, 'VALOR SEGURIDAD SOCIAL': val_men, 'DIFERENCIA': (val_1 + val_2 + val_3) - val_men, 'TIPO': tipo
            })
        else: # FIC
            rows.append({
                'NIT': nit, 'EMPRESA': empresa, 'CANT EMPLEADOS': avg_emp, 'PERIODO': periodo,
                'FIC APROPIACION 1': val_1, 'FIC APROPIACION 2': val_2, 'FIC APROPIACION 3': val_3,
                'TOTAL FIC': val_1 + val_2 + val_3, 'VALOR FIC': val_men, 'DIFERENCIA': (val_1 + val_2 + val_3) - val_men, 'TIPO': tipo
            })
            
    return pd.DataFrame(rows)


st.info("⚠️ Importante: Asegúrate de que los enlaces de Google Drive tengan el permiso configurado como 'Cualquier persona con el enlace puede leer'.")

tab1, tab2, tab3 = st.tabs(["1️⃣ Apropiación vs Pago", "2️⃣ Seguridad Social", "3️⃣ FIC"])

# ==========================================
# GESTIÓN 1: APROPIACIÓN VS PAGO (NO TOCADO, ID DE ARCHIVO ACTUALIZADO)
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
                
                aprop_cols = ['NIT', 'EMPRESA', 'PERIODO', 'CANT EMPLEADOS', 'NOMINA', 'PRESTACIONES', 'TIPO']
                for c in aprop_cols:
                    if c not in df_aprop.columns: df_aprop[c] = 0 if c not in ['EMPRESA', 'PERIODO', 'TIPO'] else ""
                        
                agg_aprop = {'EMPRESA': 'first', 'PERIODO': 'first', 'CANT EMPLEADOS': 'sum', 'NOMINA': 'sum', 'PRESTACIONES': 'sum', 'TIPO': 'first'}
                df_aprop_agrupado = df_aprop.groupby('NIT').agg(agg_aprop).reset_index()
                
                pago_cols = ['NIT', 'EMPRESA', 'CANT EMPLEADOS', 'NOMINA', 'PRESTACIONES', 'TOTAL']
                for c in pago_cols:
                    if c not in df_pago.columns: df_pago[c] = 0 if c != 'EMPRESA' else ""
                        
                agg_pago = {'EMPRESA': 'first', 'CANT EMPLEADOS': 'sum', 'NOMINA': 'sum', 'PRESTACIONES': 'sum', 'TOTAL': 'sum'}
                df_pago_agrupado = df_pago.groupby('NIT').agg(agg_pago).reset_index()
                
                cruce_1 = pd.merge(df_aprop_agrupado, df_pago_agrupado, on="NIT", how="outer", suffixes=('_Apropiado', '_Pagado'))
                cruce_1 = safe_fillna(cruce_1)
                
                cant_promedio = []
                for a, p in zip(cruce_1['CANT EMPLEADOS_Apropiado'], cruce_1['CANT EMPLEADOS_Pagado']):
                    a_val = pd.to_numeric(a, errors='coerce')
                    p_val = pd.to_numeric(p, errors='coerce')
                    a_val = a_val if not pd.isna(a_val) else 0
                    p_val = p_val if not pd.isna(p_val) else 0
                    if a_val > 0 and p_val > 0: cant_promedio.append(int(round((a_val + p_val) / 2)))
                    elif a_val > 0: cant_promedio.append(int(a_val))
                    else: cant_promedio.append(int(p_val))
                        
                df_export = pd.DataFrame()
                df_export['NIT'] = cruce_1['NIT']
                df_export['EMPRESA'] = cruce_1['EMPRESA_Apropiado'].where(cruce_1['EMPRESA_Apropiado'] != "", cruce_1['EMPRESA_Pagado'])
                df_export['CANT EMPLEADOS'] = cant_promedio
                df_export['PERIODO'] = cruce_1.get('PERIODO', "")
                df_export['NOMINA\nAPROPIACION'] = cruce_1['NOMINA_Apropiado']
                df_export['PRESTACIONES\nAPROPIACION'] = cruce_1['PRESTACIONES_Apropiado']
                df_export['TOTAL APROPIACION'] = df_export['NOMINA\nAPROPIACION'] + df_export['PRESTACIONES\nAPROPIACION']
                df_export['NOMINA\nAPROBACION'] = cruce_1['NOMINA_Pagado']
                df_export['PRESTACIONES\nAPROBACION'] = cruce_1['PRESTACIONES_Pagado']
                df_export['TOTAL APROBACION'] = cruce_1['TOTAL']
                dif = df_export['TOTAL APROPIACION'] - df_export['TOTAL APROBACION']
                prestaciones_aprop = df_export['PRESTACIONES\nAPROPIACION']
                df_export['DIFERENCIA'] = dif.where(~((dif < 0) & (abs(dif) <= prestaciones_aprop)), 0)
                df_export['TIPO'] = cruce_1.get('TIPO', "")
                
                st.session_state['df_cruce_1'] = df_export
                st.success("¡Cruce Consolidado Nacional generado con éxito!")
                st.dataframe(df_export.head())
            except Exception as e:
                st.error(f"Error procesando los archivos. Detalle: {e}")
        else:
            st.warning("Por favor, pega ambas URLs para continuar.")
            
    if 'df_cruce_1' in st.session_state:
        st.markdown("### ¿Qué deseas hacer con el resultado?")
        btn_col1, btn_col2 = st.columns(2)
        with btn_col1:
            st.download_button("📥 Descargar CONSOLIDADO NACIONAL (Excel)", data=to_excel_tab1(st.session_state['df_cruce_1']), file_name="CONSOLIDADO NACIONAL.xlsx", mime="application/vnd.ms-excel", use_container_width=True)
        with btn_col2:
            if st.button("☁️ Guardar histórico en Google Drive (NOMINA)", use_container_width=True):
                with st.spinner("Conectando con Google Drive..."):
                    exito, mensaje = guardar_en_drive(st.session_state['df_cruce_1'], "NOMINA")
                    if exito: st.success(mensaje)
                    else: st.error(mensaje)

# ==========================================
# GESTIÓN 2: SEGURIDAD SOCIAL
# ==========================================
with tab2:
    st.header("Cruce: Seguridad Social")
    st.markdown("Pega de 1 a 3 enlaces de apropiación para consolidar.")
    
    col_ap1, col_ap2, col_ap3 = st.columns(3)
    with col_ap1: url_ss_ap1 = st.text_input("URL Apropiación 1:", key="url_ss_ap1")
    with col_ap2: url_ss_ap2 = st.text_input("URL Apropiación 2 (Opcional):", key="url_ss_ap2")
    with col_ap3: url_ss_ap3 = st.text_input("URL Apropiación 3 (Opcional):", key="url_ss_ap3")
        
    url_ss_mensual = st.text_input("URL Reporte Mensual de Seguridad Social:", key="url_ss_mensual")
    
    if st.button("Generar Cruce Seguridad Social", type="primary"):
        urls_validas = [u for u in [url_ss_ap1, url_ss_ap2, url_ss_ap3] if u.strip() != ""]
        if urls_validas and url_ss_mensual:
            try:
                df_export_ss = compilar_cruce_multiple(url_ss_ap1, url_ss_ap2, url_ss_ap3, url_ss_mensual, 'SEGURIDAD SOCIAL', "SSGCON")
                st.session_state['df_cruce_2'] = df_export_ss
                st.success("¡Cruce de Seguridad Social consolidado con éxito!")
                st.dataframe(df_export_ss.head())
            except Exception as e:
                st.error(f"Error procesando. Detalle: {e}")
        else:
            st.warning("Pega al menos 1 URL de apropiación y el reporte mensual.")

    if 'df_cruce_2' in st.session_state:
        st.markdown("### ¿Qué deseas hacer con el resultado?")
        btn_col1, btn_col2 = st.columns(2)
        with btn_col1:
            st.download_button("📥 Descargar Reporte SEGURIDAD SOCIAL (Excel)", data=to_excel_ss(st.session_state['df_cruce_2']), file_name="CRUCE_SEGURIDAD_SOCIAL.xlsx", mime="application/vnd.ms-excel", use_container_width=True)
        with btn_col2:
            if st.button("☁️ Guardar histórico en Google Drive (SEGURIDAD SOCIAL)", use_container_width=True, key="save_ss"):
                with st.spinner("Conectando con Google Drive..."):
                    exito, mensaje = guardar_en_drive(st.session_state['df_cruce_2'], "SEGURIDAD SOCIAL")
                    if exito: st.success(mensaje)
                    else: st.error(mensaje)

# ==========================================
# GESTIÓN 3: FIC
# ==========================================
with tab3:
    st.header("Cruce: FIC")
    st.markdown("Pega de 1 a 3 enlaces de apropiación para consolidar.")
    
    col_fic1, col_fic2, col_fic3 = st.columns(3)
    with col_fic1: url_fic_ap1 = st.text_input("URL Apropiación 1:", key="url_fic_ap1")
    with col_fic2: url_fic_ap2 = st.text_input("URL Apropiación 2 (Opcional):", key="url_fic_ap2")
    with col_fic3: url_fic_ap3 = st.text_input("URL Apropiación 3 (Opcional):", key="url_fic_ap3")
        
    url_fic_mensual = st.text_input("URL Reporte Mensual de FIC:", key="url_fic_mensual")
    
    if st.button("Generar Cruce FIC", type="primary"):
        urls_validas = [u for u in [url_fic_ap1, url_fic_ap2, url_fic_ap3] if u.strip() != ""]
        if urls_validas and url_fic_mensual:
            try:
                df_export_fic = compilar_cruce_multiple(url_fic_ap1, url_fic_ap2, url_fic_ap3, url_fic_mensual, 'FIC', "EMPRESA")
                st.session_state['df_cruce_3'] = df_export_fic
                st.success("¡Cruce de FIC consolidado con éxito!")
                st.dataframe(df_export_fic.head())
            except Exception as e:
                st.error(f"Error procesando. Detalle: {e}")
        else:
            st.warning("Pega al menos 1 URL de apropiación y el reporte mensual.")
            
    if 'df_cruce_3' in st.session_state:
        st.markdown("### ¿Qué deseas hacer con el resultado?")
        btn_col1, btn_col2 = st.columns(2)
        with btn_col1:
            st.download_button("📥 Descargar Reporte FIC (Excel)", data=to_excel_fic(st.session_state['df_cruce_3']), file_name="CRUCE_FIC.xlsx", mime="application/vnd.ms-excel", use_container_width=True)
        with btn_col2:
            if st.button("☁️ Guardar histórico en Google Drive (FIC)", use_container_width=True, key="save_fic"):
                with st.spinner("Conectando con Google Drive..."):
                    exito, mensaje = guardar_en_drive(st.session_state['df_cruce_3'], "FIC")
                    if exito: st.success(mensaje)
                    else: st.error(mensaje)
