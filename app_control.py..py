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
    
    header_format = workbook.add_format({
        'bg_color': '#C00000', 'font_color': 'white', 'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True
    })
    worksheet.set_row(0, 35) 
    
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)
        
    for row_num in range(1, len(df) + 1):
        excel_row = row_num + 1 
        worksheet.write_formula(row_num, 6, f'=E{excel_row}+F{excel_row}')
        # NUEVA FÓRMULA DIRECTA: NOMINA APROPIACIÓN - NOMINA APROBACIÓN
        formula_dif = f'=E{excel_row}-H{excel_row}'
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
        worksheet.write_formula(row_num, 7, f'=E{excel_row}+F{excel_row}+G{excel_row}')
        worksheet.write_formula(row_num, 9, f'=H{excel_row}-I{excel_row}')
        
    worksheet.set_column('A:A', 15) 
    worksheet.set_column('B:B', 35) 
    worksheet.set_column('C:C', 16) 
    worksheet.set_column('D:D', 20) 
    worksheet.set_column('E:J', 18) 
    worksheet.set_column('K:K', 20) 
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
        worksheet.write_formula(row_num, 7, f'=E{excel_row}+F{excel_row}+G{excel_row}')
        worksheet.write_formula(row_num, 9, f'=H{excel_row}-I{excel_row}')
        
    worksheet.set_column('A:A', 15) 
    worksheet.set_column('B:B', 35) 
    worksheet.set_column('C:C', 16) 
    worksheet.set_column('D:D', 20) 
    worksheet.set_column('E:J', 18) 
    worksheet.set_column('K:K', 20) 
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

def coalesce_strings(row, cols):
    for c in cols:
        val = row.get(c)
        if pd.notna(val) and str(val).strip() != "":
            return str(val).strip()
    return ""

def calc_promedio_empleados(row, cols):
    vals = []
    for c in cols:
        val = row.get(c)
        if pd.notna(val):
            try:
                v = float(val)
                if v > 0: vals.append(v)
            except: pass
    if len(vals) > 0:
        return int(round(sum(vals) / len(vals)))
    return 0

# FUNCIÓN INTELIGENTE PARA GUARDAR EN GOOGLE SHEETS
def guardar_en_drive(dataframe, sheet_name):
    try:
        cred_dict = json.loads(st.secrets["google_credentials"])
        gc = gspread.service_account_from_dict(cred_dict)
        sh = gc.open_by_key("1W0zekg36sXn3z_n4Vpa976yS59lLhLab4T9BwzU3G_s")
        worksheet = sh.worksheet(sheet_name)
        
        df_clean = safe_fillna(dataframe.copy())
        
        col_a = worksheet.col_values(1)
        registros_reales = [x for x in col_a if str(x).strip() != ""]
        next_row = len(registros_reales) + 1
        
        datos_para_enviar = df_clean.values.tolist()
        
        for i, row in enumerate(datos_para_enviar):
            current_row = next_row + i
            if sheet_name == "NOMINA":
                row[6] = f'=E{current_row}+F{current_row}'
                # NUEVA FÓRMULA DIRECTA PARA GOOGLE DRIVE
                row[10] = f'=E{current_row}-H{current_row}'
            elif sheet_name in ["SEGURIDAD SOCIAL", "FIC"]:
                row[7] = f'=E{current_row}+F{current_row}+G{current_row}'
                row[9] = f'=H{current_row}-I{current_row}'
        
        try:
            worksheet.update(f"A{next_row}", datos_para_enviar, value_input_option='USER_ENTERED')
        except TypeError:
            worksheet.update(values=datos_para_enviar, range_name=f"A{next_row}", value_input_option='USER_ENTERED')
            
        return True, f"¡Datos y fórmulas guardados exitosamente en Google Drive ({sheet_name})!"
    except Exception as e:
        return False, f"Error al guardar en Drive: {e}"

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
                
                aprop_cols = ['NIT', 'EMPRESA', 'PERIODO', 'CANT EMPLEADOS', 'NOMINA', 'PRESTACIONES', 'TIPO']
                for c in aprop_cols:
                    if c not in df_aprop.columns: df_aprop[c] = "" if c in ['EMPRESA', 'PERIODO', 'TIPO'] else 0
                
                df_aprop['NIT'] = df_aprop['NIT'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                for col in ['CANT EMPLEADOS', 'NOMINA', 'PRESTACIONES']:
                    df_aprop[col] = pd.to_numeric(df_aprop[col], errors='coerce').fillna(0)
                        
                agg_aprop = {'EMPRESA': 'first', 'PERIODO': 'first', 'CANT EMPLEADOS': 'sum', 'NOMINA': 'sum', 'PRESTACIONES': 'sum', 'TIPO': 'first'}
                df_aprop_agrupado = df_aprop.groupby('NIT').agg(agg_aprop).reset_index()
                
                pago_cols = ['NIT', 'EMPRESA', 'CANT EMPLEADOS', 'NOMINA', 'PRESTACIONES', 'TOTAL']
                for c in pago_cols:
                    if c not in df_pago.columns: df_pago[c] = "" if c == 'EMPRESA' else 0
                
                df_pago['NIT'] = df_pago['NIT'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                for col in ['CANT EMPLEADOS', 'NOMINA', 'PRESTACIONES', 'TOTAL']:
                    df_pago[col] = pd.to_numeric(df_pago[col], errors='coerce').fillna(0)
                        
                agg_pago = {'EMPRESA': 'first', 'CANT EMPLEADOS': 'sum', 'NOMINA': 'sum', 'PRESTACIONES': 'sum', 'TOTAL': 'sum'}
                df_pago_agrupado = df_pago.groupby('NIT').agg(agg_pago).reset_index()
                
                cruce_1 = pd.merge(df_aprop_agrupado, df_pago_agrupado, on="NIT", how="outer", suffixes=('_Apropiado', '_Pagado'))
                cruce_1 = safe_fillna(cruce_1)
                
                emp_cols = [c for c in cruce_1.columns if 'EMPRESA' in c]
                per_cols = [c for c in cruce_1.columns if 'PERIODO' in c]
                tip_cols = [c for c in cruce_1.columns if 'TIPO' in c]
                
                cant_aprop_col = 'CANT EMPLEADOS_Apropiado' if 'CANT EMPLEADOS_Apropiado' in cruce_1 else 'CANT EMPLEADOS'
                cant_pago_col = 'CANT EMPLEADOS_Pagado' if 'CANT EMPLEADOS_Pagado' in cruce_1 else 'CANT EMPLEADOS'
                
                cant_promedio = []
                for a, p in zip(cruce_1.get(cant_aprop_col, [0]*len(cruce_1)), cruce_1.get(cant_pago_col, [0]*len(cruce_1))):
                    a_val = pd.to_numeric(a, errors='coerce'); a_val = a_val if not pd.isna(a_val) else 0
                    p_val = pd.to_numeric(p, errors='coerce'); p_val = p_val if not pd.isna(p_val) else 0
                    if a_val > 0 and p_val > 0: cant_promedio.append(int(round((a_val + p_val) / 2)))
                    elif a_val > 0: cant_promedio.append(int(a_val))
                    else: cant_promedio.append(int(p_val))
                        
                df_export = pd.DataFrame()
                df_export['NIT'] = cruce_1['NIT']
                df_export['EMPRESA'] = cruce_1.apply(lambda r: coalesce_strings(r, emp_cols), axis=1)
                df_export['CANT EMPLEADOS'] = cant_promedio
                df_export['PERIODO'] = cruce_1.apply(lambda r: coalesce_strings(r, per_cols), axis=1)
                
                df_export['NOMINA\nAPROPIACION'] = pd.to_numeric(cruce_1.get('NOMINA_Apropiado', 0), errors='coerce').fillna(0)
                df_export['PRESTACIONES\nAPROPIACION'] = pd.to_numeric(cruce_1.get('PRESTACIONES_Apropiado', 0), errors='coerce').fillna(0)
                df_export['TOTAL APROPIACION'] = df_export['NOMINA\nAPROPIACION'] + df_export['PRESTACIONES\nAPROPIACION']
                
                df_export['NOMINA\nAPROBACION'] = pd.to_numeric(cruce_1.get('NOMINA_Pagado', 0), errors='coerce').fillna(0)
                df_export['PRESTACIONES\nAPROBACION'] = pd.to_numeric(cruce_1.get('PRESTACIONES_Pagado', 0), errors='coerce').fillna(0)
                df_export['TOTAL APROBACION'] = pd.to_numeric(cruce_1.get('TOTAL', 0), errors='coerce').fillna(0)
                
                # NUEVO CÁLCULO DIRECTO PARA LA VISTA EN PANTALLA
                df_export['DIFERENCIA'] = df_export['NOMINA\nAPROPIACION'] - df_export['NOMINA\nAPROBACION']
                
                df_export['TIPO'] = cruce_1.apply(lambda r: coalesce_strings(r, tip_cols), axis=1)
                
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
            st.download_button(label="📥 Descargar CONSOLIDADO NACIONAL (Excel)", data=to_excel_tab1(st.session_state['df_cruce_1']), file_name="CONSOLIDADO NACIONAL.xlsx", mime="application/vnd.ms-excel", use_container_width=True)
        with btn_col2:
            if st.button("☁️ Guardar en hoja NOMINA", use_container_width=True):
                with st.spinner("Conectando con Google Drive..."):
                    exito, mensaje = guardar_en_drive(st.session_state['df_cruce_1'], "NOMINA")
                    if exito: st.success(mensaje)
                    else: st.error(mensaje)

# ==========================================
# GESTIÓN 2: SEGURIDAD SOCIAL
# ==========================================
with tab2:
    st.header("Cruce: Seguridad Social")
    st.markdown("Pega de 1 a 3 enlaces de apropiación y tu enlace de mensualidad.")
    
    col_ap1, col_ap2, col_ap3 = st.columns(3)
    with col_ap1: url_ss_ap1 = st.text_input("URL Apropiación 1:", key="url_ss_ap1")
    with col_ap2: url_ss_ap2 = st.text_input("URL Apropiación 2 (Opcional):", key="url_ss_ap2")
    with col_ap3: url_ss_ap3 = st.text_input("URL Apropiación 3 (Opcional):", key="url_ss_ap3")
        
    url_ss_mensual = st.text_input("URL Reporte Mensual de Seguridad Social:", key="url_ss_mensual")
    
    if st.button("Generar Cruce Seguridad Social", type="primary"):
        urls_validas = [u for u in [url_ss_ap1, url_ss_ap2, url_ss_ap3] if u.strip() != ""]
        if urls_validas and url_ss_mensual:
            try:
                dfs_aprop = []
                for i, url in enumerate(urls_validas):
                    link = get_direct_excel_link(url)
                    df_t = pd.read_excel(link, sheet_name="EMPRESA")
                    df_t.columns = df_t.columns.str.strip()
                    for c in ['NIT', 'EMPRESA', 'CANT EMPLEADOS', 'PERIODO', 'SEGURIDAD SOCIAL', 'TIPO']:
                        if c not in df_t.columns: df_t[c] = "" if c in ['EMPRESA', 'PERIODO', 'TIPO'] else 0
                    
                    df_t['NIT'] = df_t['NIT'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                    df_t['SEGURIDAD SOCIAL'] = pd.to_numeric(df_t['SEGURIDAD SOCIAL'], errors='coerce').fillna(0)
                    df_t['CANT EMPLEADOS'] = pd.to_numeric(df_t['CANT EMPLEADOS'], errors='coerce').fillna(0)
                            
                    agg_rules = {'EMPRESA':'first', 'PERIODO':'first', 'TIPO':'first', 'CANT EMPLEADOS':'sum', 'SEGURIDAD SOCIAL':'sum'}
                    df_agg = df_t.groupby('NIT').agg(agg_rules).reset_index()
                    df_agg.rename(columns={'SEGURIDAD SOCIAL': f'SS_AP{i+1}', 'CANT EMPLEADOS': f'EMP_AP{i+1}'}, inplace=True)
                    dfs_aprop.append(df_agg)
                    
                df_m = pd.read_excel(get_direct_excel_link(url_ss_mensual), sheet_name="SSGCON")
                df_m.columns = df_m.columns.str.strip()
                for c in ['NIT', 'EMPRESA', 'CANT EMPLEADOS', 'PERIODO', 'SEGURIDAD SOCIAL', 'TIPO']:
                    if c not in df_m.columns: df_m[c] = "" if c in ['EMPRESA', 'PERIODO', 'TIPO'] else 0
                
                df_m['NIT'] = df_m['NIT'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                df_m['SEGURIDAD SOCIAL'] = pd.to_numeric(df_m['SEGURIDAD SOCIAL'], errors='coerce').fillna(0)
                df_m['CANT EMPLEADOS'] = pd.to_numeric(df_m['CANT EMPLEADOS'], errors='coerce').fillna(0)
                        
                df_m_agg = df_m.groupby('NIT').agg(agg_rules).reset_index()
                df_m_agg.rename(columns={'SEGURIDAD SOCIAL': 'SS_MENSUAL', 'CANT EMPLEADOS': 'EMP_M'}, inplace=True)
                
                cruce = dfs_aprop[0]
                if len(dfs_aprop) > 1: cruce = pd.merge(cruce, dfs_aprop[1], on='NIT', how='outer', suffixes=('', '_2'))
                if len(dfs_aprop) > 2: cruce = pd.merge(cruce, dfs_aprop[2], on='NIT', how='outer', suffixes=('', '_3'))
                cruce = pd.merge(cruce, df_m_agg, on='NIT', how='outer', suffixes=('', '_M'))
                
                emp_cols = [c for c in cruce.columns if 'EMPRESA' in c]
                per_cols = [c for c in cruce.columns if 'PERIODO' in c]
                tip_cols = [c for c in cruce.columns if 'TIPO' in c]
                cant_cols = [c for c in cruce.columns if 'EMP_' in c]
                
                cruce['EMPRESA_FINAL'] = cruce.apply(lambda r: coalesce_strings(r, emp_cols), axis=1)
                cruce['PERIODO_FINAL'] = cruce.apply(lambda r: coalesce_strings(r, per_cols), axis=1)
                cruce['TIPO_FINAL'] = cruce.apply(lambda r: coalesce_strings(r, tip_cols), axis=1)
                cruce['CANT_EMP_PROM'] = cruce.apply(lambda r: calc_promedio_empleados(r, cant_cols), axis=1)
                
                df_export = pd.DataFrame()
                df_export['NIT'] = cruce['NIT']
                df_export['EMPRESA'] = cruce['EMPRESA_FINAL']
                df_export['CANT EMPLEADOS'] = cruce['CANT_EMP_PROM']
                df_export['PERIODO'] = cruce['PERIODO_FINAL']
                df_export['SEGURIDAD SOCIAL AP1'] = pd.to_numeric(cruce['SS_AP1'], errors='coerce').fillna(0) if 'SS_AP1' in cruce else 0
                df_export['SEGURIDAD SOCIAL AP2'] = pd.to_numeric(cruce['SS_AP2'], errors='coerce').fillna(0) if 'SS_AP2' in cruce else 0
                df_export['SEGURIDAD SOCIAL AP3'] = pd.to_numeric(cruce['SS_AP3'], errors='coerce').fillna(0) if 'SS_AP3' in cruce else 0
                df_export['TOTAL SS APROPIADO'] = df_export['SEGURIDAD SOCIAL AP1'] + df_export['SEGURIDAD SOCIAL AP2'] + df_export['SEGURIDAD SOCIAL AP3']
                df_export['VALOR SEGURIDAD SOCIAL'] = pd.to_numeric(cruce['SS_MENSUAL'], errors='coerce').fillna(0) if 'SS_MENSUAL' in cruce else 0
                df_export['DIFERENCIA'] = df_export['TOTAL SS APROPIADO'] - df_export['VALOR SEGURIDAD SOCIAL']
                df_export['TIPO'] = cruce['TIPO_FINAL']
                
                st.session_state['df_cruce_2'] = df_export
                st.success("¡Cruce de Seguridad Social consolidado!")
                st.dataframe(df_export.head())
            except Exception as e:
                st.error(f"Error procesando. Detalle: {e}")
        else:
            st.warning("Pega al menos 1 URL de apropiación y el reporte mensual.")

    if 'df_cruce_2' in st.session_state:
        st.markdown("### ¿Qué deseas hacer con el resultado?")
        btn_col1, btn_col2 = st.columns(2)
        with btn_col1:
            st.download_button("📥 Descargar Reporte (Excel)", to_excel_ss(st.session_state['df_cruce_2']), "Cruce_Seguridad_Social.xlsx", "application/vnd.ms-excel", use_container_width=True)
        with btn_col2:
            if st.button("☁️ Guardar en hoja SEGURIDAD SOCIAL", use_container_width=True):
                with st.spinner("Conectando con Google Drive..."):
                    exito, mensaje = guardar_en_drive(st.session_state['df_cruce_2'], "SEGURIDAD SOCIAL")
                    if exito: st.success(mensaje)
                    else: st.error(mensaje)

# ==========================================
# GESTIÓN 3: FIC
# ==========================================
with tab3:
    st.header("Cruce: FIC")
    st.markdown("Pega de 1 a 3 enlaces de apropiación y tu enlace mensual de FIC.")
    
    col_fic1, col_fic2, col_fic3 = st.columns(3)
    with col_fic1: url_fic_ap1 = st.text_input("URL Apropiación 1:", key="url_fic_ap1")
    with col_fic2: url_fic_ap2 = st.text_input("URL Apropiación 2 (Opcional):", key="url_fic_ap2")
    with col_fic3: url_fic_ap3 = st.text_input("URL Apropiación 3 (Opcional):", key="url_fic_ap3")
        
    url_fic_mensual = st.text_input("URL Reporte Mensual de FIC:", key="url_fic_mensual")
    
    if st.button("Generar Cruce FIC", type="primary"):
        urls_validas = [u for u in [url_fic_ap1, url_fic_ap2, url_fic_ap3] if u.strip() != ""]
        if urls_validas and url_fic_mensual:
            try:
                dfs_aprop = []
                for i, url in enumerate(urls_validas):
                    link = get_direct_excel_link(url)
                    df_t = pd.read_excel(link, sheet_name="EMPRESA")
                    df_t.columns = df_t.columns.str.strip()
                    for c in ['NIT', 'EMPRESA', 'CANT EMPLEADOS', 'PERIODO', 'FIC', 'TIPO']:
                        if c not in df_t.columns: df_t[c] = "" if c in ['EMPRESA', 'PERIODO', 'TIPO'] else 0
                    
                    df_t['NIT'] = df_t['NIT'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                    df_t['FIC'] = pd.to_numeric(df_t['FIC'], errors='coerce').fillna(0)
                    df_t['CANT EMPLEADOS'] = pd.to_numeric(df_t['CANT EMPLEADOS'], errors='coerce').fillna(0)
                            
                    agg_rules = {'EMPRESA':'first', 'PERIODO':'first', 'TIPO':'first', 'CANT EMPLEADOS':'sum', 'FIC':'sum'}
                    df_agg = df_t.groupby('NIT').agg(agg_rules).reset_index()
                    df_agg.rename(columns={'FIC': f'FIC_AP{i+1}', 'CANT EMPLEADOS': f'EMP_AP{i+1}'}, inplace=True)
                    dfs_aprop.append(df_agg)
                    
                df_m = pd.read_excel(get_direct_excel_link(url_fic_mensual), sheet_name="FICCON")
                df_m.columns = df_m.columns.str.strip()
                for c in ['NIT', 'EMPRESA', 'CANT EMPLEADOS', 'PERIODO', 'FIC', 'TIPO']:
                    if c not in df_m.columns: df_m[c] = "" if c in ['EMPRESA', 'PERIODO', 'TIPO'] else 0
                
                df_m['NIT'] = df_m['NIT'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                df_m['FIC'] = pd.to_numeric(df_m['FIC'], errors='coerce').fillna(0)
                df_m['CANT EMPLEADOS'] = pd.to_numeric(df_m['CANT EMPLEADOS'], errors='coerce').fillna(0)        
                        
                df_m_agg = df_m.groupby('NIT').agg(agg_rules).reset_index()
                df_m_agg.rename(columns={'FIC': 'FIC_MENSUAL', 'CANT EMPLEADOS': 'EMP_M'}, inplace=True)
                
                cruce = dfs_aprop[0]
                if len(dfs_aprop) > 1: cruce = pd.merge(cruce, dfs_aprop[1], on='NIT', how='outer', suffixes=('', '_2'))
                if len(dfs_aprop) > 2: cruce = pd.merge(cruce, dfs_aprop[2], on='NIT', how='outer', suffixes=('', '_3'))
                cruce = pd.merge(cruce, df_m_agg, on='NIT', how='outer', suffixes=('', '_M'))
                
                emp_cols = [c for c in cruce.columns if 'EMPRESA' in c]
                per_cols = [c for c in cruce.columns if 'PERIODO' in c]
                tip_cols = [c for c in cruce.columns if 'TIPO' in c]
                cant_cols = [c for c in cruce.columns if 'EMP_' in c]
                
                cruce['EMPRESA_FINAL'] = cruce.apply(lambda r: coalesce_strings(r, emp_cols), axis=1)
                cruce['PERIODO_FINAL'] = cruce.apply(lambda r: coalesce_strings(r, per_cols), axis=1)
                cruce['TIPO_FINAL'] = cruce.apply(lambda r: coalesce_strings(r, tip_cols), axis=1)
                cruce['CANT_EMP_PROM'] = cruce.apply(lambda r: calc_promedio_empleados(r, cant_cols), axis=1)
                
                df_export = pd.DataFrame()
                df_export['NIT'] = cruce['NIT']
                df_export['EMPRESA'] = cruce['EMPRESA_FINAL']
                df_export['CANT EMPLEADOS'] = cruce['CANT_EMP_PROM']
                df_export['PERIODO'] = cruce['PERIODO_FINAL']
                df_export['FIC AP1'] = pd.to_numeric(cruce['FIC_AP1'], errors='coerce').fillna(0) if 'FIC_AP1' in cruce else 0
                df_export['FIC AP2'] = pd.to_numeric(cruce['FIC_AP2'], errors='coerce').fillna(0) if 'FIC_AP2' in cruce else 0
                df_export['FIC AP3'] = pd.to_numeric(cruce['FIC_AP3'], errors='coerce').fillna(0) if 'FIC_AP3' in cruce else 0
                df_export['TOTAL FIC APROPIADO'] = df_export['FIC AP1'] + df_export['FIC AP2'] + df_export['FIC AP3']
                df_export['VALOR FIC'] = pd.to_numeric(cruce['FIC_MENSUAL'], errors='coerce').fillna(0) if 'FIC_MENSUAL' in cruce else 0
                df_export['DIFERENCIA'] = df_export['TOTAL FIC APROPIADO'] - df_export['VALOR FIC']
                df_export['TIPO'] = cruce['TIPO_FINAL']
                
                st.session_state['df_cruce_3'] = df_export
                st.success("¡Cruce de FIC consolidado!")
                st.dataframe(df_export.head())
            except Exception as e:
                st.error(f"Error procesando. Detalle: {e}")
        else:
            st.warning("Pega al menos 1 URL de apropiación y el reporte mensual.")

    if 'df_cruce_3' in st.session_state:
        st.markdown("### ¿Qué deseas hacer con el resultado?")
        btn_col1, btn_col2 = st.columns(2)
        with btn_col1:
            st.download_button("📥 Descargar Reporte (Excel)", to_excel_fic(st.session_state['df_cruce_3']), "Cruce_FIC.xlsx", "application/vnd.ms-excel", use_container_width=True)
        with btn_col2:
            if st.button("☁️ Guardar en hoja FIC", use_container_width=True):
                with st.spinner("Conectando con Google Drive..."):
                    exito, mensaje = guardar_en_drive(st.session_state['df_cruce_3'], "FIC")
                    if exito: st.success(mensaje)
                    else: st.error(mensaje)
