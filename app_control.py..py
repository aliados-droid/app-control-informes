import streamlit as st
import pandas as pd
import io
import json
import gspread

# Configuración de la página
st.set_page_config(page_title="Control Gerencial", page_icon="📊", layout="wide")
st.title("📊 Sistema de Control Gerencial: Apropiado vs Pagado")
st.markdown("Automatización de cruces de información por NIT para el control de pagos y apropiaciones.")

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Cruce Consolidado')
    return output.getvalue()

def to_excel_tab1(df):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='CONSOLIDADO NACIONAL')
    
    workbook = writer.book
    worksheet = writer.sheets['CONSOLIDADO NACIONAL']
    
    header_format = workbook.add_format({
        'bg_color': '#C00000',
        'font_color': 'white',
        'bold': True,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True
    })
    
    worksheet.set_row(0, 35) 
    
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)
        
    for row_num in range(1, len(df) + 1):
        excel_row = row_num + 1 
        # Columna G (Total Apropiacion) = E (Nomina Aprop) + F (Prestaciones Aprop)
        worksheet.write_formula(row_num, 6, f'=E{excel_row}+F{excel_row}')
        
        # Columna K (Diferencia) = G (Total Aprop) - J (Total Aprobacion) validando con F (Prestaciones Aprop)
        formula_dif = f'=IF(G{excel_row}-J{excel_row}<0, IF(ABS(G{excel_row}-J{excel_row})<=F{excel_row}, 0, G{excel_row}-J{excel_row}), G{excel_row}-J{excel_row})'
        worksheet.write_formula(row_num, 10, formula_dif)
        
    worksheet.set_column('A:A', 15) # NIT
    worksheet.set_column('B:B', 35) # EMPRESA
    worksheet.set_column('C:C', 16) # CANT EMPLEADOS
    worksheet.set_column('D:D', 20) # PERIODO
    worksheet.set_column('E:J', 18) # VALORES MONETARIOS
    worksheet.set_column('K:K', 15) # DIFERENCIA
    worksheet.set_column('L:L', 20) # TIPO
    
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

# FUNCIÓN PARA GUARDAR EN GOOGLE SHEETS
def guardar_en_drive(dataframe):
    try:
        cred_dict = json.loads(st.secrets["google_credentials"])
        gc = gspread.service_account_from_dict(cred_dict)
        
        # 👇 ¡OJO AQUÍ! CAMBIA ESTO POR EL ID DE TU NUEVO ARCHIVO DE GOOGLE SHEETS 👇
        sh = gc.open_by_key("1W0zekg36sXn3z_n4Vpa976yS59lLhLab4T9BwzU3G_s")
        
        worksheet = sh.worksheet("NOMINA")
        
        df_clean = safe_fillna(dataframe.copy())
        datos_para_enviar = df_clean.values.tolist()
        
        worksheet.append_rows(datos_para_enviar)
        
        return True, "¡Datos guardados exitosamente en Google Drive!"
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
                
                # SE AGREGA 'PERIODO' Y 'TIPO' A LAS COLUMNAS REQUERIDAS DE APROPIACIÓN
                aprop_cols = ['NIT', 'EMPRESA', 'PERIODO', 'CANT EMPLEADOS', 'NOMINA', 'PRESTACIONES', 'TIPO']
                for c in aprop_cols:
                    if c not in df_aprop.columns:
                        df_aprop[c] = 0 if c not in ['EMPRESA', 'PERIODO', 'TIPO'] else ""
                        
                agg_aprop = {'EMPRESA': 'first', 'PERIODO': 'first', 'CANT EMPLEADOS': 'sum', 'NOMINA': 'sum', 'PRESTACIONES': 'sum', 'TIPO': 'first'}
                df_aprop_agrupado = df_aprop.groupby('NIT').agg(agg_aprop).reset_index()
                
                pago_cols = ['NIT', 'EMPRESA', 'CANT EMPLEADOS', 'NOMINA', 'PRESTACIONES', 'TOTAL']
                for c in pago_cols:
                    if c not in df_pago.columns:
                        df_pago[c] = 0 if c != 'EMPRESA' else ""
                        
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
                    
                    if a_val > 0 and p_val > 0:
                        cant_promedio.append(int(round((a_val + p_val) / 2)))
                    elif a_val > 0:
                        cant_promedio.append(int(a_val))
                    else:
                        cant_promedio.append(int(p_val))
                        
                df_export = pd.DataFrame()
                df_export['NIT'] = cruce_1['NIT']
                df_export['EMPRESA'] = cruce_1['EMPRESA_Apropiado'].where(cruce_1['EMPRESA_Apropiado'] != "", cruce_1['EMPRESA_Pagado'])
                
                df_export['CANT EMPLEADOS'] = cant_promedio
                df_export['PERIODO'] = cruce_1.get('PERIODO', "")
                
                df_export['NOMINA\nAPROPIACION'] = cruce_1['NOMINA_Apropiado']
