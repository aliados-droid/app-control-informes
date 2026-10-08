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

# FUNCIÓN PARA GUARDAR EN GOOGLE SHEETS
def guardar_en_drive(dataframe, sheet_name):
    try:
        cred_dict = json.loads(st.secrets["google_credentials"])
        gc = gspread.service_account_from_dict(cred_dict)
        # ID DEL CONSOLIDADO NACIONAL DIRECTO
        sh = gc.open_by_key("1W0zekg36sXn3z_n4Vpa976yS59lLhLab4T9BwzU3G_s")
        worksheet = sh.worksheet(sheet_name)
        
        df_clean = safe_fillna(dataframe.copy())
        datos_para_enviar = df_clean.values.tolist()
        works
