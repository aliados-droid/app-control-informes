import streamlit as st
import pandas as pd
import io

# Configuración de la página
st.set_page_config(page_title="Control Gerencial", page_icon="📊", layout="wide")
st.title("📊 Sistema de Control Gerencial: Apropiado vs Pagado")
st.markdown("Automatización de cruces de información por NIT para el control de pagos y apropiaciones.")

# Función auxiliar para convertir DataFrame a Excel descargable
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Cruce Consolidado')
    return output.getvalue()

# Creación de pestañas para las 3 gestiones
tab1, tab2, tab3 = st.tabs(["1️⃣ Apropiación vs Pago", "2️⃣ Seguridad Social", "3️⃣ FIC"])

# ==========================================
# GESTIÓN 1: APROPIACIÓN VS PAGO
# ==========================================
with tab1:
    st.header("Cruce: Apropiación vs Pago")
    col1, col2 = st.columns(2)
    
    with col1:
        file_aprop_1 = st.file_uploader("Sube el archivo de Apropiación (o pega la URL de Drive)", type=["xlsx", "xls"], key="aprop1")
        ciudad = st.selectbox("Selecciona la región del pago:", ["Bogotá (NOMCONBOG)", "Eje Cafetero (NOMCONEJE)"])
    with col2:
        file_pago_1 = st.file_uploader("Sube el archivo de Aprobación de Pago", type=["xlsx", "xls"], key="pago1")
    
    if st.button("Generar Cruce Apropiación vs Pago", type="primary"):
        if file_aprop_1 and file_pago_1:
            try:
                # Lectura de las hojas específicas
                df_aprop = pd.read_excel(file_aprop_1, sheet_name="EMPRESA")
                
                hoja_pago = "NOMCONBOG" if "Bogotá" in ciudad else "NOMCONEJE"
                df_pago = pd.read_excel(file_pago_1, sheet_name=hoja_pago)
                
                # Tratamiento de ausencias (llenar nulos con 0)
                df_aprop.fillna(0, inplace=True)
                df_pago.fillna(0, inplace=True)
                
                # Cruce de datos por NIT (Outer join para ver diferencias)
                cruce_1 = pd.merge(df_aprop, df_pago, on="NIT", how="outer", suffixes=('_Apropiado', '_Pagado'))
                cruce_1.fillna(0, inplace=True)
                
                st.success("¡Cruce realizado con éxito!")
                st.dataframe(cruce_1.head())
                
                st.download_button(
                    label="📥 Descargar Reporte en Excel",
                    data=to_excel(cruce_1),
                    file_name="Cruce_Apropiacion_vs_Pago.xlsx",
                    mime="application/vnd.ms-excel"
                )
            except Exception as e:
                st.error(f"Error procesando los archivos: {e}")
        else:
            st.warning("Por favor, sube ambos archivos para continuar.")

# ==========================================
# GESTIÓN 2: SEGURIDAD SOCIAL
# ==========================================
with tab2:
    st.header("Cruce: Seguridad Social")
    st.markdown("Puedes subir de 1 a 3 archivos de apropiación.")
    
    col_ap1, col_ap2, col_ap3 = st.columns(3)
    with col_ap1:
        ss_aprop_1 = st.file_uploader("Apropiación 1", type=["xlsx", "xls"], key="ss_ap1")
    with col_ap2:
        ss_aprop_2 = st.file_uploader("Apropiación 2 (Opcional)", type=["xlsx", "xls"], key="ss_ap2")
    with col_ap3:
        ss_aprop_3 = st.file_uploader("Apropiación 3 (Opcional)", type=["xlsx", "xls"], key="ss_ap3")
        
    ss_mensual = st.file_uploader("Sube el Reporte Mensual de Seguridad Social", type=["xlsx", "xls"], key="ss_mensual")
    
    if st.button("Generar Cruce Seguridad Social", type="primary"):
        archivos_ss = [f for f in [ss_aprop_1, ss_aprop_2, ss_aprop_3] if f is not None]
        
        if archivos_ss and ss_mensual:
            try:
                # Consolidar apropiaciones
                lista_df_aprop = []
                for file in archivos_ss:
                    df_temp = pd.read_excel(file, sheet_name="EMPRESA")
                    lista_df_aprop.append(df_temp[["NIT", "SEGURIDAD SOCIAL"]])
                
                df_aprop_consol = pd.concat(lista_df_aprop)
                # Sumarizamos la seguridad social por NIT
                df_aprop_agrupado = df_aprop_consol.groupby("NIT")["SEGURIDAD SOCIAL"].sum().reset_index()
                df_aprop_agrupado.rename(columns={"SEGURIDAD SOCIAL": "SEGURIDAD_SOCIAL_APROPIADA"}, inplace=True)
                
                # Leer reporte mensual
                df_ss_mensual = pd.read_excel(ss_mensual, sheet_name="SSGCON")
                
                # Cruce por NIT
                cruce_2 = pd.merge(df_aprop_agrupado, df_ss_mensual, on="NIT", how="outer")
                cruce_2.fillna(0, inplace=True)
                
                st.success("¡Cruce de Seguridad Social consolidado!")
                st.dataframe(cruce_2.head())
                
                st.download_button(
                    label="📥 Descargar Reporte Seguridad Social",
                    data=to_excel(cruce_2),
                    file_name="Cruce_Seguridad_Social.xlsx"
                )
            except Exception as e:
                st.error(f"Error procesando los archivos: {e}")
        else:
            st.warning("Debes subir al menos 1 archivo de apropiación y el reporte mensual.")

# ==========================================
# GESTIÓN 3: FIC
# ==========================================
with tab3:
    st.header("Cruce: FIC")
    st.markdown("Puedes subir de 1 a 3 archivos de apropiación.")
    
    col_fic1, col_fic2, col_fic3 = st.columns(3)
    with col_fic1:
        fic_aprop_1 = st.file_uploader("Apropiación 1", type=["xlsx", "xls"], key="fic_ap1")
    with col_fic2:
        fic_aprop_2 = st.file_uploader("Apropiación 2 (Opcional)", type=["xlsx", "xls"], key="fic_ap2")
    with col_fic3:
        fic_aprop_3 = st.file_uploader("Apropiación 3 (Opcional)", type=["xlsx", "xls"], key="fic_ap3")
        
    fic_mensual = st.file_uploader("Sube el Reporte Mensual de FIC", type=["xlsx", "xls"], key="fic_mensual")
    
    if st.button("Generar Cruce FIC", type="primary"):
        archivos_fic = [f for f in [fic_aprop_1, fic_aprop_2, fic_aprop_3] if f is not None]
        
        if archivos_fic and fic_mensual:
            try:
                # Consolidar apropiaciones
                lista_df_fic = []
                for file in archivos_fic:
                    df_temp = pd.read_excel(file, sheet_name="EMPRESA")
                    lista_df_fic.append(df_temp[["NIT", "FIC"]])
                
                df_fic_consol = pd.concat(lista_df_fic)
                # Sumarizamos el FIC por NIT
                df_fic_agrupado = df_fic_consol.groupby("NIT")["FIC"].sum().reset_index()
                df_fic_agrupado.rename(columns={"FIC": "FIC_APROPIADO"}, inplace=True)
                
                # Leer reporte mensual
                df_fic_mensual = pd.read_excel(fic_mensual, sheet_name="EMPRESA")
                
                # Cruce por NIT
                cruce_3 = pd.merge(df_fic_agrupado, df_fic_mensual, on="NIT", how="outer")
                cruce_3.fillna(0, inplace=True)
                
                st.success("¡Cruce de FIC consolidado!")
                st.dataframe(cruce_3.head())
                
                st.download_button(
                    label="📥 Descargar Reporte FIC",
                    data=to_excel(cruce_3),
                    file_name="Cruce_FIC.xlsx"
                )
            except Exception as e:
                st.error(f"Error procesando los archivos: {e}")
        else:
            st.warning("Debes subir al menos 1 archivo de apropiación y el reporte mensual.")