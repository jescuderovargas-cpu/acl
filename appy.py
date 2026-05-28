import streamlit as st
import gspread
import pandas as pd

st.set_page_config(page_title="Diagnóstico", layout="wide")

@st.cache_data(ttl=600)
def diagnostico_conexiones():
    creds = st.secrets["gcp_service_account"]
    gc = gspread.service_account_from_dict(creds)
    
    # 1. Intentar abrir el archivo
    sh = gc.open_by_key("1Sm-0JmTghLfQczZ3k-W57JNZRbH10waNQj6HVOXZdgk")
    
    # 2. Obtener lista de nombres de pestañas para ver si "registro visitas" existe
    lista_pestañas = [ws.title for ws in sh.worksheets()]
    
    # 3. Intentar cargar la primera hoja
    ws = sh.get_worksheet(0)
    df = pd.DataFrame(ws.get_all_records())
    
    return lista_pestañas, df

st.title("🔍 Diagnóstico de Conexión")

try:
    pestañas, df = diagnostico_conexiones()
    st.success("¡Conexión exitosa!")
    st.write("Pestañas encontradas en el archivo:", pestañas)
    st.dataframe(df)
except Exception as e:
    st.error(f"Error detectado: {e}")
