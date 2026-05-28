import streamlit as st
import gspread
import pandas as pd

@st.cache_data(ttl=600)
def load_data():
    # Cargar credenciales desde los secretos
    creds = st.secrets["gcp_service_account"]
    
    # Conectar
    gc = gspread.service_account_from_dict(creds)
    
    # Abrir el archivo
    sh = gc.open_by_key("1Sm-0JmTghLfQczZ3k-W57JNZRbH10waNQj6HVOXZdgk")
    
    # Obtener la hoja
    ws = sh.worksheet("registro visitas")
    
    # Obtener datos
    data = ws.get_all_records()
    return pd.DataFrame(data)

st.title("Registro de Visitas (Privado)")
df = load_data()
st.dataframe(df)
