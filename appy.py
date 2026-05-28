import streamlit as st
import gspread
import pandas as pd

st.set_page_config(layout="wide")

@st.cache_data(ttl=600)
def load_data():
    # Cargar credenciales desde los secretos
    creds = st.secrets["gcp_service_account"]
    
    # Conectar
    gc = gspread.service_account_from_dict(creds)
    # Abre por ID de hoja
    sh = gc.open_by_key("1Sm-0JmTghLfQczZ3k-W57JNZRbH10waNQj6HVOXZdgk")
    # Selecciona la pestaña "registro de visitas" (cambia el índice si es necesario)
    ws = sh.worksheet("registro de visitas") 
    
    return pd.DataFrame(ws.get_all_records())

st.title("Registro de Visitas (Privado)")
df = load_data()
st.dataframe(df)
