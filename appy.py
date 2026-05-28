import streamlit as st
import gspread
import pandas as pd

st.set_page_config(page_title="Dashboard", layout="wide")

@st.cache_data(ttl=600)
def load_data():
    creds = st.secrets["gcp_service_account"]
    gc = gspread.service_account_from_dict(creds)
    
    # 1. Abrir por ID
    sh = gc.open_by_key("1Sm-0JmTghLfQczZ3k-W57JNZRbH10waNQj6HVOXZdgk")
    
    # 2. Acceder a la hoja específicamente por nombre
    # Si sigue dando error, asegúrate de que el nombre sea EXACTAMENTE "registro visitas"
    ws = sh.worksheet("registro visitas")
    
    return pd.DataFrame(ws.get_all_records())

st.title("📋 Registro de Visitas")

try:
    df = load_data()
    st.dataframe(df, use_container_width=True)
except Exception as e:
    st.error(f"Error al cargar la hoja: {e}")
    st.write("Asegúrate de que la pestaña se llame exactamente 'registro visitas' y que el bot tenga permisos de Lector.")
