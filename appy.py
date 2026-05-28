import streamlit as st
import gspread
import pandas as pd

# Configuración de página
st.set_page_config(page_title="Registro de Visitas", layout="wide")

@st.cache_data(ttl=600)
def load_data():
    # Cargar credenciales desde los secretos de Streamlit
    creds = st.secrets["gcp_service_account"]
    
    # Autorizar
    gc = gspread.service_account_from_dict(creds)
    
    # Abrir por ID (el ID es lo que está entre /d/ y /edit en tu URL)
    sh = gc.open_by_key("1Sm-0JmTghLfQczZ3k-W57JNZRbH10waNQj6HVOXZdgk")
    
    # Intentar abrir la hoja por nombre o por índice si el nombre falla
    try:
        ws = sh.worksheet("registro de visitas")
    except:
        ws = sh.get_worksheet(0) # Abre la primera pestaña si la anterior falla
        
    return pd.DataFrame(ws.get_all_records())

# Interfaz
st.title("📋 Registro de Visitas")

try:
    df = load_data()
    st.dataframe(df, use_container_width=True)
except Exception as e:
    st.error(f"Error al conectar: {e}")
    st.write("Verifica que el correo 'lectorappsheet@...' tenga permisos de Lector en el botón Compartir de tu Google Sheet.")
