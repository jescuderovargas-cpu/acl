import streamlit as st
import pandas as pd

# Configuración de página
st.set_page_config(page_title="Dashboard ACL", layout="wide")

# Lista de tus enlaces publicados como CSV
URLS = [
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vSUgGBm_yYSbJpK3-Sz-mftj2qhPNaUZJ4L7pV7PdlgDm16m0WFqX5rLxWj-rcJ06SN8TbZAPJvoz-d/pub?output=csv"
    # Si tienes más hojas, añade las URLs aquí separadas por comas
]

@st.cache_data(ttl=600)
def cargar_datos(urls):
    lista_dfs = []
    for url in urls:
        try:
            df = pd.read_csv(url)
            lista_dfs.append(df)
        except Exception as e:
            st.warning(f"No se pudo cargar una de las hojas: {e}")
    
    if lista_dfs:
        return pd.concat(lista_dfs, ignore_index=True)
    else:
        return pd.DataFrame()

# --- INTERFAZ ---
st.title("📊 Dashboard de Datos")

df = cargar_datos(URLS)

if not df.empty:
    # Buscador simple
    query = st.text_input("🔍 Buscar en los datos:")
    
    if query:
        # Filtra el DataFrame
        df_filtrado = df[df.apply(lambda row: row.astype(str).str.contains(query, case=False).any(), axis=1)]
        st.dataframe(df_filtrado, use_container_width=True)
    else:
        st.dataframe(df, use_container_width=True)
else:
    st.write("No hay datos disponibles para mostrar.")
