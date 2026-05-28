import streamlit as st
import pandas as pd

# Configuración inicial de la página
st.set_page_config(page_title="Dashboard ACL", layout="wide")

st.title("📊 Dashboard de Gestión de Datos")

# --- CONFIGURACIÓN DE LAS URLs ---
# Sustituye estas URLs por las de tus 3 Google Sheets reales
# Asegúrate de que las hojas sean públicas (archivo -> compartir -> cualquier persona con el enlace)
URLS = [
    "TU_URL_HOJA_1",
    "TU_URL_HOJA_2",
    "TU_URL_HOJA_3"
]

# --- FUNCIÓN DE CARGA ---
@st.cache_data(ttl=600)
def cargar_datos(urls):
    lista_dfs = []
    for url in urls:
        # Convertimos la URL para descargar como CSV
        csv_url = url.replace("/edit#gid=", "/export?format=csv&gid=")
        df = pd.read_csv(csv_url)
        lista_dfs.append(df)
    
    # Unimos todas las hojas en una sola
    return pd.concat(lista_dfs, ignore_index=True)

# --- EJECUCIÓN ---
try:
    with st.spinner('Cargando datos de Google Sheets...'):
        df_completo = cargar_datos(URLS)

    # Buscador
    st.subheader("Buscador General")
    query = st.text_input("Escribe para filtrar los datos:")

    if query:
        # Filtra el DataFrame si hay texto en el buscador
        df_filtrado = df_completo[df_completo.apply(lambda row: row.astype(str).str.contains(query, case=False).any(), axis=1)]
        st.write(f"Resultados encontrados: {len(df_filtrado)}")
        st.dataframe(df_filtrado, use_container_width=True)
    else:
        # Muestra todo si no hay nada escrito
        st.dataframe(df_completo, use_container_width=True)

except Exception as e:
    st.error(f"Error al cargar los datos: {e}")
    st.info("Asegúrate de que las URLs de tus Google Sheets sean correctas y estén compartidas como 'Cualquier persona con el enlace puede leer'.")
