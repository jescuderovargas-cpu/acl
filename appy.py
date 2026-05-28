import streamlit as st
import pandas as pd

# Configuración de página
st.set_page_config(page_title="Dashboard Registro de Visitas", layout="wide")

# URL directa para exportar tu hoja específica (gid=0) como CSV
URLS = [
    "https://docs.google.com/spreadsheets/d/1Sm-0JmTghLfQczZ3k-W57JNZRbH10waNQj6HVOXZdgk/export?format=csv&gid=0"
]

@st.cache_data(ttl=600)
def cargar_datos(urls):
    lista_dfs = []
    for url in urls:
        try:
            # Leemos directamente el CSV desde la URL
            df = pd.read_csv(url)
            lista_dfs.append(df)
        except Exception as e:
            st.error(f"Error al conectar con Google Sheets: {e}")
    
    if lista_dfs:
        return pd.concat(lista_dfs, ignore_index=True)
    return pd.DataFrame()

# --- INTERFAZ ---
st.title("📋 Registro de Visitas")

# Carga de datos
df = cargar_datos(URLS)

if not df.empty:
    # Buscador interactivo
    st.subheader("Buscador")
    query = st.text_input("Escribe para filtrar los registros:")

    if query:
        # Filtra el DataFrame: busca en todas las columnas si el texto coincide
        df_filtrado = df[df.apply(lambda row: row.astype(str).str.contains(query, case=False).any(), axis=1)]
        st.write(f"Resultados encontrados: {len(df_filtrado)}")
        st.dataframe(df_filtrado, use_container_width=True)
    else:
        st.dataframe(df, use_container_width=True)
else:
    st.info("Cargando datos o el documento está vacío. Asegúrate de que el enlace sea correcto.")
