import streamlit as st
import pandas as pd

# Configuración de la página
st.set_page_config(page_title="Registro de Visitas", layout="wide")

# URL de exportación directa (formato CSV)
# Esta URL es la forma más rápida y efectiva de leer una hoja privada 
# configurada como "Cualquier persona con el enlace"
URL_HOJA = "https://docs.google.com/spreadsheets/d/1Sm-0JmTghLfQczZ3k-W57JNZRbH10waNQj6HVOXZdgk/export?format=csv&gid=0"

@st.cache_data(ttl=600)
def cargar_datos():
    try:
        # Leemos el CSV directamente desde Google Sheets
        df = pd.read_csv(URL_HOJA)
        return df
    except Exception as e:
        st.error(f"Error al conectar con la hoja: {e}")
        return pd.DataFrame()

# Título y carga
st.title("📋 Registro de Visitas")

df = cargar_datos()

if not df.empty:
    # Buscador interactivo
    st.subheader("Buscador de registros")
    query = st.text_input("Escribe para filtrar:")

    if query:
        # Filtra el DataFrame basándose en el texto introducido
        df_filtrado = df[df.apply(lambda row: row.astype(str).str.contains(query, case=False).any(), axis=1)]
        st.write(f"Resultados encontrados: {len(df_filtrado)}")
        st.dataframe(df_filtrado, use_container_width=True)
    else:
        st.dataframe(df, use_container_width=True)
else:
    st.info("No se pudieron cargar los datos. Verifica que el enlace sea correcto y los permisos estén abiertos.")
