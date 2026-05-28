import streamlit as st
import pandas as pd

# Usamos la URL pública de "Publicar en la web" que funciona siempre
# si el documento está compartido como "Cualquier persona con el enlace puede leer"
url = "https://docs.google.com/spreadsheets/d/1Sm-0JmTghLfQczZ3k-W57JNZRbH10waNQj6HVOXZdgk/export?format=csv&gid=0"

st.title("Prueba de carga directa")

try:
    df = pd.read_csv(url)
    st.success("¡Datos cargados correctamente!")
    st.dataframe(df)
except Exception as e:
    st.error(f"Error: {e}")
    st.write("Si esto falla, el problema es que el archivo NO es accesible desde internet.")
