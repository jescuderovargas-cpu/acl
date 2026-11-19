import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

try:
    import plotly.express as px
    HAS_PLOTLY = True
except ImportError:
    HAS_PLOTLY = False

# ---------------------------------------------------------
# 1. CONFIGURACIÓN GENERAL
# ---------------------------------------------------------
st.set_page_config(
    page_title="Cuadro de Mandos — Registro de Visitas",
    page_icon="🌿",
    layout="wide",
)

SHEET_ID = "1Sm-0JmTghLfQczZ3k-W57JNZRbH10waNQj6HVOXZdgk"
NOMBRE_PESTANA = "Registro visitas"
ARCHIVO_LOCAL_FALLBACK = "Tecnicos.xlsx"


# ---------------------------------------------------------
# 2. CONEXIÓN GOOGLE SHEETS Y CARGA CON CACHÉ
# ---------------------------------------------------------
@st.cache_resource
def conectar_gsheets():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly",
    ]
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scopes,
    )
    return gspread.authorize(creds)


@st.cache_data(ttl=600)
def cargar_datos_gsheet(sheet_id: str, pestana: str) -> pd.DataFrame:
    cliente = conectar_gsheets()
    ws = cliente.open_by_key(sheet_id).worksheet(pestana)
    return pd.DataFrame(ws.get_all_records())


@st.cache_data(ttl=600)
def cargar_datos_excel_local(ruta_excel: str, pestana: str) -> pd.DataFrame:
    return pd.read_excel(ruta_excel, sheet_name=pestana)


def limpiar_visitas(df_raw: pd.DataFrame) -> pd.DataFrame:
    df = df_raw.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # Eliminar filas vacías sin Id ni Fecha
    if "Id" in df.columns and "Fecha" in df.columns:
        df = df.dropna(subset=["Id", "Fecha"], how="all")
        df = df[df["Id"].astype(str).str.strip() != ""]

    # Convertir fechas y semana
    if "Fecha" in df.columns:
        df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
        df = df.dropna(subset=["Fecha"])

    if "Próxima visita" in df.columns:
        df["Próxima visita"] = pd.to_datetime(df["Próxima visita"], errors="coerce")

    if "Semana" in df.columns:
        df["Semana"] = pd.to_numeric(df["Semana"], errors="coerce")

    # Limpiar espacios en campos clave
    for col in ["Técnico/a", "Agricultor/a", "Tipo de agricultor", "Género", "Campaña"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            df.loc[df[col].isin(["nan", "None", ""]), col] = None

    # Indicadores de adjuntos
    df["Tiene foto"] = (
        df["Adjuntar foto"].notna() & (df["Adjuntar foto"].astype(str).str.strip() != "")
        if "Adjuntar foto" in df.columns
        else False
    )
    df["Tiene archivo"] = (
        df["Adjuntar archivo"].notna() & (df["Adjuntar archivo"].astype(str).str.strip() != "")
        if "Adjuntar archivo" in df.columns
        else False
    )

    return df.sort_values("Fecha", ascending=False)


# ---------------------------------------------------------
# 3. BARRA LATERAL (ORIGEN Y ACTUALIZACIÓN)
# ---------------------------------------------------------
st.sidebar.title("⚙️ Configuración")

origen = st.sidebar.radio(
    "Origen de datos",
    options=["Google Sheets", "Excel local (Tecnicos.xlsx)"],
    index=0,
)

if st.sidebar.button("🔄 Actualizar datos", use_container_width=True):
    st.cache_data.clear()
    st.rerun()

try:
    if origen == "Google Sheets":
        df_raw = cargar_datos_gsheet(SHEET_ID, NOMBRE_PESTANA)
    else:
        df_raw = cargar_datos_excel_local(ARCHIVO_LOCAL_FALLBACK, NOMBRE_PESTANA)
except Exception as e:
    st.error(f"No se pudo cargar '{NOMBRE_PESTANA}': {e}")
    st.info(
        "Revisa el ID del Sheet, el nombre de la pestaña, que esté compartida con "
        "`firebase-app-hosting-compute@studio-5810984322-26d93.iam.gserviceaccount.com` "
        "y que `st.secrets['gcp_service_account']` esté configurado."
    )
    st.stop()

df = limpiar_visitas(df_raw)
if df.empty:
    st.warning("No hay registros válidos tras la limpieza.")
    st.stop()


# ---------------------------------------------------------
# 4. CABECERA Y FILTROS EN LA PARTE SUPERIOR
# ---------------------------------------------------------
st.title("📋 Cuadro de Mandos — Registro de Visitas")

with st.container(border=True):
    st.subheader("🔎 Filtros")
    f1, f2, f3, f4 = st.columns(4)

    with f1:
        if "Campaña" in df.columns and df["Campaña"].notna().any():
            campanas = sorted(df["Campaña"].dropna().unique())
            campanas_sel = st.multiselect("Campaña", campanas, default=campanas)
            if campanas_sel:
                df = df[df["Campaña"].isin(campanas_sel)]

    with f2:
        if "Técnico/a" in df.columns and df["Técnico/a"].notna().any():
            tecnicos = sorted(df["Técnico/a"].dropna().unique())
            tecnicos_sel = st.multiselect("Técnico/a", tecnicos, default=tecnicos)
            if tecnicos_sel:
                df = df[df["Técnico/a"].isin(tecnicos_sel)]

    with f3:
        if "Género" in df.columns and df["Género"].notna().any():
            generos = sorted(df["Género"].dropna().unique())
            generos_sel = st.multiselect(
                "Cultivo / Género",
                generos,
                default=[],
                placeholder="Todos los cultivos",
            )
            if generos_sel:
                df = df[df["Género"].isin(generos_sel)]

    with f4:
        if "Tipo de agricultor" in df.columns and df["Tipo de agricultor"].notna().any():
            tipos = sorted(df["Tipo de agricultor"].dropna().unique())
            tipos_sel = st.multiselect(
                "Tipo de agricultor",
                tipos,
                default=[],
                placeholder="Todos los tipos",
            )
            if tipos_sel:
                df = df[df["Tipo de agricultor"].isin(tipos_sel)]

st.caption(f"Registros filtrados: **{len(df):,}**")

if not HAS_PLOTLY:
    st.caption("Mostrando gráficos nativos de Streamlit (añade `plotly` a `requirements.txt` y reinicia la app para usar Plotly).")

if df.empty:
    st.warning("No hay registros que coincidan con los filtros seleccionados.")
    st.stop()


# ---------------------------------------------------------
# 5. KPIs Y GRÁFICOS
# ---------------------------------------------------------
total_visitas = len(df)
agricultores_unicos = df["Agricultor/a"].nunique() if "Agricultor/a" in df.columns else 0
tecnicos_activos = df["Técnico/a"].nunique() if "Técnico/a" in df.columns else 0
pct_fotos = (df["Tiene foto"].mean() * 100) if total_visitas > 0 else 0.0

k1, k2, k3, k4 = st.columns(4)
k1.metric("Total Visitas", f"{total_visitas:,}")
k2.metric("Agricultores Visitados", f"{agricultores_unicos:,}")
k3.metric("Técnicos Activos", f"{tecnicos_activos:,}")
k4.metric("Visitas con Foto", f"{pct_fotos:.1f}%")

st.divider()

col_izq, col_der = st.columns(2)

with col_izq:
    if "Técnico/a" in df.columns:
        visitas_tec = (
            df.groupby("Técnico/a", as_index=False)
            .agg(Visitas=("Id", "count"), Agricultores=("Agricultor/a", "nunique"))
            .sort_values("Visitas", ascending=False)
        )
        if HAS_PLOTLY:
            fig_tec = px.bar(
                visitas_tec,
                x="Técnico/a",
                y="Visitas",
                hover_data=["Agricultores"],
                title="Visitas por técnico/a",
                text_auto=True,
            )
            st.plotly_chart(fig_tec, use_container_width=True)
        else:
            st.subheader("Visitas por técnico/a")
            st.bar_chart(visitas_tec.set_index("Técnico/a")["Visitas"])

with col_der:
    if "Género" in df.columns:
        visitas_gen = (
            df.groupby("Género", as_index=False)
            .size()
            .rename(columns={"size": "Visitas"})
            .sort_values("Visitas", ascending=False)
        )
        if HAS_PLOTLY:
            fig_gen = px.bar(
                visitas_gen,
                x="Género",
                y="Visitas",
                title="Visitas por cultivo / género",
                text_auto=True,
            )
            st.plotly_chart(fig_gen, use_container_width=True)
        else:
            st.subheader("Visitas por cultivo / género")
            st.bar_chart(visitas_gen.set_index("Género")["Visitas"])

df["Mes"] = df["Fecha"].dt.to_period("M").astype(str)
visitas_mes = (
    df.groupby("Mes", as_index=False)
    .size()
    .rename(columns={"size": "Visitas"})
    .sort_values("Mes")
)

if HAS_PLOTLY:
    fig_mes = px.line(
        visitas_mes,
        x="Mes",
        y="Visitas",
        markers=True,
        title="Evolución mensual de visitas",
    )
    st.plotly_chart(fig_mes, use_container_width=True)
else:
    st.subheader("Evolución mensual de visitas")
    st.line_chart(visitas_mes.set_index("Mes")["Visitas"])

# Tabla de detalle
cols_visibles = [
    c for c in [
        "Fecha", "hora registro", "Campaña", "Semana", "Técnico/a",
        "Agricultor/a", "Tipo de agricultor", "Género",
        "Próxima visita", "Anotaciones", "Tiene foto", "Tiene archivo"
    ]
    if c in df.columns
]

with st.expander("Ver detalle de visitas filtradas", expanded=True):
    st.dataframe(df[cols_visibles], use_container_width=True, hide_index=True)
