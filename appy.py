import io
import os
import streamlit as st
import pandas as pd
import gspread
from gspread.exceptions import APIError
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload

try:
    import plotly.express as px
    HAS_PLOTLY = True
except ImportError:
    HAS_PLOTLY = False

# ---------------------------------------------------------
# 1. CONFIGURACIÓN GENERAL Y FUENTES DE DATOS
# ---------------------------------------------------------
st.set_page_config(
    page_title="Cuadro de Mandos",
    layout="wide",
)

st.markdown(
    """
    <style>
        .block-container {
            padding-top: 1.8rem;
            padding-bottom: 2rem;
        }
        [data-testid="stMetric"] {
            background-color: #f8f9fb;
            border: 1px solid #e6e9ef;
            padding: 14px 18px;
            border-radius: 8px;
        }
        [data-testid="stMetricLabel"] {
            font-size: 0.85rem;
            font-weight: 600;
            color: #4a5568;
            text-transform: uppercase;
            letter-spacing: 0.03em;
        }
        [data-testid="stMetricValue"] {
            font-size: 1.55rem;
            font-weight: 700;
            color: #1a202c;
        }
        h1 {
            font-weight: 700;
            letter-spacing: -0.02em;
            color: #111827;
            margin-bottom: 0.25rem;
        }
        h2, h3 {
            font-weight: 600;
            color: #1f2937;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

FUENTES = {
    "Registro visitas": {
        "file_id": "1Sm-0JmTghLfQczZ3k-W57JNZRbH10waNQj6HVOXZdgk",
        "pestana": "Registro visitas",
        "formato": "gsheet",
        "tipo": "visitas",
        "archivo_local": "Tecnicos.xlsx",
        "secreto": "gcp_service_account",
    },
    "Producción y Confección (Hoja1)": {
        "file_id": "11snS_ZAfyZY7ZMnTOe2QwvbXNYGTN8hl",
        "pestana": "Hoja1",
        "formato": "xlsx_drive",
        "tipo": "produccion",
        "archivo_local": "prueba (1).xlsx",
        "secreto": "gcp_service_account_prod",
    },
}


# ---------------------------------------------------------
# 2. CONEXIÓN Y CARGA CON CACHÉ
# ---------------------------------------------------------
@st.cache_resource
def obtener_credenciales(secreto_key: str = "gcp_service_account"):
    if secreto_key not in st.secrets:
        raise KeyError(f"No se encontró st.secrets['{secreto_key}'].")
    return Credentials.from_service_account_info(
        st.secrets[secreto_key],
        scopes=SCOPES,
    )


@st.cache_resource
def conectar_gsheets(secreto_key: str = "gcp_service_account"):
    return gspread.authorize(obtener_credenciales(secreto_key))


@st.cache_resource
def conectar_drive(secreto_key: str = "gcp_service_account"):
    return build(
        "drive",
        "v3",
        credentials=obtener_credenciales(secreto_key),
        cache_discovery=False,
    )


def descargar_xlsx_bytes(file_id: str, secreto_key: str = "gcp_service_account") -> io.BytesIO:
    drive_service = conectar_drive(secreto_key)
    request = drive_service.files().get_media(fileId=file_id, supportsAllDrives=True)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.seek(0)
    return fh


@st.cache_data(ttl=600)
def cargar_hoja_inteligente(
    file_id: str,
    pestana: str,
    formato_preferido: str = "auto",
    secreto_key: str = "gcp_service_account",
) -> pd.DataFrame:
    if not file_id or file_id.startswith("ID_"):
        raise ValueError("Falta configurar el ID real del archivo en FUENTES.")

    if formato_preferido != "xlsx_drive":
        try:
            cliente = conectar_gsheets(secreto_key)
            ws = cliente.open_by_key(file_id).worksheet(pestana)
            return pd.DataFrame(ws.get_all_records())
        except APIError as err:
            msg = str(err)
            if "not be an Office file" not in msg and "Use the Drive API" not in msg:
                raise

    fh = descargar_xlsx_bytes(file_id, secreto_key=secreto_key)
    return pd.read_excel(fh, sheet_name=pestana)


@st.cache_data(ttl=600)
def cargar_datos_excel_local(ruta_excel: str, pestana: str) -> pd.DataFrame:
    return pd.read_excel(ruta_excel, sheet_name=pestana)


# ---------------------------------------------------------
# 3. LIMPIEZA Y PREPARACIÓN DE DATOS
# ---------------------------------------------------------
def limpiar_numero_es(serie: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(serie):
        return serie.fillna(0.0)
    limpio = (
        serie.astype(str)
        .str.strip()
        .str.replace("\u00a0", "", regex=False)
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
    )
    return pd.to_numeric(limpio, errors="coerce").fillna(0.0)


def buscar_columna(df: pd.DataFrame, candidatos: list[str]) -> str | None:
    columnas_norm = {c.lower().strip(): c for c in df.columns}
    for cand in candidatos:
        clave = cand.lower().strip()
        if clave in columnas_norm:
            return columnas_norm[clave]
    for cand in candidatos:
        clave = cand.lower().strip()
        for col_norm, col_orig in columnas_norm.items():
            if clave in col_norm:
                return col_orig
    return None


def limpiar_visitas(df_raw: pd.DataFrame) -> pd.DataFrame:
    df = df_raw.copy()
    df.columns = [str(c).strip() for c in df.columns]

    if "Id" in df.columns and "Fecha" in df.columns:
        df = df.dropna(subset=["Id", "Fecha"], how="all")
        df = df[df["Id"].astype(str).str.strip() != ""]

    if "Fecha" in df.columns:
        df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
        df = df.dropna(subset=["Fecha"])
        df = df.sort_values("Fecha", ascending=False)

    if "Próxima visita" in df.columns:
        df["Próxima visita"] = pd.to_datetime(df["Próxima visita"], errors="coerce")

    if "Semana" in df.columns:
        df["Semana"] = pd.to_numeric(df["Semana"], errors="coerce")

    for col in ["Técnico/a", "Agricultor/a", "Tipo de agricultor", "Género", "Campaña"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            df.loc[df[col].isin(["nan", "None", ""]), col] = None

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
    return df


def limpiar_produccion(df_raw: pd.DataFrame) -> pd.DataFrame:
    df = df_raw.copy()
    df.columns = [str(c).strip() for c in df.columns]

    col_fecha = buscar_columna(df, ["FECHA", "Fecha"])
    if col_fecha:
        df["FECHA"] = pd.to_datetime(df[col_fecha], errors="coerce", dayfirst=True)
        df = df.dropna(subset=["FECHA"])
        df = df.sort_values("FECHA", ascending=False)

    for col in [
        "KG MOVIDOS",
        "KG CONFECCIONADOS 1",
        "KG CONFECCIONADOS 2",
        "KG CONFECCIONADOS 3",
        "TOTAL KG CONFECCIONADOS",
        "TOTAL IMPORTE",
        "TOTAL IMPORTE EXTRA",
        "TOTAL IMPORTES",
        "COSTE MO",
        "Coste Mo Estimado",
        "PERSONAL",
        "Personal",
        "OPERARIOS",
        "Operarios",
        "TRABAJADORES",
        "Trabajadores",
        "N PERSONAL",
        "Nº PERSONAL",
        "NUM PERSONAL",
    ]:
        if col in df.columns:
            df[col] = limpiar_numero_es(df[col])

    for col in ["LINEA PRODUCCIÓN", "RESPONSABLE LINEA", "GÉNERO", "PRODUCTO", "ARTÍCULO", "CLIENTE", "Campaña"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            df.loc[df[col].isin(["nan", "None", ""]), col] = None

    return df


# ---------------------------------------------------------
# 4. BARRA LATERAL Y CARGA DE FUENTE
# ---------------------------------------------------------
st.sidebar.title("Configuración")

fuente_seleccionada = st.sidebar.selectbox(
    "Selecciona hoja",
    options=list(FUENTES.keys()),
    index=0,
)
cfg = FUENTES[fuente_seleccionada]

origen = st.sidebar.radio(
    "Origen de datos",
    options=["Google Drive / Sheets", "Excel local"],
    index=0,
)

if st.sidebar.button("Actualizar datos", use_container_width=True):
    st.cache_data.clear()
    st.rerun()

secreto_actual = cfg.get("secreto", "gcp_service_account")

try:
    if origen == "Google Drive / Sheets":
        df_raw = cargar_hoja_inteligente(
            cfg["file_id"],
            cfg["pestana"],
            formato_preferido=cfg.get("formato", "auto"),
            secreto_key=secreto_actual,
        )
    else:
        ruta_local = cfg["archivo_local"]
        if not os.path.exists(ruta_local) and os.path.exists("Tecnicos.xlsx"):
            ruta_local = "Tecnicos.xlsx"
        df_
