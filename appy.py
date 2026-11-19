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
    page_icon="📊",
    layout="wide",
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
        "secreto": "gcp_service_account",  # cambia a "gcp_service_account" si usas la misma
    },
}


# ---------------------------------------------------------
# 2. CONEXIÓN Y CARGA CON CACHÉ
# ---------------------------------------------------------
@st.cache_resource
def obtener_credenciales(secreto_key: str = "gcp_service_account"):
    if secreto_key not in st.secrets:
        raise KeyError(f"No se encontró `st.secrets['{secreto_key}']`.")
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

    if "FECHA" in df.columns:
        df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce", dayfirst=True)
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
    ]:
        if col in df.columns:
            df[col] = limpiar_numero_es(df[col])

    for col in ["LINEA PRODUCCIÓN", "RESPONSABLE LINEA", "GÉNERO", "CLIENTE", "Campaña"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            df.loc[df[col].isin(["nan", "None", ""]), col] = None

    return df


# ---------------------------------------------------------
# 4. BARRA LATERAL Y CARGA DE FUENTE
# ---------------------------------------------------------
st.sidebar.title("⚙️ Configuración")

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

if st.sidebar.button("🔄 Actualizar datos", use_container_width=True):
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
        df_raw = cargar_datos_excel_local(ruta_local, cfg["pestana"])
except HttpError as e:
    st.error(f"Error de Google Drive ({e.resp.status}) en '{fuente_seleccionada}': {e}")
    st.info(
        "Revisa el ID, que el archivo esté compartido con "
        "`firebase-app-hosting-compute@studio-5810984322-26d93.iam.gserviceaccount.com` "
        "y que **Google Drive API** esté habilitada."
    )
    st.stop()
except Exception as e:
    st.error(f"No se pudo cargar '{fuente_seleccionada}' ({cfg['pestana']}): {e}")
    st.info(
        "Revisa el ID del archivo, el nombre de la pestaña, que esté compartido con "
        "`firebase-app-hosting-compute@studio-5810984322-26d93.iam.gserviceaccount.com` "
        "y que `st.secrets['gcp_service_account']` esté configurado."
    )
    st.stop()

if df_raw.empty:
    st.warning("La hoja seleccionada no tiene datos.")
    st.stop()


# ---------------------------------------------------------
# 5. VISTA A: REGISTRO DE VISITAS
# ---------------------------------------------------------
if cfg["tipo"] == "visitas":
    df = limpiar_visitas(df_raw)
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
                    "Cultivo / Género", generos, default=[], placeholder="Todos los cultivos"
                )
                if generos_sel:
                    df = df[df["Género"].isin(generos_sel)]

        with f4:
            if "Tipo de agricultor" in df.columns and df["Tipo de agricultor"].notna().any():
                tipos = sorted(df["Tipo de agricultor"].dropna().unique())
                tipos_sel = st.multiselect(
                    "Tipo de agricultor", tipos, default=[], placeholder="Todos los tipos"
                )
                if tipos_sel:
                    df = df[df["Tipo de agricultor"].isin(tipos_sel)]

    st.caption(f"Registros filtrados: **{len(df):,}**")
    if not HAS_PLOTLY:
        st.caption("Mostrando gráficos nativos de Streamlit (añade `plotly` a `requirements.txt` para usar Plotly).")
    if df.empty:
        st.warning("No hay registros que coincidan con los filtros seleccionados.")
        st.stop()

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

    df["Mes"] = df["Fecha"].dt.to_period("M").astype(str)
    visitas_mes = (
        df.groupby("Mes", as_index=False)
        .size()
        .rename(columns={"size": "Visitas"})
        .sort_values("Mes")
    )

    col_izq, col_der = st.columns(2)
    with col_izq:
        if "Técnico/a" in df.columns:
            conteo_col = "Id" if "Id" in df.columns else df.columns[0]
            visitas_tec = (
                df.groupby("Técnico/a", as_index=False)
                .agg(Visitas=(conteo_col, "count"))
                .sort_values("Visitas", ascending=False)
            )
            if HAS_PLOTLY:
                st.plotly_chart(
                    px.bar(
                        visitas_tec,
                        x="Técnico/a",
                        y="Visitas",
                        title="Visitas por técnico/a",
                        text_auto=True,
                    ),
                    use_container_width=True,
                )
            else:
                st.subheader("Visitas por técnico/a")
                st.bar_chart(visitas_tec.set_index("Técnico/a")["Visitas"])

    with col_der:
        if HAS_PLOTLY:
            st.plotly_chart(
                px.bar(
                    visitas_mes,
                    x="Mes",
                    y="Visitas",
                    title="Visitas por mes",
                    text_auto=True,
                ),
                use_container_width=True,
            )
        else:
            st.subheader("Visitas por mes")
            st.bar_chart(visitas_mes.set_index("Mes")["Visitas"])

    df["Semana_inicio"] = (
        df["Fecha"].dt.to_period("W").apply(lambda r: r.start_time.strftime("%Y-%m-%d"))
    )
    df_tec_sem = df.dropna(subset=["Técnico/a"]) if "Técnico/a" in df.columns else df.copy()
    visitas_sem_tec = (
        df_tec_sem.groupby(["Semana_inicio", "Técnico/a"], as_index=False)
        .size()
        .rename(columns={"size": "Visitas"})
        .sort_values("Semana_inicio")
    )

    if HAS_PLOTLY:
        st.plotly_chart(
            px.line(
                visitas_sem_tec,
                x="Semana_inicio",
                y="Visitas",
                color="Técnico/a",
                markers=True,
                title="Evolución semanal de visitas por técnico/a",
                labels={"Semana_inicio": "Semana (inicio)"},
            ),
            use_container_width=True,
        )
    else:
        st.subheader("Evolución semanal de visitas por técnico/a")
        piv_sem = visitas_sem_tec.pivot(
            index="Semana_inicio", columns="Técnico/a", values="Visitas"
        ).fillna(0)
        st.line_chart(piv_sem)

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


# ---------------------------------------------------------
# 6. VISTA B: PRODUCCIÓN Y CONFECCIÓN (HOJA1)
# ---------------------------------------------------------
else:
    df = limpiar_produccion(df_raw)
    st.title("🏭 Cuadro de Mandos — Producción y Confección")

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
            if "LINEA PRODUCCIÓN" in df.columns and df["LINEA PRODUCCIÓN"].notna().any():
                lineas = sorted(df["LINEA PRODUCCIÓN"].dropna().unique())
                lineas_sel = st.multiselect("Línea de producción", lineas, default=[])
                if lineas_sel:
                    df = df[df["LINEA PRODUCCIÓN"].isin(lineas_sel)]

        with f3:
            if "GÉNERO" in df.columns and df["GÉNERO"].notna().any():
                generos = sorted(df["GÉNERO"].dropna().unique())
                generos_sel = st.multiselect("Género", generos, default=[])
                if generos_sel:
                    df = df[df["GÉNERO"].isin(generos_sel)]

        with f4:
            if "CLIENTE" in df.columns and df["CLIENTE"].notna().any():
                clientes = sorted(df["CLIENTE"].dropna().unique())
                clientes_sel = st.multiselect("Cliente", clientes, default=[])
                if clientes_sel:
                    df = df[df["CLIENTE"].isin(clientes_sel)]

    st.caption(f"Registros filtrados: **{len(df):,}**")
    if not HAS_PLOTLY:
        st.caption("Mostrando gráficos nativos de Streamlit (añade `plotly` a `requirements.txt` para usar Plotly).")
    if df.empty:
        st.warning("No hay registros que coincidan con los filtros seleccionados.")
        st.stop()

    kg_mov = df["KG MOVIDOS"].sum() if "KG MOVIDOS" in df.columns else 0.0
    kg_conf = df["TOTAL KG CONFECCIONADOS"].sum() if "TOTAL KG CONFECCIONADOS" in df.columns else 0.0
    imp_tot = df["TOTAL IMPORTES"].sum() if "TOTAL IMPORTES" in df.columns else 0.0
    coste_kg = (imp_tot / kg_conf) if kg_conf > 0 else 0.0

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Kg Movidos", f"{kg_mov:,.0f} kg")
    k2.metric("Total Kg Confeccionados", f"{kg_conf:,.0f} kg")
    k3.metric("Total Importes", f"{imp_tot:,.2f} €")
    k4.metric("Coste Medio (€/kg conf.)", f"{coste_kg:,.3f} €/kg")

    st.divider()

    col_izq, col_der = st.columns(2)
    with col_izq:
        if "LINEA PRODUCCIÓN" in df.columns and "TOTAL KG CONFECCIONADOS" in df.columns:
            por_linea = (
                df.groupby("LINEA PRODUCCIÓN", as_index=False)["TOTAL KG CONFECCIONADOS"]
                .sum()
                .sort_values("TOTAL KG CONFECCIONADOS", ascending=False)
            )
            if HAS_PLOTLY:
                st.plotly_chart(
                    px.bar(
                        por_linea,
                        x="LINEA PRODUCCIÓN",
                        y="TOTAL KG CONFECCIONADOS",
                        title="Kg confeccionados por línea de producción",
                        text_auto=".2s",
                    ),
                    use_container_width=True,
                )
            else:
                st.subheader("Kg confeccionados por línea de producción")
                st.bar_chart(por_linea.set_index("LINEA PRODUCCIÓN")["TOTAL KG CONFECCIONADOS"])

    with col_der:
        if "FECHA" in df.columns and "TOTAL IMPORTES" in df.columns:
            df["Mes"] = df["FECHA"].dt.to_period("M").astype(str)
            por_mes = (
                df.groupby("Mes", as_index=False)["TOTAL IMPORTES"]
                .sum()
                .sort_values("Mes")
            )
            if HAS_PLOTLY:
                st.plotly_chart(
                    px.bar(
                        por_mes,
                        x="Mes",
                        y="TOTAL IMPORTES",
                        title="Importes por mes (€)",
                        text_auto=".2s",
                    ),
                    use_container_width=True,
                )
            else:
                st.subheader("Importes por mes (€)")
                st.bar_chart(por_mes.set_index("Mes")["TOTAL IMPORTES"])

    if (
        "FECHA" in df.columns
        and "LINEA PRODUCCIÓN" in df.columns
        and "TOTAL KG CONFECCIONADOS" in df.columns
    ):
        df["Semana_inicio"] = (
            df["FECHA"].dt.to_period("W").apply(lambda r: r.start_time.strftime("%Y-%m-%d"))
        )
        sem_linea = (
            df.dropna(subset=["LINEA PRODUCCIÓN"])
            .groupby(["Semana_inicio", "LINEA PRODUCCIÓN"], as_index=False)["TOTAL KG CONFECCIONADOS"]
            .sum()
            .sort_values("Semana_inicio")
        )
        if HAS_PLOTLY:
            st.plotly_chart(
                px.line(
                    sem_linea,
                    x="Semana_inicio",
                    y="TOTAL KG CONFECCIONADOS",
                    color="LINEA PRODUCCIÓN",
                    markers=True,
                    title="Evolución semanal de kg confeccionados por línea",
                    labels={"Semana_inicio": "Semana (inicio)"},
                ),
                use_container_width=True,
            )
        else:
            st.subheader("Evolución semanal de kg confeccionados por línea")
            piv_l = sem_linea.pivot(
                index="Semana_inicio", columns="LINEA PRODUCCIÓN", values="TOTAL KG CONFECCIONADOS"
            ).fillna(0)
            st.line_chart(piv_l)

    with st.expander("Ver detalle de producción filtrada", expanded=True):
        st.dataframe(df, use_container_width=True, hide_index=True)
