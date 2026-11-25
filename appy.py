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
    initial_sidebar_state="auto",
)

st.markdown(
    """
    <style>
        .block-container {
            padding-top: 1.5rem;
            padding-bottom: 2rem;
            padding-left: 1.5rem;
            padding-right: 1.5rem;
        }
        [data-testid="stMetric"] {
            background-color: #f8f9fb;
            border: 1px solid #e6e9ef;
            padding: 12px 16px;
            border-radius: 8px;
        }
        [data-testid="stMetricLabel"] {
            font-size: 0.8rem;
            font-weight: 600;
            color: #4a5568;
            text-transform: uppercase;
            letter-spacing: 0.03em;
        }
        [data-testid="stMetricValue"] {
            font-size: 1.45rem;
            font-weight: 700;
            color: #1a202c;
        }
        h1 {
            font-weight: 700;
            letter-spacing: -0.02em;
            color: #111827;
            margin-bottom: 0.2rem;
        }
        h2, h3 {
            font-weight: 600;
            color: #1f2937;
        }
        @media (max-width: 768px) {
            .block-container {
                padding-top: 0.8rem;
                padding-bottom: 1.2rem;
                padding-left: 0.75rem;
                padding-right: 0.75rem;
            }
            h1 {
                font-size: 1.45rem !important;
                line-height: 1.2 !important;
            }
            h2 {
                font-size: 1.15rem !important;
            }
            h3 {
                font-size: 1.05rem !important;
            }
            [data-testid="stMetric"] {
                padding: 10px 12px;
            }
            [data-testid="stMetricLabel"] {
                font-size: 0.72rem;
            }
            [data-testid="stMetricValue"] {
                font-size: 1.25rem;
            }
            .stButton > button {
                min-height: 2.6rem;
            }
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
    "Previsiones": {
        "file_id": "1O3cZG22pczeI4OgDHvXe0yF9nwTNc_u4f3L9k6t0Ch8",
        "pestana": "Entregas",
        "formato": "gsheet",
        "tipo": "entregas",
        "archivo_local": "DB_CalidadACL.xlsx",
        "secreto": "gcp_service_account",
    },
}


def render_plotly_responsive(fig, height: int = 320, rotar_x: bool = True) -> None:
    fig.update_layout(
        height=height,
        margin=dict(l=10, r=10, t=42, b=15),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="left",
            x=0,
            title_text="",
        ),
        font=dict(size=12),
    )
    if rotar_x:
        fig.update_xaxes(tickangle=-35, automargin=True)
    fig.update_yaxes(automargin=True)
    st.plotly_chart(
        fig,
        use_container_width=True,
        config={"displayModeBar": False, "responsive": True},
    )


def exigir_contrasena() -> None:
    if st.session_state.get("autenticado", False):
        return

    clave_esperada = st.secrets.get("app_auth", {}).get("password")
    if not clave_esperada:
        st.error("No está configurada `st.secrets['app_auth']['password']`.")
        st.stop()

    st.title("Acceso al Cuadro de Mandos")
    st.caption("Introduce la contraseña para continuar")

    with st.form("login_form", clear_on_submit=False):
        clave_introducida = st.text_input("Contraseña", type="password")
        enviado = st.form_submit_button("Entrar", use_container_width=True)

    if enviado:
        if clave_introducida == clave_esperada:
            st.session_state["autenticado"] = True
            st.rerun()
        else:
            st.error("Contraseña incorrecta.")

    st.stop()


exigir_contrasena()


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


def calcular_coste_medio_agrupado(
    df: pd.DataFrame,
    grupo_cols: str | list[str],
    col_importe: str = "TOTAL IMPORTES",
    col_kg: str = "TOTAL KG CONFECCIONADOS",
    nombre_metrica: str = "Coste Medio MO (€/kg)",
) -> pd.DataFrame:
    if isinstance(grupo_cols, str):
        grupo_cols = [grupo_cols]
    df_sub = df.dropna(subset=grupo_cols).copy()
    if df_sub.empty:
        return pd.DataFrame(columns=grupo_cols + ["Importes", "Kg Confeccionados", nombre_metrica])

    agg = (
        df_sub.groupby(grupo_cols, as_index=False)
        .agg(
            Importes=(col_importe, "sum"),
            Kg_Confeccionados=(col_kg, "sum"),
        )
    )
    agg[nombre_metrica] = agg.apply(
        lambda r: (r["Importes"] / r["Kg_Confeccionados"]) if r["Kg_Confeccionados"] > 0 else 0.0,
        axis=1,
    )
    return agg


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

    for col in ["Técnico/a", "Agricultor/a", "Tipo de agricultor", "Género", "Campaña", "Anotaciones"]:
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

    col_semana = buscar_columna(df, ["SEMANA", "Semana", "SEM"])
    if col_semana:
        df["Semana"] = pd.to_numeric(df[col_semana], errors="coerce")

    col_fecha = buscar_columna(df, ["FECHA", "Fecha"])
    if col_fecha:
        df["FECHA"] = pd.to_datetime(df[col_fecha], errors="coerce", dayfirst=True)
        df = df.dropna(subset=["FECHA"])
        df = df.sort_values("FECHA", ascending=False)

    for col in [
        "KG CONFECCIONADOS 1",
        "KG CONFECCIONADOS 2",
        "KG CONFECCIONADOS 3",
        "TOTAL KG CONFECCIONADOS",
        "TOTAL IMPORTE",
        "TOTAL IMPORTE EXTRA",
        "TOTAL IMPORTES",
        "COSTE MO",
        "Coste Mo Estimado",
        "DESVIACION",
        "DESVIACIÓN",
        "DESVIACION MO",
        "DESVIACIÓN MO",
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

    for col in [
        "FAMILIA",
        "Familia",
        "CONFECCION",
        "CONFECCIÓN",
        "LINEA PRODUCCIÓN",
        "RESPONSABLE LINEA",
        "JEFA DE LINEA",
        "JEFE DE LINEA",
        "GÉNERO",
        "PRODUCTO",
        "ARTÍCULO",
        "CLIENTE",
        "Campaña",
    ]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            df.loc[df[col].isin(["nan", "None", ""]), col] = None

    return df


def limpiar_entregas(df_raw: pd.DataFrame) -> pd.DataFrame:
    df = df_raw.copy()
    df.columns = [str(c).strip() for c in df.columns]

    if "ID_Entrega" in df.columns:
        df = df.dropna(subset=["ID_Entrega"])
        df = df[df["ID_Entrega"].astype(str).str.strip() != ""]

    col_fecha = buscar_columna(df, ["Fecha_", "FECHA", "Fecha"])
    if col_fecha:
        df["Fecha_"] = pd.to_datetime(df[col_fecha], errors="coerce", dayfirst=True)
        df = df.dropna(subset=["Fecha_"])
        df = df.sort_values("Fecha_", ascending=False)

    if "Semana" in df.columns:
        df["Semana"] = pd.to_numeric(df["Semana"], errors="coerce")

    for col in ["Kg", "Kg_Entregados", "Dif_Kg"]:
        if col in df.columns:
            df[col] = limpiar_numero_es(df[col])

    for col in ["Gestion", "Campaña", "Responsable", "Nombre", "Producto", "Calibre", "Ubicacion"]:
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

if st.sidebar.button("Actualizar datos", use_container_width=True):
    st.cache_data.clear()
    st.rerun()

if st.sidebar.button("Cerrar sesión", use_container_width=True):
    st.session_state["autenticado"] = False
    st.rerun()

secreto_actual = cfg.get("secreto", "gcp_service_account")
correo_cuenta_servicio = (
    st.secrets.get(secreto_actual, {}).get("client_email", secreto_actual)
    if secreto_actual in st.secrets
    else secreto_actual
)

try:
    df_raw = cargar_hoja_inteligente(
        cfg["file_id"],
        cfg["pestana"],
        formato_preferido=cfg.get("formato", "auto"),
        secreto_key=secreto_actual,
    )
except HttpError as e:
    st.error(f"Error de Google Drive ({e.resp.status}) en '{fuente_seleccionada}': {e}")
    st.info(
        f"Revisa el ID del archivo, que esté compartido como Lector con `{correo_cuenta_servicio}` "
        "y que Google Drive API esté habilitada."
    )
    st.stop()
except Exception as e:
    st.error(f"No se pudo cargar '{fuente_seleccionada}' ({cfg['pestana']}): {e}")
    st.info(
        f"Revisa el ID del archivo, el nombre de la pestaña, que esté compartido con "
        f"`{correo_cuenta_servicio}` y que `st.secrets['{secreto_actual}']` esté configurado."
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
    st.title("Cuadro de Mandos — Registro de Visitas")
    st.caption("Seguimiento comercial y técnico en campo")

    with st.expander("Filtros", expanded=False):
        f1, f2 = st.columns(2)

        with f1:
            if "Campaña" in df.columns and df["Campaña"].notna().any():
                if "Fecha" in df.columns:
                    campanas = (
                        df.dropna(subset=["Campaña"])
                        .groupby("Campaña")["Fecha"]
                        .max()
                        .sort_values(ascending=True)
                        .index.tolist()
                    )
                else:
                    campanas = sorted(df["Campaña"].dropna().unique())
                ultima_campana = [campanas[-1]] if campanas else []
                campanas_sel = st.multiselect("Campaña", campanas, default=ultima_campana)
                if campanas_sel:
                    df = df[df["Campaña"].isin(campanas_sel)]

        with f2:
            if "Técnico/a" in df.columns and df["Técnico/a"].notna().any():
                tecnicos = sorted(df["Técnico/a"].dropna().unique())
                tecnicos_sel = st.multiselect("Técnico/a", tecnicos, default=tecnicos)
                if tecnicos_sel:
                    df = df[df["Técnico/a"].isin(tecnicos_sel)]

        f3, f4 = st.columns(2)
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

    st.caption(f"Registros filtrados: {len(df):,}")
    if df.empty:
        st.warning("No hay registros que coincidan con los filtros seleccionados.")
        st.stop()

    if "Anotaciones" in df.columns and df["Anotaciones"].notna().any():
        mask_precio = df["Anotaciones"].str.contains(r"\bpreci\w*", case=False, na=False)
        df_alertas_precio = df[mask_precio]
        if not df_alertas_precio.empty:
            st.warning(
                f"Alerta: se han detectado {len(df_alertas_precio)} visita(s) con menciones a 'precio' en Anotaciones."
            )
            cols_alerta = [
                c for c in ["Fecha", "Agricultor/a", "Técnico/a", "Género", "Anotaciones"]
                if c in df_alertas_precio.columns
            ]
            with st.expander("Ver visitas con mención a 'precio'", expanded=False):
                st.dataframe(df_alertas_precio[cols_alerta], use_container_width=True, hide_index=True)

    total_visitas = len(df)
    agricultores_unicos = df["Agricultor/a"].nunique() if "Agricultor/a" in df.columns else 0
    tecnicos_activos = df["Técnico/a"].nunique() if "Técnico/a" in df.columns else 0

    k1, k2, k3 = st.columns(3)
    k1.metric("Total Visitas", f"{total_visitas:,}")
    k2.metric("Agricultores Visitados", f"{agricultores_unicos:,}")
    k3.metric("Técnicos Activos", f"{tecnicos_activos:,}")

    # -----------------------------------------------------
            # ALERTAS: AGRICULTORES SIN VISITA Y > 30 DÍAS (1 MES)
            # -----------------------------------------------------
            if {"Agricultor/a", "Fecha"}.issubset(df.columns):
                df_valid_vis = df.dropna(subset=["Agricultor/a", "Fecha"]).copy()
                if not df_valid_vis.empty:
                    fecha_ref = df_valid_vis["Fecha"].max()

                    resumen_agr = (
                        df_valid_vis.groupby("Agricultor/a", as_index=False)
                        .agg(
                            Visitas=("Fecha", "count"),
                            Ultima_Visita=("Fecha", "max"),
                            Tecnico=(
                                "Técnico/a",
                                lambda s: s.dropna().iloc[-1] if s.notna().any() else "-",
                            ),
                            Genero=(
                                "Género",
                                lambda s: s.dropna().iloc[-1] if s.notna().any() else "-",
                            ),
                        )
                    )
                    resumen_agr["Dias_sin_visita"] = (
                        fecha_ref - resumen_agr["Ultima_Visita"]
                    ).dt.days
                    resumen_agr["Ultima_Visita_str"] = resumen_agr["Ultima_Visita"].dt.strftime(
                        "%Y-%m-%d"
                    )

                    # Agricultores sin ninguna visita registrada (si existe columna catálogo)
                    col_catalogo_agr = buscar_columna(
                        df_raw, ["AGRICULTORES", "LISTA AGRICULTORES"]
                    )
                    sin_visita_df = pd.DataFrame()
                    if col_catalogo_agr and col_catalogo_agr in df_raw.columns:
                        todos_agr = sorted(
                            {
                                str(x).strip()
                                for x in df_raw[col_catalogo_agr].dropna()
                                if str(x).strip() not in ("", "nan", "None")
                            }
                        )
                        visitados_set = set(resumen_agr["Agricultor/a"].dropna())
                        nunca_visitados = sorted(set(todos_agr) - visitados_set)
                        if nunca_visitados:
                            sin_visita_df = pd.DataFrame(
                                {
                                    "Agricultor/a": nunca_visitados,
                                    "Visitas": 0,
                                    "Ultima_Visita_str": "Sin visitas",
                                    "Dias_sin_visita": None,
                                    "Tecnico": "-",
                                    "Genero": "-",
                                }
                            )

                    if not sin_visita_df.empty:
                        st.error(
                            f"Alerta: {len(sin_visita_df)} agricultor(es) no tienen ninguna visita registrada."
                        )
                        with st.expander(
                            "Ver agricultores sin ninguna visita", expanded=False
                        ):
                            st.dataframe(
                                sin_visita_df[
                                    ["Agricultor/a", "Visitas", "Ultima_Visita_str"]
                                ],
                                use_container_width=True,
                                hide_index=True,
                            )

                    # Agricultores con más de 30 días (1 mes) sin visita
                    df_mas_de_un_mes = resumen_agr[
                        resumen_agr["Dias_sin_visita"] > 30
                    ].sort_values("Dias_sin_visita", ascending=False)

                    if not df_mas_de_un_mes.empty:
                        st.warning(
                            f"Alerta: {len(df_mas_de_un_mes)} agricultor(es) llevan más de 30 días (1 mes) sin visita."
                        )

                        import matplotlib
                        matplotlib.use("Agg")
                        import matplotlib.pyplot as plt

                        df_tabla_mes = df_mas_de_un_mes[
                            [
                                "Agricultor/a",
                                "Tecnico",
                                "Genero",
                                "Ultima_Visita_str",
                                "Dias_sin_visita",
                                "Visitas",
                            ]
                        ].rename(
                            columns={
                                "Ultima_Visita_str": "Última visita",
                                "Dias_sin_visita": "Días sin visita",
                                "Tecnico": "Técnico/a",
                                "Genero": "Género",
                            }
                        )

                        with st.expander(
                            "Ver agricultores con más de 1 mes sin visita", expanded=False
                        ):
                            st.dataframe(
                                df_tabla_mes, use_container_width=True, hide_index=True
                            )

                            c_dl1, c_dl2 = st.columns(2)

                            # 1. Descarga CSV (compatible con Excel en español)
                            csv_bytes = df_tabla_mes.to_csv(
                                index=False, sep=";"
                            ).encode("utf-8-sig")
                            with c_dl1:
                                st.download_button(
                                    label="Descargar CSV",
                                    data=csv_bytes,
                                    file_name="agricultores_mas_1_mes_sin_visita.csv",
                                    mime="text/csv",
                                    use_container_width=True,
                                    key="dl_csv_sin_visita",
                                )

                            # 2. Imagen PNG de la tabla para descargar o copiar
                            n_filas = len(df_tabla_mes)
                            fig_h = max(1.8, 0.55 + n_filas * 0.36)
                            fig, ax = plt.subplots(figsize=(9.5, fig_h), dpi=200)
                            ax.axis("off")

                            tabla_mpl = ax.table(
                                cellText=df_tabla_mes.astype(str).values,
                                colLabels=df_tabla_mes.columns,
                                loc="center",
                                cellLoc="left",
                            )
                            tabla_mpl.auto_set_font_size(False)
                            tabla_mpl.set_fontsize(9)
                            tabla_mpl.scale(1, 1.35)

                            for (row, col), cell in tabla_mpl.get_celld().items():
                                cell.set_edgecolor("#e2e8f0")
                                if row == 0:
                                    cell.set_facecolor("#1f2937")
                                    cell.set_text_props(
                                        color="white", weight="bold"
                                    )
                                elif row % 2 == 0:
                                    cell.set_facecolor("#f8f9fb")
                                else:
                                    cell.set_facecolor("#ffffff")

                            buf_img = io.BytesIO()
                            plt.savefig(
                                buf_img, format="png", bbox_inches="tight", dpi=200
                            )
                            plt.close(fig)
                            buf_img.seek(0)
                            img_bytes = buf_img.getvalue()

                            with c_dl2:
                                st.download_button(
                                    label="Descargar imagen (PNG)",
                                    data=img_bytes,
                                    file_name="agricultores_mas_1_mes_sin_visita.png",
                                    mime="image/png",
                                    use_container_width=True,
                                    key="dl_png_sin_visita",
                                )

                            with st.expander(
                                "Vista previa de imagen para copiar", expanded=False
                            ):
                                st.image(img_bytes, use_container_width=True)

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
                fig_tec = px.bar(
                    visitas_tec,
                    x="Técnico/a",
                    y="Visitas",
                    title="Visitas por técnico/a",
                    text_auto=True,
                )
                render_plotly_responsive(fig_tec, height=310)
            else:
                st.subheader("Visitas por técnico/a")
                st.bar_chart(visitas_tec.set_index("Técnico/a")["Visitas"])

    with col_der:
        if HAS_PLOTLY:
            fig_mes = px.bar(
                visitas_mes,
                x="Mes",
                y="Visitas",
                title="Visitas por mes",
                text_auto=True,
            )
            render_plotly_responsive(fig_mes, height=310)
        else:
            st.subheader("Visitas por mes")
            st.bar_chart(visitas_mes.set_index("Mes")["Visitas"])

    df["Dia"] = df["Fecha"].dt.strftime("%Y-%m-%d")
    df_tec_dia = df.dropna(subset=["Técnico/a"]) if "Técnico/a" in df.columns else df.copy()
    visitas_dia_tec = (
        df_tec_dia.groupby(["Dia", "Técnico/a"], as_index=False)
        .size()
        .rename(columns={"size": "Visitas"})
        .sort_values("Dia")
    )

    if HAS_PLOTLY:
        fig_dia = px.bar(
            visitas_dia_tec,
            x="Dia",
            y="Visitas",
            color="Técnico/a",
            barmode="group",
            title="Visitas por día y técnico/a",
            labels={"Dia": "Día"},
        )
        render_plotly_responsive(fig_dia, height=330)
    else:
        st.subheader("Visitas por día y técnico/a")
        piv_dia = visitas_dia_tec.pivot(
            index="Dia", columns="Técnico/a", values="Visitas"
        ).fillna(0)
        st.line_chart(piv_dia)

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
        fig_sem = px.line(
            visitas_sem_tec,
            x="Semana_inicio",
            y="Visitas",
            color="Técnico/a",
            markers=True,
            title="Evolución semanal de visitas por técnico/a",
            labels={"Semana_inicio": "Semana (inicio)"},
        )
        render_plotly_responsive(fig_sem, height=330)
    else:
        st.subheader("Evolución semanal de visitas por técnico/a")
        piv_sem = visitas_sem_tec.pivot(
            index="Semana_inicio", columns="Técnico/a", values="Visitas"
        ).fillna(0)
        st.line_chart(piv_sem)

    cols_visibles = [
        c for c in [
            "Fecha", "Agricultor/a", "Técnico/a", "Género",
            "Anotaciones", "Campaña", "Semana", "Tipo de agricultor",
            "Próxima visita", "Tiene archivo"
        ]
        if c in df.columns
    ]

    # -----------------------------------------------------
    # RANKING DE AGRICULTORES MÁS VISITADOS
    # -----------------------------------------------------
    if "Agricultor/a" in df.columns and df["Agricultor/a"].notna().any():
        st.divider()
        st.subheader("Ranking de agricultores más visitados")

        conteo_col_agr = "Id" if "Id" in df.columns else df.columns[0]
        ranking_agr = (
            df.dropna(subset=["Agricultor/a"])
            .groupby("Agricultor/a", as_index=False)
            .agg(
                Visitas=(conteo_col_agr, "count"),
                Ultima_Visita=("Fecha", "max") if "Fecha" in df.columns else (conteo_col_agr, "count"),
                Tecnico=("Técnico/a", lambda s: s.dropna().iloc[-1] if "Técnico/a" in df.columns and s.notna().any() else "-"),
                Genero=("Género", lambda s: s.dropna().iloc[-1] if "Género" in df.columns and s.notna().any() else "-"),
            )
            .sort_values("Visitas", ascending=False)
            .reset_index(drop=True)
        )
        ranking_agr.insert(0, "Puesto", ranking_agr.index + 1)

        if "Fecha" in df.columns:
            ranking_agr["Última visita"] = pd.to_datetime(ranking_agr["Ultima_Visita"]).dt.strftime("%Y-%m-%d")
        else:
            ranking_agr["Última visita"] = "-"

        top_n = st.slider(
            "Mostrar en gráfico (Top N)",
            min_value=5,
            max_value=min(30, max(5, len(ranking_agr))),
            value=min(10, len(ranking_agr)),
            step=5,
            key="top_n_agricultores",
        )
        top_agr_graf = ranking_agr.head(top_n).sort_values("Visitas", ascending=True)

        if HAS_PLOTLY:
            fig_rank = px.bar(
                top_agr_graf,
                x="Visitas",
                y="Agricultor/a",
                orientation="h",
                text_auto=True,
                title=f"Top {top_n} agricultores más visitados",
                labels={"Agricultor/a": "", "Visitas": "Número de visitas"},
            )
            render_plotly_responsive(fig_rank, height=max(280, len(top_agr_graf) * 32), rotar_x=False)
        else:
            st.bar_chart(top_agr_graf.set_index("Agricultor/a")["Visitas"])

        st.dataframe(
            ranking_agr[["Puesto", "Agricultor/a", "Visitas", "Última visita", "Tecnico", "Genero"]].rename(
                columns={"Tecnico": "Técnico/a", "Genero": "Género"}
            ),
            use_container_width=True,
            hide_index=True,
        )
    with st.expander("Detalle de visitas filtradas", expanded=False):
        busqueda_visitas = st.text_input(
            "Buscar en el detalle",
            placeholder="Escribe agricultor, técnico, cultivo, anotación...",
            key="buscador_detalle_visitas",
        ).strip()

        df_detalle_visitas = df[cols_visibles]
        if busqueda_visitas:
            mask_busqueda_visitas = (
                df_detalle_visitas.astype(str)
                .apply(
                    lambda col: col.str.contains(
                        busqueda_visitas, case=False, na=False, regex=False
                    )
                )
                .any(axis=1)
            )
            df_detalle_visitas = df_detalle_visitas[mask_busqueda_visitas]

        st.caption(f"Mostrando {len(df_detalle_visitas):,} de {len(df):,} registros")
        st.dataframe(df_detalle_visitas, use_container_width=True, hide_index=True)


# ---------------------------------------------------------
# 6. VISTA B: PRODUCCIÓN Y CONFECCIÓN (HOJA1)
# ---------------------------------------------------------
elif cfg["tipo"] == "produccion":
    df = limpiar_produccion(df_raw)
    st.title("Cuadro de Mandos — Producción y Confección")
    st.caption("Coste medio de mano de obra, estimado y desviaciones")

    col_familia = buscar_columna(df, ["FAMILIA", "GÉNERO", "PRODUCTO", "ARTÍCULO", "GENERO"])
    col_confeccion = buscar_columna(df, ["CONFECCION", "CONFECCIÓN", "TIPO CONFECCION", "TIPO CONFECCIÓN"])
    col_jefa = buscar_columna(
        df,
        ["JEFA DE LINEA", "JEFE DE LINEA", "RESPONSABLE LINEA", "RESPONSABLE", "ENCARGADA"],
    )
    col_coste_mo = buscar_columna(df, ["COSTE MO"])
    col_coste_est = buscar_columna(df, ["Coste Mo Estimado", "COSTE MO ESTIMADO"])
    col_desv = buscar_columna(df, ["DESVIACION", "DESVIACIÓN", "DESVIACION MO", "DESVIACIÓN MO"])

    with st.expander("Filtros", expanded=False):
        f1, f2 = st.columns(2)

        with f1:
            if "Campaña" in df.columns and df["Campaña"].notna().any():
                if "FECHA" in df.columns:
                    campanas = (
                        df.dropna(subset=["Campaña"])
                        .groupby("Campaña")["FECHA"]
                        .max()
                        .sort_values(ascending=True)
                        .index.tolist()
                    )
                else:
                    campanas = sorted(df["Campaña"].dropna().unique())
                ultima_campana = [campanas[-1]] if campanas else []
                campanas_sel = st.multiselect("Campaña", campanas, default=ultima_campana)
                if campanas_sel:
                    df = df[df["Campaña"].isin(campanas_sel)]

        with f2:
            if "Semana" in df.columns and df["Semana"].notna().any():
                if "FECHA" in df.columns:
                    semanas = (
                        df.dropna(subset=["Semana"])
                        .assign(Semana_int=lambda d: d["Semana"].astype(int))
                        .groupby("Semana_int")["FECHA"]
                        .max()
                        .sort_values(ascending=True)
                        .index.tolist()
                    )
                else:
                    semanas = sorted(df["Semana"].dropna().astype(int).unique())
                semanas_sel = st.multiselect("Semana", semanas, default=[])
                if semanas_sel:
                    df = df[df["Semana"].isin(semanas_sel)]

        f3, f4 = st.columns(2)
        with f3:
            if col_familia and df[col_familia].notna().any():
                familias = sorted(df[col_familia].dropna().unique())
                familias_sel = st.multiselect("Familia de producto", familias, default=[])
                if familias_sel:
                    df = df[df[col_familia].isin(familias_sel)]

        with f4:
            if col_confeccion and df[col_confeccion].notna().any():
                confecciones = sorted(df[col_confeccion].dropna().unique())
                confecciones_sel = st.multiselect("Confección", confecciones, default=[])
                if confecciones_sel:
                    df = df[df[col_confeccion].isin(confecciones_sel)]

        f5, f6 = st.columns(2)
        with f5:
            if "LINEA PRODUCCIÓN" in df.columns and df["LINEA PRODUCCIÓN"].notna().any():
                lineas = sorted(df["LINEA PRODUCCIÓN"].dropna().unique())
                lineas_sel = st.multiselect("Línea de producción", lineas, default=[])
                if lineas_sel:
                    df = df[df["LINEA PRODUCCIÓN"].isin(lineas_sel)]

        with f6:
            if col_jefa and df[col_jefa].notna().any():
                jefas = sorted(df[col_jefa].dropna().unique())
                jefas_sel = st.multiselect("Jefa de línea", jefas, default=[])
                if jefas_sel:
                    df = df[df[col_jefa].isin(jefas_sel)]

    st.caption(f"Registros filtrados: {len(df):,}")
    if df.empty:
        st.warning("No hay registros que coincidan con los filtros seleccionados.")
        st.stop()

    col_personal = buscar_columna(
        df,
        ["PERSONAL", "OPERARIOS", "TRABAJADORES", "N PERSONAL", "Nº PERSONAL", "NUM PERSONAL"],
    )
    if col_personal:
        if "FECHA" in df.columns:
            personal_diario = df.groupby("FECHA")[col_personal].max()
            personal_val = personal_diario.mean() if not personal_diario.empty else 0.0
            etiqueta_personal = "Media Personal / Día"
        else:
            personal_val = df[col_personal].mean() if not df.empty else 0.0
            etiqueta_personal = "Media Personal"
    else:
        personal_val = df[col_jefa].nunique() if col_jefa else 0.0
        etiqueta_personal = "Personal / Responsables"

    kg_conf = df["TOTAL KG CONFECCIONADOS"].sum() if "TOTAL KG CONFECCIONADOS" in df.columns else 0.0
    imp_tot = df["TOTAL IMPORTES"].sum() if "TOTAL IMPORTES" in df.columns else 0.0
    coste_medio_mo_kg = (imp_tot / kg_conf) if kg_conf > 0 else 0.0

    k1, k2 = st.columns(2)
    k1.metric(etiqueta_personal, f"{personal_val:,.1f}")
    k2.metric("Total Kg Confeccionados", f"{kg_conf:,.0f} kg")

    k3, k4 = st.columns(2)
    k3.metric("Total Importes", f"{imp_tot:,.2f} €")
    k4.metric("Coste Medio MO (€/kg)", f"{coste_medio_mo_kg:,.3f} €/kg")

    st.divider()

    g1, g2 = st.columns(2)

    with g1:
        if col_familia and "TOTAL IMPORTES" in df.columns and "TOTAL KG CONFECCIONADOS" in df.columns:
            df_fam = calcular_coste_medio_agrupado(
                df, col_familia, nombre_metrica="Coste Medio MO (€/kg)"
            ).sort_values("Coste Medio MO (€/kg)", ascending=False)
            if HAS_PLOTLY:
                fig_fam = px.bar(
                    df_fam,
                    x=col_familia,
                    y="Coste Medio MO (€/kg)",
                    title="Coste medio MO por familia (€/kg)",
                    text_auto=".3f",
                )
                render_plotly_responsive(fig_fam, height=320)
            else:
                st.subheader("Coste medio MO por familia (€/kg)")
                st.bar_chart(df_fam.set_index(col_familia)["Coste Medio MO (€/kg)"])
        else:
            st.info("No se encuentran columnas de familia, TOTAL IMPORTES o TOTAL KG CONFECCIONADOS.")

    with g2:
        if col_jefa and "TOTAL IMPORTES" in df.columns and "TOTAL KG CONFECCIONADOS" in df.columns:
            df_jefa = calcular_coste_medio_agrupado(
                df, col_jefa, nombre_metrica="Coste Medio MO (€/kg)"
            ).sort_values("Coste Medio MO (€/kg)", ascending=False)
            if HAS_PLOTLY:
                fig_jefa = px.bar(
                    df_jefa,
                    x=col_jefa,
                    y="Coste Medio MO (€/kg)",
                    title="Coste medio MO por jefa de línea (€/kg)",
                    text_auto=".3f",
                )
                render_plotly_responsive(fig_jefa, height=320)
            else:
                st.subheader("Coste medio MO por jefa de línea (€/kg)")
                st.bar_chart(df_jefa.set_index(col_jefa)["Coste Medio MO (€/kg)"])
        else:
            st.info("No se encuentra columna de jefa/responsable de línea.")

    st.divider()

    g3, g4 = st.columns(2)

    with g3:
        if "FECHA" in df.columns and "TOTAL IMPORTES" in df.columns and "TOTAL KG CONFECCIONADOS" in df.columns:
            df["Semana_inicio"] = (
                df["FECHA"].dt.to_period("W").apply(lambda r: r.start_time.strftime("%Y-%m-%d"))
            )
            df_sem = calcular_coste_medio_agrupado(
                df, "Semana_inicio", nombre_metrica="Coste Medio MO (€/kg)"
            ).sort_values("Semana_inicio")
            if HAS_PLOTLY:
                fig_sem_mo = px.line(
                    df_sem,
                    x="Semana_inicio",
                    y="Coste Medio MO (€/kg)",
                    markers=True,
                    title="Coste medio MO semanal (€/kg)",
                    labels={"Semana_inicio": "Semana"},
                )
                render_plotly_responsive(fig_sem_mo, height=320)
            else:
                st.subheader("Coste medio MO semanal (€/kg)")
                st.line_chart(df_sem.set_index("Semana_inicio")["Coste Medio MO (€/kg)"])

    with g4:
        if "FECHA" in df.columns and "TOTAL IMPORTES" in df.columns and "TOTAL KG CONFECCIONADOS" in df.columns:
            modo_comp = st.radio(
                "Eje comparativa interanual",
                options=["Mes", "Semana del año"],
                horizontal=True,
                index=0,
            )
            df["Año"] = df["FECHA"].dt.year.astype(str)
            if modo_comp == "Mes":
                df["Tramo_anual"] = df["FECHA"].dt.strftime("%m")
                etiqueta_tramo = "Mes (01-12)"
            else:
                df["Tramo_anual"] = df["FECHA"].dt.isocalendar().week.astype(str).str.zfill(2)
                etiqueta_tramo = "Semana del año"

            df_inter = calcular_coste_medio_agrupado(
                df, ["Tramo_anual", "Año"], nombre_metrica="Coste Medio MO (€/kg)"
            ).sort_values(["Tramo_anual", "Año"])

            if HAS_PLOTLY:
                fig_inter = px.line(
                    df_inter,
                    x="Tramo_anual",
                    y="Coste Medio MO (€/kg)",
                    color="Año",
                    markers=True,
                    title="Coste medio MO vs. años anteriores (€/kg)",
                    labels={"Tramo_anual": etiqueta_tramo},
                )
                render_plotly_responsive(fig_inter, height=300, rotar_x=False)
            else:
                st.subheader("Coste medio MO vs. años anteriores (€/kg)")
                piv_inter = df_inter.pivot(
                    index="Tramo_anual", columns="Año", values="Coste Medio MO (€/kg)"
                )
                st.line_chart(piv_inter)

    st.divider()

    st.subheader("Desviaciones globales por familia y confección")
    if col_familia and col_confeccion and "TOTAL KG CONFECCIONADOS" in df.columns:
        df_desv_base = df.dropna(subset=[col_familia, col_confeccion]).copy()
        if not df_desv_base.empty:
            df_desv_base["_kg"] = df_desv_base["TOTAL KG CONFECCIONADOS"]
            df_desv_base["_imp"] = (
                df_desv_base["TOTAL IMPORTES"] if "TOTAL IMPORTES" in df_desv_base.columns else 0.0
            )
            df_desv_base["_est_pond"] = (
                df_desv_base[col_coste_est] * df_desv_base["_kg"]
                if col_coste_est
                else 0.0
            )
            df_desv_base["_real_pond"] = (
                df_desv_base[col_coste_mo] * df_desv_base["_kg"]
                if col_coste_mo
                else df_desv_base["_imp"]
            )
            df_desv_base["_desv_pond"] = (
                df_desv_base[col_desv] * df_desv_base["_kg"]
                if col_desv
                else (df_desv_base["_real_pond"] - df_desv_base["_est_pond"])
            )

            tabla_desv = (
                df_desv_base.groupby([col_familia, col_confeccion], as_index=False)
                .agg(
                    Kg_Confeccionados=("_kg", "sum"),
                    Importes_Totales=("_imp", "sum"),
                    Real_Pond=("_real_pond", "sum"),
                    Est_Pond=("_est_pond", "sum"),
                    Desv_Pond=("_desv_pond", "sum"),
                )
            )
            tabla_desv["Coste Real (€/kg)"] = tabla_desv.apply(
                lambda r: (r["Importes_Totales"] / r["Kg_Confeccionados"])
                if r["Kg_Confeccionados"] > 0
                else 0.0,
                axis=1,
            )
            tabla_desv["Coste Estimado (€/kg)"] = tabla_desv.apply(
                lambda r: (r["Est_Pond"] / r["Kg_Confeccionados"])
                if r["Kg_Confeccionados"] > 0
                else 0.0,
                axis=1,
            )
            tabla_desv["Desviación Global (€/kg)"] = tabla_desv.apply(
                lambda r: (r["Desv_Pond"] / r["Kg_Confeccionados"])
                if r["Kg_Confeccionados"] > 0
                else 0.0,
                axis=1,
            )
            tabla_desv["Desviación (%)"] = tabla_desv.apply(
                lambda r: ((r["Desviación Global (€/kg)"] / r["Coste Estimado (€/kg)"]) * 100)
                if r["Coste Estimado (€/kg)"] > 0
                else 0.0,
                axis=1,
            )
            tabla_desv = tabla_desv.sort_values("Desviación Global (€/kg)", ascending=True)
            tabla_desv["Familia_Confeccion"] = (
                tabla_desv[col_familia].astype(str) + " — " + tabla_desv[col_confeccion].astype(str)
            )

            if HAS_PLOTLY:
                altura_desv = max(300, len(tabla_desv) * 34)
                fig_desv = px.bar(
                    tabla_desv,
                    y="Familia_Confeccion",
                    x="Desviación Global (€/kg)",
                    color=col_confeccion,
                    orientation="h",
                    title="Desviación global real vs. estimado (€/kg)",
                    text_auto=".3f",
                    labels={
                        "Familia_Confeccion": "",
                        "Desviación Global (€/kg)": "Desviación (€/kg)",
                    },
                )
                render_plotly_responsive(fig_desv, height=altura_desv, rotar_x=False)
            else:
                st.bar_chart(tabla_desv.set_index("Familia_Confeccion")["Desviación Global (€/kg)"])

            tabla_desv_vista = tabla_desv.sort_values("Desviación Global (€/kg)", ascending=False)
            cols_tabla_desv = [
                col_familia,
                col_confeccion,
                "Kg_Confeccionados",
                "Coste Estimado (€/kg)",
                "Coste Real (€/kg)",
                "Desviación Global (€/kg)",
                "Desviación (%)",
            ]
            st.dataframe(
                tabla_desv_vista[cols_tabla_desv].style.format({
                    "Kg_Confeccionados": "{:,.0f}",
                    "Coste Estimado (€/kg)": "{:,.3f}",
                    "Coste Real (€/kg)": "{:,.3f}",
                    "Desviación Global (€/kg)": "{:+,.3f}",
                    "Desviación (%)": "{:+,.1f}%",
                }),
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.info("No hay datos de familia y confección con los filtros actuales.")
    else:
        st.info("No se encuentran las columnas de familia, confección o TOTAL KG CONFECCIONADOS.")

    cols_prioritarias_prod = [
        c for c in [
            "FECHA", "Semana", "Campaña", col_familia, col_confeccion,
            "LINEA PRODUCCIÓN", col_jefa, "TOTAL KG CONFECCIONADOS",
            "TOTAL IMPORTES", col_coste_mo, col_coste_est, col_desv
        ]
        if c and c in df.columns
    ]
    cols_detalle_prod = list(dict.fromkeys(cols_prioritarias_prod + list(df.columns)))

    with st.expander("Detalle de producción filtrada", expanded=False):
        st.dataframe(df[cols_detalle_prod], use_container_width=True, hide_index=True)


# ---------------------------------------------------------
# 7. VISTA C: PREVISIONES / ENTREGAS
# ---------------------------------------------------------
else:
    df = limpiar_entregas(df_raw)
    st.title("Cuadro de Mandos — Previsiones y Entregas")
    st.caption("Seguimiento de kilos previstos por producto, semana y fecha")

    with st.expander("Filtros", expanded=False):
        f1, f2 = st.columns(2)

        with f1:
            if "Campaña" in df.columns and df["Campaña"].notna().any():
                if "Fecha_" in df.columns:
                    campanas = (
                        df.dropna(subset=["Campaña"])
                        .groupby("Campaña")["Fecha_"]
                        .max()
                        .sort_values(ascending=True)
                        .index.tolist()
                    )
                else:
                    campanas = sorted(df["Campaña"].dropna().unique())
                ultima_campana = [campanas[-1]] if campanas else []
                campanas_sel = st.multiselect("Campaña", campanas, default=ultima_campana)
                if campanas_sel:
                    df = df[df["Campaña"].isin(campanas_sel)]

        with f2:
            if "Semana" in df.columns and df["Semana"].notna().any():
                if "Fecha_" in df.columns:
                    semanas = (
                        df.dropna(subset=["Semana"])
                        .assign(Semana_int=lambda d: d["Semana"].astype(int))
                        .groupby("Semana_int")["Fecha_"]
                        .max()
                        .sort_values(ascending=True)
                        .index.tolist()
                    )
                else:
                    semanas = sorted(df["Semana"].dropna().astype(int).unique())
                ultima_semana = [semanas[-1]] if semanas else []
                semanas_sel = st.multiselect("Semana", semanas, default=ultima_semana)
                if semanas_sel:
                    df = df[df["Semana"].isin(semanas_sel)]

        f3, f4 = st.columns(2)
        with f3:
            if "Producto" in df.columns and df["Producto"].notna().any():
                productos = sorted(df["Producto"].dropna().unique())
                productos_sel = st.multiselect("Familia de producto", productos, default=[])
                if productos_sel:
                    df = df[df["Producto"].isin(productos_sel)]

        with f4:
            if "Gestion" in df.columns and df["Gestion"].notna().any():
                gestiones = sorted(df["Gestion"].dropna().unique())
                gestiones_sel = st.multiselect("Gestión", gestiones, default=[])
                if gestiones_sel:
                    df = df[df["Gestion"].isin(gestiones_sel)]

        f5, f6 = st.columns(2)
        with f5:
            if "Fecha_" in df.columns and df["Fecha_"].notna().any():
                min_fecha = df["Fecha_"].min().date()
                max_fecha = df["Fecha_"].max().date()
                rango_fecha = st.date_input(
                    "Fecha",
                    value=(min_fecha, max_fecha),
                    min_value=min_fecha,
                    max_value=max_fecha,
                )
                if isinstance(rango_fecha, tuple) and len(rango_fecha) == 2:
                    f_ini, f_fin = rango_fecha
                    df = df[(df["Fecha_"].dt.date >= f_ini) & (df["Fecha_"].dt.date <= f_fin)]
                elif isinstance(rango_fecha, tuple) and len(rango_fecha) == 1:
                    df = df[df["Fecha_"].dt.date >= rango_fecha[0]]

        with f6:
            if "Responsable" in df.columns and df["Responsable"].notna().any():
                responsables = sorted(df["Responsable"].dropna().unique())
                responsables_sel = st.multiselect("Responsable", responsables, default=[])
                if responsables_sel:
                    df = df[df["Responsable"].isin(responsables_sel)]

    st.caption(f"Registros filtrados: {len(df):,}")
    if df.empty:
        st.warning("No hay registros que coincidan con los filtros seleccionados.")
        st.stop()

    kg_prev = df["Kg"].sum() if "Kg" in df.columns else 0.0

    k1, _ = st.columns([1, 2])
    k1.metric("Total Kg Previstos", f"{kg_prev:,.0f} kg")

    st.divider()

    if "Fecha_" in df.columns and "Kg" in df.columns and "Producto" in df.columns:
        c_g1, c_g2 = st.columns(2)
        with c_g1:
            granularidad_prev = st.radio(
                "Agrupación temporal",
                options=["Semana", "Día"],
                horizontal=True,
                index=0,
            )
        with c_g2:
            modo_barras = st.radio(
                "Tipo de barras",
                options=["Total por producto", "Desglosado por periodo"],
                horizontal=True,
                index=0,
            )

        if granularidad_prev == "Día":
            df["Periodo_prev"] = df["Fecha_"].dt.strftime("%Y-%m-%d")
            etiqueta_per = "Día"
        else:
            df["Periodo_prev"] = df["Fecha_"].dt.to_period("W").apply(
                lambda r: r.start_time.strftime("%Y-%m-%d")
            )
            etiqueta_per = "Semana (inicio)"

        df_prod_base = df.dropna(subset=["Producto"]).copy()

        if modo_barras == "Total por producto":
            df_prev_graf = (
                df_prod_base.groupby("Producto", as_index=False)["Kg"]
                .sum()
                .sort_values("Kg", ascending=False)
            )
            if HAS_PLOTLY:
                fig_prev = px.bar(
                    df_prev_graf,
                    x="Producto",
                    y="Kg",
                    text_auto=".2s",
                    title=f"Kg previstos por producto ({granularidad_prev.lower()})",
                    labels={"Producto": "Producto", "Kg": "Kg previstos"},
                )
                render_plotly_responsive(fig_prev, height=330)
            else:
                st.subheader(f"Kg previstos por producto ({granularidad_prev.lower()})")
                st.bar_chart(df_prev_graf.set_index("Producto")["Kg"])
        else:
            df_prev_graf = (
                df_prod_base.groupby(["Producto", "Periodo_prev"], as_index=False)["Kg"]
                .sum()
                .sort_values(["Kg", "Periodo_prev"], ascending=[False, True])
            )
            if HAS_PLOTLY:
                fig_prev_desg = px.bar(
                    df_prev_graf,
                    x="Producto",
                    y="Kg",
                    color="Periodo_prev",
                    barmode="group",
                    title=f"Kg previstos por producto y {granularidad_prev.lower()}",
                    labels={"Producto": "Producto", "Kg": "Kg previstos", "Periodo_prev": etiqueta_per},
                )
                render_plotly_responsive(fig_prev_desg, height=340)
            else:
                st.subheader(f"Kg previstos por producto y {granularidad_prev.lower()}")
                piv_prev = df_prev_graf.pivot(
                    index="Producto", columns="Periodo_prev", values="Kg"
                ).fillna(0)
                st.bar_chart(piv_prev)

    cols_visibles = [
        c for c in [
            "Fecha_", "Producto", "Kg", "Semana", "Nombre",
            "Responsable", "Gestion", "Calibre", "Campaña", "Ubicacion", "Observaciones"
        ]
        if c in df.columns
    ]
    with st.expander("Detalle de previsiones", expanded=False):
        busqueda_prev = st.text_input(
            "Buscar en el detalle",
            placeholder="Escribe nombre, producto, responsable, ubicación...",
            key="buscador_detalle_previsiones",
        ).strip()

        df_detalle_prev = df[cols_visibles]
        if busqueda_prev:
            mask_busqueda = (
                df_detalle_prev.astype(str)
                .apply(lambda col: col.str.contains(busqueda_prev, case=False, na=False, regex=False))
                .any(axis=1)
            )
            df_detalle_prev = df_detalle_prev[mask_busqueda]

        st.caption(f"Mostrando {len(df_detalle_prev):,} de {len(df):,} registros")
        st.dataframe(df_detalle_prev, use_container_width=True, hide_index=True)
