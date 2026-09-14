import streamlit as st
import pandas as pd

from io import BytesIO
from pathlib import Path

from optimizer import optimizar_desde_excel


# ============================================================
# CONFIGURACIÓN
# ============================================================

st.set_page_config(
    page_title="Optimizador Universal",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ============================================================
# RUTAS
# ============================================================

BASE_DIR = Path(__file__).parent

HERO_PATH = (
    BASE_DIR
    / "assets"
    / "hero-campus.png"
)


# ============================================================
# CSS
# ============================================================

st.markdown(
    """
    <style>

    /* --------------------------------------------------------
       BASE
    -------------------------------------------------------- */

    .stApp {
        background: #F7F9FC;
    }

    .block-container {
        max-width: 1320px;
        padding-top: 1.5rem;
        padding-bottom: 4rem;
    }

    html, body, [class*="css"] {
        font-family:
            Arial,
            "Helvetica Neue",
            Helvetica,
            sans-serif;
        color: #17233B;
    }


    /* --------------------------------------------------------
       TÍTULOS
    -------------------------------------------------------- */

    h1, h2, h3 {
        font-family:
            Georgia,
            "Times New Roman",
            serif !important;

        color: #0E2A52 !important;
    }

    h1 {
        font-size: 3rem !important;
        letter-spacing: -0.035em;
        line-height: 1.03;
    }

    h2 {
        font-size: 1.65rem !important;
        margin-top: 1.7rem !important;
    }

    h3 {
        font-size: 1.15rem !important;
    }


    /* --------------------------------------------------------
       TEXTO SECUNDARIO
    -------------------------------------------------------- */

    .stCaption,
    div[data-testid="stCaptionContainer"] {
        color: #718096;
    }


        /* --------------------------------------------------------
       SIDEBAR
    -------------------------------------------------------- */

    section[data-testid="stSidebar"] {
        background: #102E55;
        border-right: none;
        min-width: 300px !important;
        max-width: 300px !important;
    }

    section[data-testid="stSidebar"] > div {
        padding-top: 1.4rem;
        padding-left: 1.15rem;
        padding-right: 1.15rem;
    }

    section[data-testid="stSidebar"] * {
        color: #EEF5FF;
    }

    section[data-testid="stSidebar"] h1 {
        font-size: 1.85rem !important;
        line-height: 1.05 !important;
        letter-spacing: -0.02em;
        white-space: normal !important;
        word-break: keep-all !important;
        overflow-wrap: normal !important;
        hyphens: none !important;
        color: white !important;
        max-width: 100%;
    }

    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3 {
        color: white !important;
    }

    section[data-testid="stSidebar"] p {
        line-height: 1.55;
    }

    section[data-testid="stSidebar"] hr {
        border-color: rgba(255, 255, 255, 0.14);
    }

    section[data-testid="stSidebar"] div[data-testid="stCaptionContainer"] {
        color: #BFD0E6 !important;
    }


    /* --------------------------------------------------------
       CONTENEDORES / TARJETAS
    -------------------------------------------------------- */

    div[data-testid="stVerticalBlockBorderWrapper"] {
        background: white;
        border-color: #DFE7F1 !important;
        border-radius: 14px;
        box-shadow: 0 5px 18px rgba(24, 49, 83, 0.04);
    }


    /* --------------------------------------------------------
       MÉTRICAS
    -------------------------------------------------------- */

    div[data-testid="stMetric"] {
        background: white;
        border: 1px solid #DFE7F1;
        border-radius: 13px;
        padding: 1rem;
        min-height: 110px;
        box-shadow: 0 4px 14px rgba(27, 52, 86, 0.04);
    }

    div[data-testid="stMetricLabel"] {
        color: #687990;
        font-size: 0.82rem;
        font-weight: 600;
    }

    div[data-testid="stMetricValue"] {
        color: #102E55;
        font-family:
            Georgia,
            "Times New Roman",
            serif;
        font-weight: 700;
    }


    /* --------------------------------------------------------
       BOTÓN PRINCIPAL
    -------------------------------------------------------- */

    div.stButton > button {
        min-height: 48px;
        border-radius: 8px;
        font-weight: 700;
    }

    div.stButton > button[kind="primary"] {
        background: #1769E0;
        color: white;
        border: 1px solid #1769E0;
    }

    div.stButton > button[kind="primary"]:hover {
        background: #1059C2;
        border-color: #1059C2;
    }


    /* --------------------------------------------------------
       BOTONES DE DESCARGA
    -------------------------------------------------------- */

    div.stDownloadButton > button {
        min-height: 48px;
        border-radius: 8px;
        border: 1px solid #CCD9E7;
        background: white;
        color: #15355F;
        font-weight: 700;
    }

    div.stDownloadButton > button:hover {
        border-color: #1769E0;
        color: #1769E0;
    }


    /* --------------------------------------------------------
       FILE UPLOADER
    -------------------------------------------------------- */

    section[data-testid="stFileUploaderDropzone"] {
        background: #FBFCFE;
        border: 1.5px dashed #B7CAE4;
        border-radius: 12px;
        min-height: 140px;
    }


    /* --------------------------------------------------------
       DATAFRAME
    -------------------------------------------------------- */

    div[data-testid="stDataFrame"] {
        border: 1px solid #DFE7F1;
        border-radius: 10px;
        overflow: hidden;
    }


    /* --------------------------------------------------------
       ALERTAS
    -------------------------------------------------------- */

    div[data-testid="stAlert"] {
        border-radius: 10px;
    }


    /* --------------------------------------------------------
       TABS
    -------------------------------------------------------- */

    button[data-baseweb="tab"] {
        font-weight: 600;
    }

    button[data-baseweb="tab"][aria-selected="true"] {
        color: #1769E0;
    }


    /* --------------------------------------------------------
       IMÁGENES
    -------------------------------------------------------- */

    div[data-testid="stImage"] img {
        border-radius: 14px;
    }


    /* --------------------------------------------------------
       SEPARADORES
    -------------------------------------------------------- */

    hr {
        border-color: #E2E8F0;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# FUNCIONES
# ============================================================

def generar_excel(df_solucion):

    salida = BytesIO()

    with pd.ExcelWriter(
        salida,
        engine="openpyxl"
    ) as writer:

        # Horario general
        df_solucion.to_excel(
            writer,
            sheet_name="Horario General",
            index=False
        )

        # Carga docente
        carga_docente = (
            df_solucion
            .groupby("Profesor")
            .size()
            .reset_index(name="Clases")
            .sort_values(
                "Clases",
                ascending=False
            )
        )

        carga_docente.to_excel(
            writer,
            sheet_name="Carga Docente",
            index=False
        )

        # Uso de espacios
        uso_espacios = (
            df_solucion
            .groupby("Espacio")
            .size()
            .reset_index(name="Clases")
            .sort_values(
                "Clases",
                ascending=False
            )
        )

        uso_espacios.to_excel(
            writer,
            sheet_name="Uso de Espacios",
            index=False
        )

        # Por grupo
        for grupo in df_solucion["Grupo"].unique():

            df_grupo = df_solucion[
                df_solucion["Grupo"] == grupo
            ]

            nombre_hoja = (
                f"Grupo {grupo}"
            )[:31]

            df_grupo.to_excel(
                writer,
                sheet_name=nombre_hoja,
                index=False
            )

        # Por profesor
        for profesor in df_solucion[
            "Profesor"
        ].unique():

            df_profesor = df_solucion[
                df_solucion["Profesor"]
                == profesor
            ]

            nombre_hoja = (
                f"Prof {profesor}"
            )[:31]

            df_profesor.to_excel(
                writer,
                sheet_name=nombre_hoja,
                index=False
            )

        # Ajustes visuales
        for hoja in writer.book.worksheets:

            hoja.freeze_panes = "A2"

            for columna in hoja.columns:

                letra = (
                    columna[0]
                    .column_letter
                )

                longitud_maxima = 0

                for celda in columna:

                    if celda.value is not None:

                        longitud_maxima = max(
                            longitud_maxima,
                            len(
                                str(
                                    celda.value
                                )
                            )
                        )

                hoja.column_dimensions[
                    letra
                ].width = min(
                    longitud_maxima + 3,
                    40
                )

    salida.seek(0)

    return salida


def calcular_conflictos(df):

    conflicto_grupo = (
        df.duplicated(
            subset=[
                "Grupo",
                "Horario"
            ],
            keep=False
        ).sum()
    )

    conflicto_profesor = (
        df.duplicated(
            subset=[
                "Profesor",
                "Horario"
            ],
            keep=False
        ).sum()
    )

    conflicto_espacio = (
        df.duplicated(
            subset=[
                "Espacio",
                "Horario"
            ],
            keep=False
        ).sum()
    )

    return (
        conflicto_grupo
        + conflicto_profesor
        + conflicto_espacio
    )


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.title(
        "Optimizador\nUniversal"
    )

    st.caption(
        "Plataforma de optimización de "
        "horarios, espacios y recursos."
    )

    st.divider()

    st.subheader(
        "Acerca del proyecto"
    )

    st.write(
        "Sistema diseñado para apoyar "
        "la asignación eficiente de "
        "recursos mediante modelos "
        "de optimización."
    )

    st.divider()

    st.subheader(
        "Aplicaciones"
    )

    st.write(
        "🎓 Instituciones educativas"
    )

    st.write(
        "🏢 Organizaciones y empresas"
    )

    st.write(
        "🏫 Gestión de espacios"
    )

    st.write(
        "👥 Asignación de personal"
    )

    st.divider()

    st.subheader(
        "Motor"
    )

    st.write(
        "Investigación de Operaciones"
    )

    st.write(
        "OR-Tools · CP-SAT"
    )

    st.write(
        "Python · Streamlit"
    )

    st.divider()

    st.caption(
        "Optimizador Universal · V5"
    )


# ============================================================
# HERO
# ============================================================

with st.container(border=True):

    hero_left, hero_right = st.columns(
        [1.0, 1.35],
        gap="large"
    )

    with hero_left:

        st.caption(
            "PLANIFICA · OPTIMIZA · TRANSFORMA"
        )

        st.title(
            "Optimizador Universal"
        )

        st.markdown(
            "### Convierte datos operativos en horarios eficientes."
        )

        st.write(
            """
            Una plataforma flexible para instituciones educativas
            y organizaciones que buscan aprovechar mejor sus recursos
            mediante modelos de optimización.
            """
        )

        st.info(
            "📊 Basada en Investigación de Operaciones"
        )

        st.caption(
            "Una mejor organización construye más oportunidades."
        )

    with hero_right:

        if HERO_PATH.exists():

            st.image(
                str(HERO_PATH),
                use_container_width=True
            )

        else:

            st.warning(
                "No se encontró assets/hero-campus.png"
            )

# ============================================================
# BLOQUES DE APLICACIÓN
# ============================================================

st.write("")

ap1, ap2, ap3 = st.columns(3)

with ap1:

    with st.container(
        border=True
    ):

        st.subheader(
            "🎓 Instituciones educativas"
        )

        st.write(
            "Optimiza aulas, horarios, "
            "personal y recursos académicos."
        )

with ap2:

    with st.container(
        border=True
    ):

        st.subheader(
            "🏢 Organizaciones y empresas"
        )

        st.write(
            "Planifica espacios, personal "
            "y recursos de manera eficiente."
        )

with ap3:

    with st.container(
        border=True
    ):

        st.subheader(
            "🌐 Solución flexible"
        )

        st.write(
            "Adaptable a distintos contextos, "
            "restricciones y necesidades."
        )


# ============================================================
# CÓMO FUNCIONA
# ============================================================

st.header(
    "¿Cómo funciona?"
)

st.caption(
    "Cuatro pasos para convertir tus datos "
    "en una solución optimizada."
)

paso1, paso2, paso3, paso4 = (
    st.columns(4)
)


with paso1:

    with st.container(
        border=True
    ):

        st.markdown("### ①")

        st.subheader(
            "Carga tu plantilla"
        )

        st.write(
            "Sube el archivo Excel con "
            "grupos, personal, espacios "
            "y horarios."
        )


with paso2:

    with st.container(
        border=True
    ):

        st.markdown("### ②")

        st.subheader(
            "Ejecuta la optimización"
        )

        st.write(
            "El motor analiza restricciones "
            "y busca una solución factible."
        )


with paso3:

    with st.container(
        border=True
    ):

        st.markdown("### ③")

        st.subheader(
            "Analiza los resultados"
        )

        st.write(
            "Revisa indicadores, horarios, "
            "carga docente y uso de espacios."
        )


with paso4:

    with st.container(
        border=True
    ):

        st.markdown("### ④")

        st.subheader(
            "Descarga tus archivos"
        )

        st.write(
            "Exporta la solución generada "
            "en Excel o CSV."
        )


# ============================================================
# 1. CARGAR PLANTILLA
# ============================================================

st.header(
    "1. Cargar plantilla"
)

st.caption(
    "Selecciona el archivo Excel "
    "que contiene los datos a optimizar."
)

with st.container(
    border=True
):

    archivo_excel = (
        st.file_uploader(
            "Sube tu plantilla Excel",
            type=["xlsx"]
        )
    )


# ============================================================
# CONTROL DE ESTADO
# ============================================================

if archivo_excel is not None:

    archivo_id = (
        archivo_excel.name,
        archivo_excel.size
    )

    if (
        st.session_state.get(
            "archivo_id"
        )
        != archivo_id
    ):

        st.session_state[
            "archivo_id"
        ] = archivo_id

        st.session_state.pop(
            "resultado_optimizacion",
            None
        )

    st.success(
        f"Archivo cargado correctamente: "
        f"{archivo_excel.name}"
    )


    # ========================================================
    # 2. EJECUTAR OPTIMIZACIÓN
    # ========================================================

    st.header(
        "2. Generar horario óptimo"
    )

    st.caption(
        "El motor buscará una asignación "
        "que respete las restricciones configuradas."
    )

    if st.button(
        "⚙️ Ejecutar optimización",
        type="primary",
        use_container_width=True
    ):

        with st.spinner(
            "Analizando restricciones y "
            "buscando la mejor solución..."
        ):

            try:

                resultado = (
                    optimizar_desde_excel(
                        archivo_excel
                    )
                )

            except Exception as error:

                resultado = {
                    "estado": "ERROR",
                    "mensaje": str(error)
                }

        st.session_state[
            "resultado_optimizacion"
        ] = resultado


# ============================================================
# RECUPERAR RESULTADO
# ============================================================

resultado = (
    st.session_state.get(
        "resultado_optimizacion"
    )
)


# ============================================================
# RESULTADO OK
# ============================================================

if (
    resultado is not None
    and resultado.get(
        "estado"
    ) == "OK"
):

    df_solucion = (
        resultado[
            "solucion"
        ].copy()
    )

    conflictos = (
        calcular_conflictos(
            df_solucion
        )
    )

    st.success(
        "✅ Optimización completada "
        "correctamente."
    )


    # ========================================================
    # 3. DASHBOARD
    # ========================================================

    st.header(
        "3. Dashboard"
    )

    st.caption(
        "Resumen de la solución generada."
    )

    met1, met2, met3, met4 = (
        st.columns(4)
    )

    met1.metric(
        "Eventos",
        len(
            df_solucion
        )
    )

    met2.metric(
        "Grupos",
        df_solucion[
            "Grupo"
        ].nunique()
    )

    met3.metric(
        "Profesores utilizados",
        df_solucion[
            "Profesor"
        ].nunique()
    )

    met4.metric(
        "Espacios utilizados",
        df_solucion[
            "Espacio"
        ].nunique()
    )

    met5, met6, met7, met8 = (
        st.columns(4)
    )

    met5.metric(
        "Bloques utilizados",
        df_solucion[
            "Horario"
        ].nunique()
    )

    met6.metric(
        "Materias",
        df_solucion[
            "Materia"
        ].nunique()
    )

    met7.metric(
        "Conflictos",
        conflictos
    )

    met8.metric(
        "Eventos programados",
        "100%"
    )

    if conflictos == 0:

        st.success(
            "Validación completada: "
            "no se detectaron conflictos "
            "de grupo, profesor o espacio."
        )

    else:

        st.warning(
            f"Se detectaron "
            f"{conflictos} posibles conflictos."
        )


    # ========================================================
    # 4. RESULTADOS
    # ========================================================

    st.header(
        "4. Resultados"
    )

    st.caption(
        "Explora la solución desde "
        "diferentes perspectivas."
    )

    tab1, tab2, tab3, tab4 = (
        st.tabs(
            [
                "📅 Horario general",
                "👥 Por grupo",
                "🎓 Carga docente",
                "🏫 Uso de espacios"
            ]
        )
    )


    # --------------------------------------------------------
    # GENERAL
    # --------------------------------------------------------

    with tab1:

        st.subheader(
            "Horario optimizado"
        )

        st.dataframe(
            df_solucion,
            use_container_width=True,
            hide_index=True
        )


    # --------------------------------------------------------
    # GRUPO
    # --------------------------------------------------------

    with tab2:

        grupo_seleccionado = (
            st.selectbox(
                "Selecciona un grupo",
                sorted(
                    df_solucion[
                        "Grupo"
                    ].unique()
                )
            )
        )

        df_grupo = (
            df_solucion[
                df_solucion[
                    "Grupo"
                ]
                == grupo_seleccionado
            ]
        )

        st.dataframe(
            df_grupo,
            use_container_width=True,
            hide_index=True
        )


    # --------------------------------------------------------
    # DOCENTES
    # --------------------------------------------------------

    with tab3:

        carga_docente = (
            df_solucion
            .groupby(
                "Profesor"
            )
            .size()
            .reset_index(
                name="Clases asignadas"
            )
            .sort_values(
                "Clases asignadas",
                ascending=False
            )
        )

        st.dataframe(
            carga_docente,
            use_container_width=True,
            hide_index=True
        )

        st.bar_chart(
            carga_docente.set_index(
                "Profesor"
            )
        )


    # --------------------------------------------------------
    # ESPACIOS
    # --------------------------------------------------------

    with tab4:

        uso_espacios = (
            df_solucion
            .groupby(
                "Espacio"
            )
            .size()
            .reset_index(
                name="Clases asignadas"
            )
            .sort_values(
                "Clases asignadas",
                ascending=False
            )
        )

        st.dataframe(
            uso_espacios,
            use_container_width=True,
            hide_index=True
        )

        st.bar_chart(
            uso_espacios.set_index(
                "Espacio"
            )
        )


    # ========================================================
    # 5. DESCARGAS
    # ========================================================

    st.header(
        "5. Descargar resultados"
    )

    st.caption(
        "Exporta la solución para "
        "consulta, entrega o análisis."
    )

    excel_resultado = (
        generar_excel(
            df_solucion
        )
    )

    csv_resultado = (
        df_solucion
        .to_csv(
            index=False
        )
        .encode(
            "utf-8-sig"
        )
    )

    descarga1, descarga2 = (
        st.columns(2)
    )

    with descarga1:

        st.download_button(
            label=(
                "📊 Descargar Excel profesional"
            ),
            data=excel_resultado,
            file_name=(
                "Horario_Optimizado_"
                "Optimizador_Universal.xlsx"
            ),
            mime=(
                "application/"
                "vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True
        )

    with descarga2:

        st.download_button(
            label=(
                "📄 Descargar CSV"
            ),
            data=csv_resultado,
            file_name=(
                "Horario_Optimizado.csv"
            ),
            mime="text/csv",
            use_container_width=True
        )


# ============================================================
# SIN SOLUCIÓN
# ============================================================

elif (
    resultado is not None
    and resultado.get(
        "estado"
    ) == "SIN_SOLUCION"
):

    st.error(
        "No se encontró una solución factible."
    )

    st.warning(
        resultado.get(
            "mensaje",
            ""
        )
    )

    with st.expander(
        "¿Qué puedo revisar?"
    ):

        st.write(
            "Revisa la disponibilidad del "
            "personal, capacidad y tipo de "
            "espacios, materias habilitadas "
            "y bloques horarios disponibles."
        )


# ============================================================
# ERROR
# ============================================================

elif resultado is not None:

    st.error(
        "No fue posible procesar "
        "la plantilla."
    )

    st.write(
        resultado.get(
            "mensaje",
            "Error desconocido."
        )
    )


# ============================================================
# ESTADO INICIAL
# ============================================================

elif archivo_excel is None:

    st.info(
        "👆 Carga una plantilla Excel "
        "para comenzar."
    )


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "Optimizador Universal · "
    "Investigación de Operaciones "
    "aplicada a la asignación eficiente "
    "de recursos."
)
