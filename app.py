import streamlit as st
import pandas as pd
import base64

from io import BytesIO
from pathlib import Path

from optimizer import optimizar_desde_excel


# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

st.set_page_config(
    page_title="Optimizador Universal",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ============================================================
# IMAGEN DEL HERO
# ============================================================

RUTA_HERO = Path("assets/hero-campus.png")

if RUTA_HERO.exists():
    hero_base64 = base64.b64encode(
        RUTA_HERO.read_bytes()
    ).decode()

    hero_background = (
        f"linear-gradient("
        f"90deg,"
        f"rgba(248,250,252,0.99) 0%,"
        f"rgba(248,250,252,0.96) 36%,"
        f"rgba(248,250,252,0.65) 55%,"
        f"rgba(248,250,252,0.05) 78%"
        f"),"
        f"url('data:image/png;base64,{hero_base64}')"
    )
else:
    hero_background = (
        "linear-gradient(135deg,#EFF6FF,#DBEAFE)"
    )


# ============================================================
# CSS V4
# ============================================================

st.markdown(
    f"""
<style>

/* ==========================================================
   BASE
   ========================================================== */

html, body, [class*="css"] {{
    font-family:
        Arial,
        "Helvetica Neue",
        Helvetica,
        sans-serif;
}}

.stApp {{
    background:
        #F6F8FB;
    color:
        #172033;
}}

.block-container {{
    max-width: 1320px;
    padding-top: 1.7rem;
    padding-bottom: 4rem;
}}


/* ==========================================================
   TIPOGRAFÍA EDITORIAL
   ========================================================== */

.editorial-title,
.section-heading,
.hero-title {{
    font-family:
        Georgia,
        "Times New Roman",
        serif;
}}

.hero-title {{
    font-size: 3rem;
    line-height: 1.03;
    letter-spacing: -0.035em;
    color: #10264A;
    margin: 0.25rem 0 0.8rem 0;
    font-weight: 700;
}}

.hero-kicker {{
    font-size: 0.78rem;
    font-weight: 700;
    letter-spacing: 0.13em;
    text-transform: uppercase;
    color: #3768A8;
    margin-bottom: 0.65rem;
}}

.hero-lead {{
    font-family:
        Georgia,
        "Times New Roman",
        serif;
    font-size: 1.08rem;
    color: #22395D;
    line-height: 1.55;
    max-width: 620px;
}}

.hero-text {{
    color: #52637B;
    line-height: 1.6;
    max-width: 600px;
    margin-top: 0.65rem;
    font-size: 0.96rem;
}}


/* ==========================================================
   HERO
   ========================================================== */

.hero-panel {{
    background-image: {hero_background};
    background-size: cover;
    background-position: center right;
    min-height: 350px;

    border:
        1px solid #DCE5F0;

    border-radius:
        18px;

    padding:
        3rem 2.4rem;

    display:
        flex;

    align-items:
        center;

    box-shadow:
        0 12px 32px
        rgba(28, 50, 84, 0.08);

    margin-bottom:
        2rem;
}}

.hero-content {{
    width: 57%;
}}

.hero-badge {{
    display: inline-block;
    padding: 0.43rem 0.75rem;

    background:
        rgba(231,241,255,0.92);

    color:
        #215B9F;

    border:
        1px solid #C8DCF7;

    border-radius:
        999px;

    font-size:
        0.78rem;

    font-weight:
        700;

    margin-top:
        1.25rem;
}}


/* ==========================================================
   TÍTULOS DE SECCIÓN
   ========================================================== */

.section-heading {{
    font-size:
        1.52rem;

    font-weight:
        700;

    color:
        #10264A;

    margin:
        2.1rem 0 0.25rem 0;
}}

.section-description {{
    color:
        #718096;

    font-size:
        0.91rem;

    margin-bottom:
        1.15rem;
}}

.blue-line {{
    height: 3px;
    width: 44px;
    background: #2563EB;
    border-radius: 20px;
    margin: 0.45rem 0 1.15rem 0;
}}


/* ==========================================================
   TARJETAS DEL FLUJO
   ========================================================== */

.flow-card {{
    background: white;

    border:
        1px solid #DFE7F1;

    border-radius:
        13px;

    padding:
        1.15rem;

    min-height:
        190px;

    box-shadow:
        0 4px 15px
        rgba(31, 55, 90, 0.035);
}}

.flow-number {{
    width:
        32px;

    height:
        32px;

    display:
        flex;

    align-items:
        center;

    justify-content:
        center;

    border-radius:
        50%;

    background:
        #E1EDFF;

    color:
        #1D63D5;

    font-weight:
        700;

    margin-bottom:
        1rem;
}}

.flow-icon {{
    font-size:
        1.5rem;

    margin-bottom:
        0.5rem;
}}

.flow-title {{
    font-family:
        Georgia,
        "Times New Roman",
        serif;

    color:
        #10264A;

    font-size:
        1.05rem;

    font-weight:
        700;

    margin-bottom:
        0.45rem;
}}

.flow-text {{
    color:
        #607087;

    font-size:
        0.9rem;

    line-height:
        1.55;
}}


/* ==========================================================
   BLOQUES PRINCIPALES
   ========================================================== */

.content-panel {{
    background:
        white;

    border:
        1px solid #E1E8F0;

    border-radius:
        14px;

    padding:
        1.2rem 1.35rem;

    margin:
        1rem 0;

    box-shadow:
        0 5px 18px
        rgba(30, 55, 90, 0.035);
}}


/* ==========================================================
   MÉTRICAS
   ========================================================== */

div[data-testid="stMetric"] {{
    background:
        white;

    border:
        1px solid #DFE7F1;

    border-radius:
        13px;

    padding:
        1rem 1rem 0.85rem 1rem;

    min-height:
        112px;

    box-shadow:
        0 4px 14px
        rgba(27, 52, 86, 0.035);
}}

div[data-testid="stMetricLabel"] {{
    color:
        #687990;

    font-size:
        0.82rem;

    font-weight:
        600;
}}

div[data-testid="stMetricValue"] {{
    color:
        #10264A;

    font-family:
        Georgia,
        "Times New Roman",
        serif;

    font-weight:
        700;
}}


/* ==========================================================
   BOTONES
   ========================================================== */

div.stButton > button {{
    min-height:
        48px;

    border-radius:
        8px;

    font-weight:
        700;

    border:
        1px solid #1D5FDA;

    transition:
        0.2s ease;
}}

div.stButton > button[kind="primary"] {{
    background:
        #1769E0;

    color:
        white;
}}

div.stButton > button[kind="primary"]:hover {{
    background:
        #1057BE;

    border-color:
        #1057BE;
}}

div.stDownloadButton > button {{
    min-height:
        48px;

    border-radius:
        8px;

    background:
        white;

    border:
        1px solid #CCD9E7;

    color:
        #17345B;

    font-weight:
        700;
}}

div.stDownloadButton > button:hover {{
    border-color:
        #2A69C7;

    color:
        #195EBF;
}}


/* ==========================================================
   FILE UPLOADER
   ========================================================== */

section[data-testid="stFileUploaderDropzone"] {{
    background:
        #FBFCFE;

    border:
        1.5px dashed #B9CCE5;

    border-radius:
        12px;

    min-height:
        145px;
}}


/* ==========================================================
   DATAFRAME
   ========================================================== */

div[data-testid="stDataFrame"] {{
    border:
        1px solid #DEE6EF;

    border-radius:
        10px;

    overflow:
        hidden;
}}


/* ==========================================================
   ALERTAS
   ========================================================== */

div[data-testid="stAlert"] {{
    border-radius:
        10px;
}}


/* ==========================================================
   TABS
   ========================================================== */

button[data-baseweb="tab"] {{
    font-weight:
        600;
}}

button[data-baseweb="tab"][aria-selected="true"] {{
    color:
        #175FCD;
}}


/* ==========================================================
   SIDEBAR
   ========================================================== */

section[data-testid="stSidebar"] {{
    background:
        #11294B;
}}

section[data-testid="stSidebar"] > div {{
    padding-top:
        1.2rem;
}}

.sidebar-brand {{
    font-family:
        Georgia,
        "Times New Roman",
        serif;

    color:
        white;

    font-size:
        1.55rem;

    line-height:
        1.15;

    font-weight:
        700;

    margin:
        0.3rem 0 0.6rem 0;
}}

.sidebar-copy {{
    color:
        #D5E0EF;

    font-size:
        0.9rem;

    line-height:
        1.6;
}}

.sidebar-label {{
    color:
        #89A9D1;

    font-size:
        0.74rem;

    text-transform:
        uppercase;

    letter-spacing:
        0.11em;

    font-weight:
        700;

    margin-top:
        1.4rem;

    margin-bottom:
        0.75rem;
}}

.sidebar-item {{
    color:
        #F0F5FB;

    font-size:
        0.92rem;

    padding:
        0.43rem 0;

    border-bottom:
        1px solid
        rgba(255,255,255,0.055);
}}


/* ==========================================================
   FOOTER
   ========================================================== */

.footer {{
    border-top:
        1px solid #DEE5ED;

    margin-top:
        2.5rem;

    padding-top:
        1.2rem;

    color:
        #8492A6;

    font-size:
        0.82rem;
}}

</style>
""",
    unsafe_allow_html=True
)


# ============================================================
# FUNCIONES AUXILIARES
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

            df_grupo.to_excel(
                writer,
                sheet_name=f"Grupo {grupo}"[:31],
                index=False
            )

        # Por profesor
        for profesor in df_solucion["Profesor"].unique():

            df_profesor = df_solucion[
                df_solucion["Profesor"]
                == profesor
            ]

            df_profesor.to_excel(
                writer,
                sheet_name=f"Prof {profesor}"[:31],
                index=False
            )

        # Formato básico
        for hoja in writer.book.worksheets:

            hoja.freeze_panes = "A2"

            for columna in hoja.columns:

                letra = columna[0].column_letter
                longitud_maxima = 0

                for celda in columna:

                    if celda.value is not None:

                        longitud_maxima = max(
                            longitud_maxima,
                            len(str(celda.value))
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
            subset=["Grupo", "Horario"],
            keep=False
        ).sum()
    )

    conflicto_profesor = (
        df.duplicated(
            subset=["Profesor", "Horario"],
            keep=False
        ).sum()
    )

    conflicto_espacio = (
        df.duplicated(
            subset=["Espacio", "Horario"],
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

    st.markdown(
        """
        <div class="sidebar-brand">
            Optimizador<br>Universal
        </div>

        <div class="sidebar-copy">
            Plataforma de optimización de horarios,
            espacios y recursos.
        </div>

        <div class="sidebar-label">
            Flujo de trabajo
        </div>

        <div class="sidebar-item">
            01 · Cargar plantilla
        </div>

        <div class="sidebar-item">
            02 · Ejecutar optimización
        </div>

        <div class="sidebar-item">
            03 · Analizar resultados
        </div>

        <div class="sidebar-item">
            04 · Descargar archivos
        </div>

        <div class="sidebar-label">
            Motor
        </div>

        <div class="sidebar-item">
            Investigación de Operaciones
        </div>

        <div class="sidebar-item">
            OR-Tools · CP-SAT
        </div>

        <div class="sidebar-item">
            Python · Streamlit
        </div>
        """,
        unsafe_allow_html=True
    )


# ============================================================
# HERO
# ============================================================

st.markdown(
    """
    <div class="hero-panel">

        <div class="hero-content">

            <div class="hero-kicker">
                PLANIFICA · OPTIMIZA · TRANSFORMA
            </div>

            <div class="hero-title">
                Optimizador Universal
            </div>

            <div class="hero-lead">
                Convierte datos operativos
                en horarios eficientes.
            </div>

            <div class="hero-text">
                Una plataforma flexible para instituciones
                educativas y organizaciones que buscan
                aprovechar mejor sus recursos mediante
                modelos de optimización.
            </div>

            <div class="hero-badge">
                Basada en Investigación de Operaciones
            </div>

        </div>

    </div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# CÓMO FUNCIONA
# ============================================================

st.markdown(
    '<div class="section-heading">¿Cómo funciona?</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="blue-line"></div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="section-description">'
    'Cuatro pasos para convertir tus datos en una solución optimizada.'
    '</div>',
    unsafe_allow_html=True
)

p1, p2, p3, p4 = st.columns(4)

with p1:

    st.markdown(
        """
        <div class="flow-card">
            <div class="flow-number">1</div>
            <div class="flow-icon">📄</div>
            <div class="flow-title">
                Carga tu plantilla
            </div>
            <div class="flow-text">
                Sube el archivo Excel con la información
                de grupos, personal, espacios y horarios.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

with p2:

    st.markdown(
        """
        <div class="flow-card">
            <div class="flow-number">2</div>
            <div class="flow-icon">⚙️</div>
            <div class="flow-title">
                Ejecuta la optimización
            </div>
            <div class="flow-text">
                El motor analiza restricciones
                y busca una solución factible.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

with p3:

    st.markdown(
        """
        <div class="flow-card">
            <div class="flow-number">3</div>
            <div class="flow-icon">📊</div>
            <div class="flow-title">
                Analiza los resultados
            </div>
            <div class="flow-text">
                Revisa horarios, indicadores,
                carga docente y uso de espacios.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

with p4:

    st.markdown(
        """
        <div class="flow-card">
            <div class="flow-number">4</div>
            <div class="flow-icon">⬇️</div>
            <div class="flow-title">
                Descarga tus archivos
            </div>
            <div class="flow-text">
                Exporta la solución generada
                en Excel o CSV.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


# ============================================================
# CARGAR PLANTILLA
# ============================================================

st.markdown(
    '<div class="section-heading">'
    '1. Cargar plantilla'
    '</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="blue-line"></div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="section-description">'
    'Selecciona el archivo Excel con los datos que deseas optimizar.'
    '</div>',
    unsafe_allow_html=True
)

archivo_excel = st.file_uploader(
    "Sube tu plantilla Excel",
    type=["xlsx"],
    label_visibility="collapsed"
)


# ============================================================
# MANEJO DE CAMBIO DE ARCHIVO
# ============================================================

if archivo_excel is not None:

    nombre_actual = archivo_excel.name

    if (
        st.session_state.get(
            "archivo_actual"
        )
        != nombre_actual
    ):

        st.session_state[
            "archivo_actual"
        ] = nombre_actual

        st.session_state.pop(
            "resultado_optimizacion",
            None
        )

    st.success(
        f"Archivo cargado correctamente: "
        f"{archivo_excel.name}"
    )

    # ========================================================
    # EJECUTAR
    # ========================================================

    st.markdown(
        '<div class="section-heading">'
        '2. Generar horario óptimo'
        '</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="blue-line"></div>',
        unsafe_allow_html=True
    )

    if st.button(
        "⚙️ Ejecutar optimización",
        type="primary",
        use_container_width=True
    ):

        with st.spinner(
            "Analizando restricciones y buscando "
            "la mejor solución disponible..."
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
# RESULTADO GUARDADO
# ============================================================

resultado = st.session_state.get(
    "resultado_optimizacion"
)


if resultado is not None:

    # ========================================================
    # RESULTADO CORRECTO
    # ========================================================

    if resultado.get("estado") == "OK":

        df_solucion = resultado[
            "solucion"
        ].copy()

        conflictos = calcular_conflictos(
            df_solucion
        )

        st.success(
            "Optimización completada correctamente."
        )

        # ====================================================
        # DASHBOARD
        # ====================================================

        st.markdown(
            '<div class="section-heading">'
            '3. Dashboard'
            '</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            '<div class="blue-line"></div>',
            unsafe_allow_html=True
        )

        st.markdown(
            '<div class="section-description">'
            'Resumen ejecutivo de la solución generada.'
            '</div>',
            unsafe_allow_html=True
        )

        m1, m2, m3, m4 = st.columns(4)

        m1.metric(
            "Eventos",
            len(df_solucion)
        )

        m2.metric(
            "Grupos",
            df_solucion[
                "Grupo"
            ].nunique()
        )

        m3.metric(
            "Profesores utilizados",
            df_solucion[
                "Profesor"
            ].nunique()
        )

        m4.metric(
            "Espacios utilizados",
            df_solucion[
                "Espacio"
            ].nunique()
        )

        m5, m6, m7, m8 = st.columns(4)

        m5.metric(
            "Bloques utilizados",
            df_solucion[
                "Horario"
            ].nunique()
        )

        m6.metric(
            "Materias",
            df_solucion[
                "Materia"
            ].nunique()
        )

        m7.metric(
            "Conflictos",
            conflictos
        )

        m8.metric(
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
                f"Se detectaron {conflictos} "
                "posibles conflictos."
            )

        # ====================================================
        # RESULTADOS
        # ====================================================

        st.markdown(
            '<div class="section-heading">'
            '4. Resultados'
            '</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            '<div class="blue-line"></div>',
            unsafe_allow_html=True
        )

        tab1, tab2, tab3, tab4 = st.tabs(
            [
                "Horario general",
                "Por grupo",
                "Carga docente",
                "Uso de espacios"
            ]
        )

        # ----------------------------------------------------
        # HORARIO GENERAL
        # ----------------------------------------------------

        with tab1:

            st.markdown(
                "### Horario optimizado"
            )

            st.dataframe(
                df_solucion,
                use_container_width=True,
                hide_index=True
            )

        # ----------------------------------------------------
        # POR GRUPO
        # ----------------------------------------------------

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

            df_grupo = df_solucion[
                df_solucion["Grupo"]
                == grupo_seleccionado
            ]

            st.dataframe(
                df_grupo,
                use_container_width=True,
                hide_index=True
            )

        # ----------------------------------------------------
        # CARGA DOCENTE
        # ----------------------------------------------------

        with tab3:

            carga_docente = (
                df_solucion
                .groupby("Profesor")
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

        # ----------------------------------------------------
        # USO DE ESPACIOS
        # ----------------------------------------------------

        with tab4:

            uso_espacios = (
                df_solucion
                .groupby("Espacio")
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

        # ====================================================
        # DESCARGAS
        # ====================================================

        st.markdown(
            '<div class="section-heading">'
            '5. Descargar resultados'
            '</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            '<div class="blue-line"></div>',
            unsafe_allow_html=True
        )

        st.markdown(
            '<div class="section-description">'
            'Lleva la solución a donde la necesites.'
            '</div>',
            unsafe_allow_html=True
        )

        excel_resultado = generar_excel(
            df_solucion
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

        d1, d2 = st.columns(2)

        with d1:

            st.download_button(
                "📊 Descargar Excel profesional",
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

        with d2:

            st.download_button(
                "📄 Descargar CSV",
                data=csv_resultado,
                file_name=(
                    "Horario_Optimizado.csv"
                ),
                mime="text/csv",
                use_container_width=True
            )


    # ========================================================
    # SIN SOLUCIÓN
    # ========================================================

    elif resultado.get(
        "estado"
    ) == "SIN_SOLUCION":

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
                """
                Revisa la disponibilidad del personal,
                la capacidad y tipo de los espacios,
                las materias habilitadas y la cantidad
                de bloques horarios disponibles.
                """
            )


    # ========================================================
    # ERROR
    # ========================================================

    else:

        st.error(
            "No fue posible procesar la plantilla."
        )

        st.write(
            resultado.get(
                "mensaje",
                "Error desconocido."
            )
        )


elif archivo_excel is None:

    st.info(
        "Carga una plantilla Excel "
        "para comenzar la optimización."
    )


# ============================================================
# FOOTER
# ============================================================

st.markdown(
    """
    <div class="footer">
        Optimizador Universal · Plataforma de optimización
        basada en Investigación de Operaciones
    </div>
    """,
    unsafe_allow_html=True
)
