import streamlit as st
import pandas as pd
from io import BytesIO

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
# ESTILOS
# ============================================================

st.markdown(
    """
    <style>

    /* Fondo general */
    .stApp {
        background-color: #F8FAFC;
    }

    /* Contenedor principal */
    .block-container {
        padding-top: 2rem;
        padding-bottom: 3rem;
        max-width: 1400px;
    }

    /* Hero */
    .hero {
        background: linear-gradient(
            135deg,
            #0F172A 0%,
            #1E3A8A 55%,
            #2563EB 100%
        );
        padding: 2.3rem 2.5rem;
        border-radius: 22px;
        color: white;
        margin-bottom: 1.8rem;
        box-shadow: 0 18px 45px rgba(15, 23, 42, 0.15);
    }

    .hero h1 {
        font-size: 2.5rem;
        margin-bottom: 0.4rem;
        color: white;
    }

    .hero p {
        font-size: 1.06rem;
        color: #E2E8F0;
        margin-bottom: 0;
        line-height: 1.7;
    }

    .badge {
        display: inline-block;
        background: rgba(255, 255, 255, 0.14);
        color: white;
        padding: 0.38rem 0.75rem;
        border-radius: 999px;
        font-size: 0.84rem;
        font-weight: 600;
        margin-bottom: 1rem;
        border: 1px solid rgba(255,255,255,0.18);
    }

    /* Tarjetas de pasos */
    .step-card {
        background: white;
        border: 1px solid #E2E8F0;
        border-radius: 16px;
        padding: 1rem 1.1rem;
        min-height: 125px;
        box-shadow: 0 7px 20px rgba(15, 23, 42, 0.05);
    }

    .step-number {
        width: 32px;
        height: 32px;
        border-radius: 10px;
        background: #DBEAFE;
        color: #1D4ED8;
        font-weight: 800;
        display: flex;
        align-items: center;
        justify-content: center;
        margin-bottom: 0.7rem;
    }

    .step-title {
        font-weight: 700;
        color: #0F172A;
        margin-bottom: 0.2rem;
    }

    .step-text {
        color: #64748B;
        font-size: 0.92rem;
        line-height: 1.4;
    }

    /* Secciones */
    .section-title {
        font-size: 1.35rem;
        font-weight: 800;
        color: #0F172A;
        margin-top: 1rem;
        margin-bottom: 0.35rem;
    }

    .section-subtitle {
        color: #64748B;
        margin-bottom: 1rem;
    }

    /* Métricas */
    div[data-testid="stMetric"] {
        background: white;
        border: 1px solid #E2E8F0;
        padding: 1rem;
        border-radius: 16px;
        box-shadow: 0 6px 18px rgba(15, 23, 42, 0.05);
    }

    div[data-testid="stMetricLabel"] {
        color: #64748B;
        font-weight: 600;
    }

    div[data-testid="stMetricValue"] {
        color: #0F172A;
        font-weight: 800;
    }

    /* Botones */
    div.stButton > button {
        border-radius: 12px;
        font-weight: 700;
        min-height: 48px;
        border: none;
    }

    div.stButton > button[kind="primary"] {
        background: linear-gradient(
            90deg,
            #2563EB,
            #14B8A6
        );
        color: white;
    }

    div.stDownloadButton > button {
        border-radius: 12px;
        min-height: 46px;
        font-weight: 700;
    }

    /* Dataframes */
    div[data-testid="stDataFrame"] {
        border: 1px solid #E2E8F0;
        border-radius: 14px;
        overflow: hidden;
    }

    /* Alertas */
    div[data-testid="stAlert"] {
        border-radius: 14px;
    }

    /* Sidebar */
    section[data-testid="stSidebar"] {
        background: #0F172A;
    }

    section[data-testid="stSidebar"] * {
        color: #E2E8F0;
    }

    /* Footer */
    .footer {
        margin-top: 2rem;
        padding-top: 1.2rem;
        border-top: 1px solid #E2E8F0;
        color: #94A3B8;
        font-size: 0.88rem;
        text-align: center;
    }

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

        df_solucion.to_excel(
            writer,
            sheet_name="Horario General",
            index=False
        )

        carga_docente = (
            df_solucion
            .groupby("Profesor")
            .size()
            .reset_index(name="Clases")
            .sort_values("Clases", ascending=False)
        )

        carga_docente.to_excel(
            writer,
            sheet_name="Carga Docente",
            index=False
        )

        uso_espacios = (
            df_solucion
            .groupby("Espacio")
            .size()
            .reset_index(name="Clases")
            .sort_values("Clases", ascending=False)
        )

        uso_espacios.to_excel(
            writer,
            sheet_name="Uso de Espacios",
            index=False
        )

        for grupo in df_solucion["Grupo"].unique():

            df_grupo = df_solucion[
                df_solucion["Grupo"] == grupo
            ]

            df_grupo.to_excel(
                writer,
                sheet_name=f"Grupo {grupo}"[:31],
                index=False
            )

        for profesor in df_solucion["Profesor"].unique():

            df_profesor = df_solucion[
                df_solucion["Profesor"] == profesor
            ]

            df_profesor.to_excel(
                writer,
                sheet_name=f"Prof {profesor}"[:31],
                index=False
            )

        for hoja in writer.book.worksheets:

            hoja.freeze_panes = "A2"

            for columna in hoja.columns:

                longitud = 0
                letra = columna[0].column_letter

                for celda in columna:

                    if celda.value is not None:
                        longitud = max(
                            longitud,
                            len(str(celda.value))
                        )

                hoja.column_dimensions[letra].width = min(
                    longitud + 3,
                    40
                )

    salida.seek(0)

    return salida


def calcular_conflictos(df):

    conflictos_grupo = df.duplicated(
        subset=["Grupo", "Horario"],
        keep=False
    ).sum()

    conflictos_profesor = df.duplicated(
        subset=["Profesor", "Horario"],
        keep=False
    ).sum()

    conflictos_espacio = df.duplicated(
        subset=["Espacio", "Horario"],
        keep=False
    ).sum()

    return (
        conflictos_grupo
        + conflictos_profesor
        + conflictos_espacio
    )


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.markdown("## ⚙️ Optimizador Universal")

    st.markdown(
        """
        Plataforma para la asignación inteligente
        de horarios, espacios y recursos.
        """
    )

    st.divider()

    st.markdown("### Flujo")

    st.markdown(
        """
        **1.** Carga tu plantilla  
        **2.** Ejecuta la optimización  
        **3.** Analiza resultados  
        **4.** Descarga tus archivos
        """
    )

    st.divider()

    st.markdown("### Tecnología")

    st.markdown(
        """
        Python  
        OR-Tools CP-SAT  
        Streamlit  
        Pandas
        """
    )


# ============================================================
# HERO
# ============================================================

st.markdown(
    """
    <div class="hero">
        <div class="badge">
            Optimización inteligente con Investigación de Operaciones
        </div>

        <h1>Optimizador Universal</h1>

        <p>
            Transforma datos operativos en asignaciones eficientes.
            Genera horarios considerando disponibilidad de personal,
            capacidad de espacios, recursos y restricciones configurables.
        </p>
    </div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# FLUJO VISUAL
# ============================================================

st.markdown(
    '<div class="section-title">Cómo funciona</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="section-subtitle">'
    'Un flujo simple para convertir tu plantilla en una solución optimizada.'
    '</div>',
    unsafe_allow_html=True
)

paso1, paso2, paso3, paso4 = st.columns(4)

with paso1:
    st.markdown(
        """
        <div class="step-card">
            <div class="step-number">1</div>
            <div class="step-title">Carga</div>
            <div class="step-text">
                Sube la plantilla Excel con los datos de la institución.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

with paso2:
    st.markdown(
        """
        <div class="step-card">
            <div class="step-number">2</div>
            <div class="step-title">Optimiza</div>
            <div class="step-text">
                El motor analiza restricciones y busca una solución factible.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

with paso3:
    st.markdown(
        """
        <div class="step-card">
            <div class="step-number">3</div>
            <div class="step-title">Analiza</div>
            <div class="step-text">
                Revisa indicadores, horarios, carga docente y uso de espacios.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

with paso4:
    st.markdown(
        """
        <div class="step-card">
            <div class="step-number">4</div>
            <div class="step-title">Descarga</div>
            <div class="step-text">
                Exporta los resultados en Excel o CSV.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

st.write("")


# ============================================================
# CARGA DE ARCHIVO
# ============================================================

st.markdown(
    '<div class="section-title">📤 Cargar plantilla</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="section-subtitle">'
    'Selecciona el archivo Excel que contiene los datos a optimizar.'
    '</div>',
    unsafe_allow_html=True
)

archivo_excel = st.file_uploader(
    "Archivo Excel",
    type=["xlsx"],
    label_visibility="collapsed"
)


# ============================================================
# PROCESAMIENTO
# ============================================================

if archivo_excel is not None:

    st.success(
        f"Archivo listo: {archivo_excel.name}"
    )

    st.markdown(
        '<div class="section-title">⚙️ Ejecutar optimización</div>',
        unsafe_allow_html=True
    )

    if st.button(
        "🚀 Generar horario óptimo",
        type="primary",
        use_container_width=True
    ):

        with st.spinner(
            "Analizando restricciones y buscando una solución..."
        ):

            try:

                resultado = optimizar_desde_excel(
                    archivo_excel
                )

            except Exception as error:

                resultado = {
                    "estado": "ERROR",
                    "mensaje": str(error)
                }

        if resultado is None:

            st.error(
                "El motor terminó sin devolver un resultado."
            )

            st.stop()

        # ========================================================
        # OK
        # ========================================================

        if resultado["estado"] == "OK":

            df_solucion = resultado[
                "solucion"
            ].copy()

            conflictos = calcular_conflictos(
                df_solucion
            )

            st.success(
                "✅ Optimización completada correctamente."
            )

            # ====================================================
            # DASHBOARD
            # ====================================================

            st.markdown(
                '<div class="section-title">📊 Dashboard</div>',
                unsafe_allow_html=True
            )

            st.markdown(
                '<div class="section-subtitle">'
                'Resumen ejecutivo de la solución generada.'
                '</div>',
                unsafe_allow_html=True
            )

            c1, c2, c3, c4 = st.columns(4)

            c1.metric(
                "Eventos",
                len(df_solucion)
            )

            c2.metric(
                "Grupos",
                df_solucion["Grupo"].nunique()
            )

            c3.metric(
                "Profesores utilizados",
                df_solucion["Profesor"].nunique()
            )

            c4.metric(
                "Espacios utilizados",
                df_solucion["Espacio"].nunique()
            )

            c5, c6, c7, c8 = st.columns(4)

            c5.metric(
                "Bloques utilizados",
                df_solucion["Horario"].nunique()
            )

            c6.metric(
                "Materias",
                df_solucion["Materia"].nunique()
            )

            c7.metric(
                "Conflictos",
                conflictos
            )

            c8.metric(
                "Eventos programados",
                "100%"
            )

            if conflictos == 0:

                st.success(
                    "Validación completada: "
                    "no se detectaron conflictos de grupo, "
                    "profesor o espacio."
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
                '<div class="section-title">📅 Resultados</div>',
                unsafe_allow_html=True
            )

            tabs = st.tabs(
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

            with tabs[0]:

                st.dataframe(
                    df_solucion,
                    use_container_width=True,
                    hide_index=True
                )

            # ----------------------------------------------------
            # POR GRUPO
            # ----------------------------------------------------

            with tabs[1]:

                grupo_seleccionado = st.selectbox(
                    "Selecciona un grupo",
                    sorted(
                        df_solucion[
                            "Grupo"
                        ].unique()
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

            with tabs[2]:

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

            with tabs[3]:

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
                '<div class="section-title">📥 Descargar resultados</div>',
                unsafe_allow_html=True
            )

            st.markdown(
                '<div class="section-subtitle">'
                'Exporta la solución para consulta, entrega o análisis.'
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
                .encode("utf-8-sig")
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
                    file_name="Horario_Optimizado.csv",
                    mime="text/csv",
                    use_container_width=True
                )

        # ========================================================
        # SIN SOLUCIÓN
        # ========================================================

        elif resultado["estado"] == "SIN_SOLUCION":

            st.error(
                "No se encontró una solución factible."
            )

            st.warning(
                resultado["mensaje"]
            )

            with st.expander(
                "¿Qué puedo revisar?"
            ):

                st.write(
                    """
                    Revisa la disponibilidad del personal,
                    capacidad y tipo de espacios,
                    materias habilitadas y número de bloques
                    horarios disponibles.
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

else:

    st.info(
        "Sube una plantilla Excel para comenzar."
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
