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
    layout="wide"
)


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def generar_excel(df_solucion):
    """
    Genera un Excel profesional en memoria con varias hojas.
    """

    salida = BytesIO()

    with pd.ExcelWriter(
        salida,
        engine="openpyxl"
    ) as writer:

        # --------------------------------------------------------
        # HOJA GENERAL
        # --------------------------------------------------------

        df_solucion.to_excel(
            writer,
            sheet_name="Horario General",
            index=False
        )

        # --------------------------------------------------------
        # CARGA DOCENTE
        # --------------------------------------------------------

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

        # --------------------------------------------------------
        # USO DE ESPACIOS
        # --------------------------------------------------------

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

        # --------------------------------------------------------
        # POR GRUPO
        # --------------------------------------------------------

        for grupo in df_solucion["Grupo"].unique():

            df_grupo = df_solucion[
                df_solucion["Grupo"] == grupo
            ]

            nombre_hoja = f"Grupo {grupo}"[:31]

            df_grupo.to_excel(
                writer,
                sheet_name=nombre_hoja,
                index=False
            )

        # --------------------------------------------------------
        # POR PROFESOR
        # --------------------------------------------------------

        for profesor in df_solucion["Profesor"].unique():

            df_profesor = df_solucion[
                df_solucion["Profesor"] == profesor
            ]

            nombre_hoja = f"Prof {profesor}"[:31]

            df_profesor.to_excel(
                writer,
                sheet_name=nombre_hoja,
                index=False
            )

        # --------------------------------------------------------
        # AJUSTAR COLUMNAS
        # --------------------------------------------------------

        for hoja in writer.book.worksheets:

            for columna in hoja.columns:

                longitud_maxima = 0
                letra = columna[0].column_letter

                for celda in columna:

                    valor = celda.value

                    if valor is not None:

                        longitud_maxima = max(
                            longitud_maxima,
                            len(str(valor))
                        )

                hoja.column_dimensions[
                    letra
                ].width = min(
                    longitud_maxima + 3,
                    40
                )

            hoja.freeze_panes = "A2"

    salida.seek(0)

    return salida


def calcular_conflictos(df):
    """
    Comprueba conflictos básicos en la solución.
    """

    conflictos_grupo = (
        df.duplicated(
            subset=["Grupo", "Horario"],
            keep=False
        ).sum()
    )

    conflictos_profesor = (
        df.duplicated(
            subset=["Profesor", "Horario"],
            keep=False
        ).sum()
    )

    conflictos_espacio = (
        df.duplicated(
            subset=["Espacio", "Horario"],
            keep=False
        ).sum()
    )

    total = (
        conflictos_grupo
        + conflictos_profesor
        + conflictos_espacio
    )

    return total


# ============================================================
# ENCABEZADO
# ============================================================

st.title("⚙️ Optimizador Universal")

st.subheader(
    "Optimización inteligente de horarios, "
    "espacios y recursos"
)

st.write(
    """
    Plataforma basada en Investigación de Operaciones
    para generar asignaciones eficientes considerando
    disponibilidad de personal, capacidad de espacios,
    recursos y restricciones operativas.
    """
)

st.divider()


# ============================================================
# CARGAR ARCHIVO
# ============================================================

st.header("📤 1. Cargar plantilla")

st.write(
    "Carga el archivo Excel con la información de "
    "grupos, personal, espacios, horarios y disponibilidad."
)

archivo_excel = st.file_uploader(
    "Selecciona tu archivo Excel",
    type=["xlsx"]
)


# ============================================================
# PROCESAR ARCHIVO
# ============================================================

if archivo_excel is not None:

    st.success(
        f"✅ Archivo cargado: {archivo_excel.name}"
    )

    st.divider()

    st.header("⚙️ 2. Optimización")

    if st.button(
        "🚀 Generar horario óptimo",
        type="primary",
        use_container_width=True
    ):

        with st.spinner(
            "Analizando restricciones y buscando "
            "una solución factible..."
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

        # ========================================================
        # PROTECCIÓN POR SI EL MOTOR NO DEVUELVE RESULTADO
        # ========================================================

        if resultado is None:

            st.error(
                "El motor terminó sin devolver un resultado."
            )

            st.stop()

        # ========================================================
        # SOLUCIÓN ENCONTRADA
        # ========================================================

        if resultado["estado"] == "OK":

            df_solucion = resultado["solucion"].copy()

            conflictos = calcular_conflictos(
                df_solucion
            )

            st.success(
                "🎉 Horario optimizado correctamente."
            )

            # ====================================================
            # DASHBOARD
            # ====================================================

            st.header("📊 3. Dashboard")

            col1, col2, col3, col4 = st.columns(4)

            col1.metric(
                "Eventos",
                len(df_solucion)
            )

            col2.metric(
                "Grupos",
                df_solucion["Grupo"].nunique()
            )

            col3.metric(
                "Profesores",
                df_solucion["Profesor"].nunique()
            )

            col4.metric(
                "Espacios utilizados",
                df_solucion["Espacio"].nunique()
            )

            col5, col6, col7, col8 = st.columns(4)

            col5.metric(
                "Bloques horarios",
                df_solucion["Horario"].nunique()
            )

            col6.metric(
                "Materias",
                df_solucion["Materia"].nunique()
            )

            col7.metric(
                "Conflictos",
                conflictos
            )

            porcentaje = 100

            col8.metric(
                "Eventos programados",
                f"{porcentaje}%"
            )

            # ====================================================
            # VALIDACIÓN
            # ====================================================

            if conflictos == 0:

                st.success(
                    "✅ Validación completada: "
                    "no se detectaron conflictos de "
                    "grupo, profesor o espacio."
                )

            else:

                st.warning(
                    f"⚠️ Se detectaron "
                    f"{conflictos} posibles conflictos."
                )

            st.divider()

            # ====================================================
            # PESTAÑAS
            # ====================================================

            st.header("📅 4. Resultados")

            tab1, tab2, tab3, tab4 = st.tabs(
                [
                    "📅 Horario general",
                    "👥 Por grupo",
                    "👩‍🏫 Carga docente",
                    "🏫 Uso de espacios"
                ]
            )

            # ----------------------------------------------------
            # HORARIO GENERAL
            # ----------------------------------------------------

            with tab1:

                st.subheader(
                    "Horario optimizado"
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

                grupos_disponibles = sorted(
                    df_solucion[
                        "Grupo"
                    ].unique()
                )

                grupo_seleccionado = st.selectbox(
                    "Selecciona un grupo",
                    grupos_disponibles
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

            st.divider()

            # ====================================================
            # DESCARGAS
            # ====================================================

            st.header("📥 5. Descargar resultados")

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

            col_descarga1, col_descarga2 = (
                st.columns(2)
            )

            with col_descarga1:

                st.download_button(
                    label=(
                        "📊 Descargar Excel "
                        "profesional"
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

            with col_descarga2:

                st.download_button(
                    label="📄 Descargar CSV",
                    data=csv_resultado,
                    file_name=(
                        "Horario_Optimizado.csv"
                    ),
                    mime="text/csv",
                    use_container_width=True
                )

            st.success(
                f"✅ {len(df_solucion)} eventos "
                "fueron programados correctamente."
            )

        # ========================================================
        # MODELO SIN SOLUCIÓN
        # ========================================================

        elif resultado["estado"] == "SIN_SOLUCION":

            st.error(
                "❌ No se encontró una solución factible."
            )

            st.warning(
                resultado["mensaje"]
            )

            st.info(
                """
                Revisa principalmente:
                disponibilidad del personal,
                capacidad de los espacios,
                tipos de recurso,
                materias habilitadas y
                cantidad de bloques horarios.
                """
            )

        # ========================================================
        # ERROR
        # ========================================================

        else:

            st.error(
                "❌ No fue posible procesar la plantilla."
            )

            st.write(
                resultado.get(
                    "mensaje",
                    "Error desconocido."
                )
            )

else:

    st.info(
        "👆 Sube una plantilla Excel para comenzar."
    )


# ============================================================
# PIE
# ============================================================

st.divider()

st.caption(
    "Optimizador Universal · "
    "Plataforma basada en Investigación de Operaciones"
)
