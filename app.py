import streamlit as st
from optimizer import optimizar_desde_excel
st.set_page_config(
    page_title="Optimizador Universal",
    page_icon="⚙️",
    layout="wide"
)

st.title("⚙️ Optimizador Universal")

st.write(
    "Sistema inteligente para la optimización "
    "de horarios, espacios y recursos."
)

st.success("🚀 La aplicación está funcionando correctamente.")
st.divider()

st.divider()

st.subheader("📤 Cargar plantilla")

st.write(
    "Carga la plantilla Excel con la información de grupos, "
    "personal, espacios, horarios y disponibilidad."
)

archivo_excel = st.file_uploader(
    "Selecciona tu archivo Excel",
    type=["xlsx"]
)

if archivo_excel is not None:

    st.success(f"Archivo cargado: {archivo_excel.name}")

    st.divider()

    st.subheader("⚙️ Procesar plantilla")

    if st.button(
        "🚀 Generar horario óptimo",
        type="primary",
        use_container_width=True
    ):

        with st.spinner(
            "Analizando restricciones y buscando la mejor solución..."
        ):

            resultado = optimizar_desde_excel(
                archivo_excel
            )

        # ========================================================
        # RESULTADO CORRECTO
        # ========================================================

        if resultado["estado"] == "OK":

            st.success(
                "✅ Horario optimizado correctamente."
            )

            st.subheader("📊 Resumen")

            col1, col2, col3 = st.columns(3)

            col1.metric(
                "Grupos",
                resultado["grupos"]
            )

            col2.metric(
                "Personal",
                resultado["personal"]
            )

            col3.metric(
                "Espacios",
                resultado["espacios"]
            )

            col4, col5, col6 = st.columns(3)

            col4.metric(
                "Bloques horarios",
                resultado["horarios"]
            )

            col5.metric(
                "Eventos programados",
                resultado["eventos"]
            )

            col6.metric(
                "Registros disponibilidad",
                resultado["disponibilidad"]
            )

            st.divider()

            # ====================================================
            # MOSTRAR HORARIO
            # ====================================================

            st.subheader("📅 Horario optimizado")

            df_solucion = resultado["solucion"]

            st.dataframe(
                df_solucion,
                use_container_width=True,
                hide_index=True
            )

            st.success(
                f"🎉 Se programaron "
                f"{len(df_solucion)} eventos correctamente."
            )

        # ========================================================
        # SIN SOLUCIÓN
        # ========================================================

        elif resultado["estado"] == "SIN_SOLUCION":

            st.error(
                "❌ No se encontró una solución factible."
            )

            st.warning(
                resultado["mensaje"]
            )

        # ========================================================
        # OTRO ERROR
        # ========================================================

        else:

            st.error(
                "❌ No fue posible procesar la plantilla."
            )

            st.write(
                resultado["mensaje"]
            )