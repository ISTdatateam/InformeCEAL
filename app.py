# app.py

import streamlit as st
import pandas as pd
from data_processing import cargar_datos, procesar_recomendaciones, convertir_columnas, convertir_fechas
from graph_generation import generar_grafico_principal, generar_graficos_por_te3
from report_generation import generar_informe, agregar_tabla_ges_por_dimension
import data_processing
from io import BytesIO
from docx import Document
from datetime import datetime
import pyodbc
import threading

# 1. Configuración de la página (Debe ser la primera función de Streamlit)
st.set_page_config(page_title="Generador de Informes de Riesgos Psicosociales")


# Inicializar dimensiones en el estado de sesión
if 'dimensiones' not in st.session_state:
    st.session_state['dimensiones'] = []

# 2. Título y descripción
st.title("Generador de Informes de Riesgos Psicosociales")
st.write("""
Esta aplicación permite generar informes técnicos basados en datos de riesgos psicosociales.
Por favor, cargue los archivos necesarios y siga las instrucciones.
""")

# 3. Sección 1: Carga de archivos
st.header("1. Carga de archivos")

# Cargar el archivo Excel con múltiples hojas
uploaded_file_combined = st.file_uploader("Selecciona el archivo 'combined_output.xlsx'", type="xlsx")

# Cargar los archivos de recomendaciones y códigos CIIU
# uploaded_file_rec = st.file_uploader("Selecciona el archivo 'Recomendaciones2.xlsx'", type="xlsx")
# uploaded_file_ciiu = st.file_uploader("Selecciona el archivo 'ciiu.xlsx'", type="xlsx")
# Cargar el archivo 'resultados.xlsx'
# uploaded_file_resultados = st.file_uploader("Selecciona el archivo 'resultados.xlsx'", type="xlsx")

#precargas
uploaded_file_rec = 'Recomendaciones2.xlsx'
uploaded_file_ciiu = 'ciiu.xlsx'
uploaded_file_resultados = 'resultados.xlsx'

if all([uploaded_file_combined, uploaded_file_rec, uploaded_file_ciiu, uploaded_file_resultados]):
    # Mapear los archivos cargados
    uploaded_files = {
        'combined': uploaded_file_combined,
        'recomendaciones': uploaded_file_rec,
        'ciiu': uploaded_file_ciiu,
        'resultados': uploaded_file_resultados
    }

    # Cargar y procesar los datos
    data_dict = cargar_datos(uploaded_files)
    if data_dict:
        # Convertir columnas a string donde sea necesario
        data_dict = convertir_columnas(data_dict)

        # Procesar recomendaciones
        df_recomendaciones = procesar_recomendaciones(data_dict['df_rec'], data_dict['df_ciiu'])

        # Convertir fechas
        convertir_fechas(data_dict['df_res_com'])

        # Asegurarse de que 'top_glosas' esté cargado
        top_glosas = data_dict.get('top_glosas', pd.DataFrame())

        # Sección 2: Selección de CUV para generar el informe
        st.header("2. Seleccionar CUV para generar el informe")

        cuvs_disponibles = data_dict['summary_df']['CUV'].unique()
        selected_cuv = st.selectbox("Selecciona un CUV", cuvs_disponibles)

        # Filtrar los datos para el CUV seleccionado
        datos = data_dict['df_res_com'][data_dict['df_res_com']['CUV'] == selected_cuv]
        estado = data_dict['summary_df'][data_dict['summary_df']['CUV'] == selected_cuv]

        if datos.empty or estado.empty:
            st.error(f"No se encontraron datos para el CUV {selected_cuv}.")
        else:
            # Obtener la primera fila de datos
            datos = datos.iloc[0]
            estado_riesgo = estado['Riesgo'].values[0]

            # Sección 3: Generación y visualización de gráficos
            st.header("3. Gráfico principal de resultados")

            # Generar y mostrar el gráfico principal
            fig_principal = generar_grafico_principal(data_dict['df_resultados_porcentaje'], selected_cuv)
            if fig_principal:
                st.pyplot(fig_principal)
            else:
                st.warning("No se pudo generar el gráfico principal.")

            # Sección 4: Generación de gráficos por TE3
            st.header("4. Generación de gráficos por GES")

            # Generar y mostrar los gráficos por TE3
            figs_te3 = generar_graficos_por_te3(data_dict['df_porcentajes_niveles'], selected_cuv)

            if figs_te3:
                for fig_te3, te3 in figs_te3:
                    st.subheader(f"Gráfico para GES: {te3}")
                    st.pyplot(fig_te3)
            else:
                st.info("No se generaron gráficos adicionales por TE3.")

            # Sección 5: Prescripciones de medidas
            st.header("5. Prescripciones de medidas")

            # Iterar sobre las dimensiones para mostrar campos de entrada
            for idx, dim in enumerate(st.session_state.dimensiones):
                st.subheader(f"Dimensión en riesgo: **{dim['dimension']}**")

                # Resultados (No editables)
                st.markdown("**Resultados:**")
                st.text(dim.get('resultados', "No disponible"))

                # Preguntas clave (No editables)
                st.markdown("**Preguntas clave:**")
                for pregunta in dim.get('preguntas', []):
                    st.text(f"- {pregunta}")

                # Campo editable para la explicación
                st.markdown("**Explicación:**")
                dim['explicacion'] = st.text_area(
                    f"Explicación para {dim['dimension']}", value=dim.get('explicacion', ""), key=f"explicacion_{idx}"
                )

                # Medidas propuestas (editable y prellenado)
                st.markdown("**Medidas propuestas:**")
                medidas = dim.get('medidas', [])
                for i, medida in enumerate(medidas):
                    medida['medida'] = st.text_area(
                        f"Medida {i + 1} para {dim['dimension']}", value=medida['medida'], key=f"medida_{idx}_{i}"
                    )
                    medida['fecha_monitoreo'] = st.selectbox(
                        f"Fecha monitoreo para medida {i + 1}",
                        options=["01-03-2024", "01-06-2024", "01-09-2024"],
                        index=["01-03-2024", "01-06-2024", "01-09-2024"].index(medida['fecha_monitoreo']),
                        key=f"fecha_{idx}_{i}"
                    )
                    medida['responsable'] = st.text_input(
                        f"Responsable seguimiento para medida {i + 1}", value=medida['responsable'],
                        key=f"responsable_{idx}_{i}"
                    )

                    # Botón para eliminar una medida
                    if st.button(f"Eliminar medida {i + 1} para {dim['dimension']}", key=f"eliminar_medida_{idx}_{i}"):
                        medidas.pop(i)
                        st.experimental_rerun()

                # Botón para agregar una nueva medida
                if st.button(f"Agregar nueva medida para {dim['dimension']}", key=f"agregar_medida_{idx}"):
                    medidas.append({'medida': '', 'fecha_monitoreo': "01-03-2024", 'responsable': ''})
                    st.experimental_rerun()

                # Guardar las medidas en el estado de sesión
                dim['medidas'] = medidas

                # Separador entre dimensiones
                st.markdown("---")

            # Sección 6: Generación del informe en Word
            st.header("6. Generación del informe en Word")


            def on_generate_report():
                doc_buffer = generar_informe(
                    df_res_com=data_dict['df_res_com'],
                    summary_df=data_dict['summary_df'],
                    df_resultados_porcentaje=data_dict['df_resultados_porcentaje'],
                    df_porcentajes_niveles=data_dict['df_porcentajes_niveles'],
                    CUV=selected_cuv,
                    df_resumen=data_dict['df_resumen'],
                    df_res_dimTE3=data_dict['df_res_dimTE3'],
                    generar_grafico_principal=generar_grafico_principal,
                    generar_graficos_por_te3=generar_graficos_por_te3,
                    dimensiones=st.session_state.dimensiones,
                    cuv=selected_cuv,
                    df_recomendaciones=df_recomendaciones,
                    top_glosas=top_glosas,
                    datos_df=data_dict['df_res_com']
                )

                if doc_buffer:
                    st.download_button(
                        label="Descargar informe",
                        data=doc_buffer,
                        file_name=f"Informe_{selected_cuv}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                else:
                    st.error("No se pudo generar el informe.")


            if st.button("Generar informe en Word"):
                with st.spinner("Generando el informe, por favor espera..."):
                    on_generate_report()


else:
    st.info("Por favor, carga todos los archivos requeridos para continuar.")


