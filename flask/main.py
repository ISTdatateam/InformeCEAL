# main.py

import logging

# Importar funciones de los módulos
from flask.data_loading import load_excel_files, create_age_range
from flask.data_processing import (
    calculate_scores,
    create_summary,
    calculate_percentage_responses,
)
from flask.report_generation import (
    generate_main_chart,
    generate_te3_charts,
    create_word_report,
)
from flask import config


def main():
    # Configuración de logging
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    # Cargar rutas desde el archivo de configuración
    folder_path = config.FOLDER_PATH
    output_path = config.OUTPUT_PATH
    resultados_path = config.RESULTADOS_PATH
    output_archivos = config.OUTPUT_ARCHIVOS

    # Cargar datos estáticos
    df_ceal = config.DF_CEAL
    df_risk_intervals = config.DF_RISK_INTERVALS

    # Paso 1: Cargar y procesar los datos
    logging.info("Cargando y procesando datos...")
    combined_df = load_excel_files(folder_path)
    combined_df = create_age_range(combined_df)

    # Paso 2: Calcular puntajes y niveles de riesgo
    logging.info("Calculando puntajes y niveles de riesgo...")
    combined_df = calculate_scores(combined_df, df_ceal, df_risk_intervals)
    summary_df, df_resultados_porcentaje = create_summary(combined_df, df_ceal, df_risk_intervals)

    # Paso 3: Calcular porcentajes de respuestas
    logging.info("Calculando porcentajes de respuestas...")
    df_resultados_porcentaje = calculate_percentage_responses(combined_df, df_ceal, df_risk_intervals)

    # Paso 4: Generar informes y gráficos
    logging.info("Generando informes y gráficos...")
    for cuv in summary_df['CUV'].unique():
        # Obtener datos específicos para el CUV
        datos_cuv = summary_df[summary_df['CUV'] == cuv]
        estado_riesgo = datos_cuv['Riesgo'].values[0]

        # Generar gráficos
        img_path_main = generate_main_chart(df_resultados_porcentaje, cuv)
        img_paths_te3 = generate_te3_charts(df_resultados_porcentaje, cuv)

        # Generar informe en Word
        output_file = f"{output_archivos}/{cuv}_informe.docx"
        create_word_report(
            datos_cuv,
            estado_riesgo,
            img_path_main,
            img_paths_te3,
            output_file
        )

    # Paso 5: Guardar resultados en Excel
    logging.info("Guardando resultados en Excel...")
    combined_df.to_excel(output_path, index=False)

    logging.info("Proceso completado exitosamente.")

if __name__ == "__main__":
    main()
