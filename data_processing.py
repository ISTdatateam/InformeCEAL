import pandas as pd
import unicodedata
import streamlit as st


def cargar_datos(uploaded_files):
    try:
        df_res_com = pd.read_excel(uploaded_files['resultados'], sheet_name='Datos')
        combined_df_base_complet3 = pd.read_excel(uploaded_files['combined'], sheet_name='basecompleta')
        summary_df = pd.read_excel(uploaded_files['combined'], sheet_name='Summary')
        df_resultados_porcentaje = pd.read_excel(uploaded_files['combined'], sheet_name='resultado')
        df_porcentajes_niveles = pd.read_excel(uploaded_files['combined'], sheet_name='df_porcentajes_niveles')
        df_res_dimTE3 = pd.read_excel(uploaded_files['combined'], sheet_name='df_res_dimTE3')
        df_resumen = pd.read_excel(uploaded_files['combined'], sheet_name='df_resumen')
        top_glosas = pd.read_excel(uploaded_files['combined'], sheet_name='top_glosas')

        df_rec = pd.read_excel(uploaded_files['recomendaciones'], sheet_name='Hoja1')
        df_ciiu = pd.read_excel(uploaded_files['ciiu'], sheet_name='ciiu')

        df_resultados = pd.read_excel(uploaded_files['resultados'], sheet_name='Datos', usecols=['CUV', 'Folio'])
        df_res_com_resultados = pd.read_excel(uploaded_files['resultados'], sheet_name='Datos')
        df_resultados['CUV'] = pd.to_numeric(df_resultados['CUV'], errors='coerce').astype('Int64')

        st.success("Todos los archivos se cargaron exitosamente.")
        return {
            'df_res_com': df_res_com,
            'combined_df_base_complet3': combined_df_base_complet3,
            'summary_df': summary_df,
            'df_resultados_porcentaje': df_resultados_porcentaje,
            'df_porcentajes_niveles': df_porcentajes_niveles,
            'df_res_dimTE3': df_res_dimTE3,
            'df_resumen': df_resumen,
            'top_glosas': top_glosas,
            'df_rec': df_rec,
            'df_ciiu': df_ciiu,
            'df_resultados': df_resultados,
            'df_res_com_resultados': df_res_com_resultados
        }
    except Exception as e:
        st.error(f"Error al cargar los datos: {e}")
        return None


def normalizar_texto(texto):
    if isinstance(texto, str):
        texto = texto.strip().lower()
        texto = ''.join(
            c for c in unicodedata.normalize('NFD', texto)
            if unicodedata.category(c) != 'Mn'
        )
        return texto
    else:
        return ''


def procesar_recomendaciones(df_rec, df_ciiu):
    df_reco = pd.DataFrame(columns=['Dimensión', 'Rubro', 'Recomendación'])
    for _, row in df_rec.iterrows():
        dimension = row['Dimensión']
        for i in range(0, len(row) - 1, 2):
            rubro_col = f'Rubro.{i // 2}' if i // 2 > 0 else 'Rubro'
            recomendacion_col = f'Recomendación.{i // 2}' if i // 2 > 0 else 'Recomendación'

            if rubro_col in row and recomendacion_col in row:
                rubro = row[rubro_col]
                recomendacion = row[recomendacion_col]

                if pd.notna(rubro) and pd.notna(recomendacion):
                    df_reco = pd.concat(
                        [df_reco, pd.DataFrame([[dimension, rubro, recomendacion]], columns=df_reco.columns)],
                        ignore_index=True
                    )
    df_recomendaciones = pd.merge(df_ciiu, df_reco, left_on='Sección', right_on='Rubro', how='left')
    df_recomendaciones['ciiu'] = df_recomendaciones['ciiu'].astype(str)
    return df_recomendaciones


def convertir_columnas(df_dict):
    for key, df in df_dict.items():
        if 'CUV' in df.columns:
            df['CUV'] = df['CUV'].astype(str)
    return df_dict


def convertir_fechas(df, columna_fecha_inicio='Fecha Inicio'):
    if columna_fecha_inicio in df.columns:
        df['Fecha Inicio'] = pd.to_datetime(df[columna_fecha_inicio], errors='coerce').dt.strftime('%d-%m-%Y')
        st.success("Columna 'Fecha Inicio' procesada correctamente.")
    else:
        st.error(
            f"La columna '{columna_fecha_inicio}' no se encontró en el DataFrame 'df_res_com'. Verifica el nombre de la columna."
        )
        st.write("Columnas disponibles en 'df_res_com':", df.columns.tolist())
