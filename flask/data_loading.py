# data_loading.py

import os
import pandas as pd
import re
import logging

# Configuración de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def rename_duplicate_columns(df):
    """
    Renombra columnas duplicadas en un DataFrame agregando un sufijo numérico.

    Parámetros:
    df (pd.DataFrame): DataFrame con posibles columnas duplicadas.

    Retorna:
    pd.DataFrame: DataFrame con columnas renombradas.
    """
    # Crear un diccionario para contar las apariciones de cada columna
    col_counts = {}
    new_columns = []
    for col in df.columns:
        if col in col_counts:
            col_counts[col] += 1
            new_columns.append(f"{col}_{col_counts[col]}")
        else:
            col_counts[col] = 0
            new_columns.append(col)
    df.columns = new_columns
    return df

def extract_rut_cuv(file_name):
    """
    Extrae RUT, CUV y CDT del nombre de un archivo.

    Parámetros:
    file_name (str): Nombre del archivo.

    Retorna:
    tuple: (rut, cuv, cdtm) si se encontraron, de lo contrario, (None, None, None).
    """
    # Expresiones regulares para extraer RUT, CUV y CDT
    rut_match = re.search(r'\d{8}-[\dkK]', file_name)
    cuv_match = re.search(r'(\d+)(?=\.xlsx)', file_name)
    cdt_match = re.search(r'\d{8}-[\dkK]-(.*?)-\d+\.xlsx', file_name)

    if rut_match and cuv_match and cdt_match:
        rut = rut_match.group(0)
        cuv = cuv_match.group(1)
        cdtm = cdt_match.group(1)
        return rut, cuv, cdtm
    else:
        logging.warning(f"No se pudo extraer RUT o CUV del archivo: {file_name}")
        return None, None, None

def load_excel_files(folder_path):
    """
    Carga y combina archivos Excel desde una carpeta específica.

    Parámetros:
    folder_path (str): Ruta de la carpeta que contiene los archivos Excel.

    Retorna:
    pd.DataFrame: DataFrame combinado de todos los archivos.
    """
    # Obtener lista de archivos .xlsx en la carpeta
    file_list = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]

    data_frames = []

    for file_name in file_list:
        file_path = os.path.join(folder_path, file_name)
        logging.info(f"Procesando archivo: {file_name}")

        try:
            # Leer la hoja "BaseCompleta"
            df = pd.read_excel(file_path, sheet_name='BaseCompleta', header=1, usecols='C:CO')
            df = rename_duplicate_columns(df)
            df.rename(columns={'DD1': 'Genero', 'DD2': 'Edad'}, inplace=True)
            df['Genero'] = df['Genero'].replace({1: 'Hombre', 2: 'Mujer', 3: 'NcOtro', 4: 'NcOtro'})

            # Extraer RUT y CUV del nombre del archivo
            rut, cuv, cdtm = extract_rut_cuv(file_name)
            if rut and cuv and cdtm:
                df['RUT_empleador'] = rut
                df['CUV'] = cuv
                df['CDT_glosa'] = cdtm
            else:
                logging.warning(f"No se pudo extraer RUT o CUV del archivo: {file_name}")

            data_frames.append(df)

        except Exception as e:
            logging.error(f"Error procesando el archivo {file_name}: {e}")

    if data_frames:
        # Combinar todos los DataFrames
        combined_df = pd.concat(data_frames, ignore_index=True)
        logging.info("Archivos combinados exitosamente.")
        return combined_df
    else:
        logging.warning("No se encontraron DataFrames para combinar.")
        return pd.DataFrame()  # Retorna un DataFrame vacío si no hay datos

def create_age_range(df):
    """
    Crea una nueva columna 'Rango Edad' en el DataFrame basado en la edad.

    Parámetros:
    df (pd.DataFrame): DataFrame con la columna 'Edad'.

    Retorna:
    pd.DataFrame: DataFrame con la nueva columna 'Rango Edad'.
    """
    bins = [18, 25, 36, 49, float('inf')]
    labels = ['18 a 25', '26 a 36', '37 a 49', '50 o más']
    df['Rango Edad'] = pd.cut(df['Edad'], bins=bins, labels=labels, right=False)
    return df

# Si deseas probar el módulo directamente
if __name__ == "__main__":
    folder_path = r'path\to\your\folder'  # Actualiza esta ruta según tu estructura
    combined_df = load_excel_files(folder_path)
    combined_df = create_age_range(combined_df)
    # Imprimir las primeras filas para verificar
    print(combined_df.head())
