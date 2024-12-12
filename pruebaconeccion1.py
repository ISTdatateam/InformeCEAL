import streamlit as st
import pyodbc
import pandas as pd
import logging

# Configuración de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Configuración de la base de datos para SQL Server
server = '170.110.40.38'
database = 'ept_modprev'
username = 'usr_ept_modprev'
password = 'C(Q5N:6+5sIt'
driver = '{ODBC Driver 17 for SQL Server}'

# Función para conectarse a la base de datos
def get_db_connection():
    try:
        connection = pyodbc.connect(
            f'DRIVER={driver};'
            f'SERVER={server};'
            f'DATABASE={database};'
            f'UID={username};'
            f'PWD={password}'
        )
        return connection
    except pyodbc.Error as e:
        st.error(f"Error al conectar a la base de datos: {e}")
        return None

# Interfaz de usuario: campo de búsqueda para el CUV
cuv_valor = st.text_input("Ingresa el CUV que deseas buscar:")

# Botón para ejecutar la búsqueda
if st.button("Buscar"):
    if cuv_valor.strip() == "":
        st.warning("Por favor ingresa un valor de CUV antes de continuar.")
    else:
        # Función para cargar los datos desde la tabla con el CUV ingresado
        def load_data_database(cuv):
            query = "SELECT * FROM informeCEAL_combinado WHERE CUV = ?"
            connection = get_db_connection()

            if connection is not None:
                try:
                    df = pd.read_sql(query, connection, params=[cuv])
                    return df
                except pd.io.sql.DatabaseError as e:
                    st.error(f"Error al ejecutar la consulta SQL: {e}")
                    return pd.DataFrame()  # Retorna un DataFrame vacío en caso de error
                finally:
                    connection.close()
            else:
                return pd.DataFrame()  # Retorna un DataFrame vacío si la conexión falla

        # Cargamos los datos usando el valor ingresado
        datos = load_data_database(cuv_valor)

        # Mostramos el resultado en la app Streamlit
        if not datos.empty:
            st.dataframe(datos)
        else:
            st.info("No se encontraron registros para el CUV proporcionado.")
