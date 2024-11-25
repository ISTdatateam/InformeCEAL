import streamlit as st
import pyodbc
import pandas as pd

# Configuración de la base de datos para SQL Server
server = '170.110.40.38'
database = 'ept_modprev'
username = 'usr_ept_modprev'
password = 'C(Q5N:6+5sIt'
driver = '{ODBC Driver 17 for SQL Server}'

# Función para conectarse a la base de datos
def get_db_connection():
    return pyodbc.connect(
        f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}'
    )

# Consulta SQL para obtener combinaciones únicas
def fetch_unique_combinations():
    query = """
    SELECT DISTINCT RUT, CUV, CdT
    FROM archivo_combinado
    """
    connection = get_db_connection()
    df = pd.read_sql(query, connection)
    connection.close()
    return df

# Interfaz Streamlit
st.title("Consulta a SQL Server")

st.write("Extrayendo combinaciones únicas de los campos: RUT, CUV y CdT de la tabla archivo_combinado.")

if st.button("Ejecutar consulta"):
    try:
        # Obtener los datos
        data = fetch_unique_combinations()
        st.write(f"Total de combinaciones únicas encontradas: {len(data)}")
        st.dataframe(data)  # Mostrar el DataFrame en Streamlit
    except Exception as e:
        st.error(f"Error al ejecutar la consulta: {e}")
