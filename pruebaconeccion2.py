import streamlit as st
import pyodbc
import pandas as pd
import logging
import os
from dotenv import load_dotenv

# Cargar variables de entorno desde un archivo .env (asegúrate de tener este archivo configurado)
load_dotenv()

# Configuración de logging para monitorear la aplicación
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Configuración de la base de datos para SQL Server utilizando variables de entorno
server = os.getenv('DB_SERVER', '170.110.40.38')  # Reemplaza con tu servidor si es necesario
database = os.getenv('DB_DATABASE', 'ept_modprev')
username = os.getenv('DB_USERNAME', 'usr_ept_modprev')
password = os.getenv('DB_PASSWORD', 'C(Q5N:6+5sIt')  # Asegúrate de manejar esto de forma segura
driver = '{ODBC Driver 17 for SQL Server}'


# Función para conectarse a la base de datos
def get_db_connection():
    """
    Establece una conexión con la base de datos SQL Server.
    Retorna el objeto de conexión si es exitosa, de lo contrario retorna None.
    """
    try:
        connection = pyodbc.connect(
            f'DRIVER={driver};'
            f'SERVER={server};'
            f'DATABASE={database};'
            f'UID={username};'
            f'PWD={password}'
        )
        logging.info("Conexión a la base de datos establecida exitosamente.")
        return connection
    except pyodbc.Error as e:
        st.error(f"Error al conectar a la base de datos: {e}")
        logging.error(f"Error al conectar a la base de datos: {e}")
        return None


# Función para consultar una tabla específica por CUV
def consultar_tabla(tabla, cuv):
    """
    Realiza una consulta SQL a una tabla específica filtrando por CUV.

    Parámetros:
    - tabla: Nombre de la tabla en la base de datos.
    - cuv: Valor del CUV a filtrar.

    Retorna:
    - DataFrame con los resultados de la consulta.
    """
    query = f"SELECT * FROM {tabla} WHERE CUV = ?"
    connection = get_db_connection()

    if connection:
        try:
            df = pd.read_sql(query, connection, params=[cuv])
            logging.info(f"Consulta ejecutada en la tabla '{tabla}' para CUV: {cuv}")
            return df
        except Exception as e:
            st.error(f"Error al consultar la tabla '{tabla}': {e}")
            logging.error(f"Error al consultar la tabla '{tabla}': {e}")
            return pd.DataFrame()
        finally:
            connection.close()
            logging.info("Conexión a la base de datos cerrada.")
    else:
        return pd.DataFrame()


# Lista de tablas a consultar
tablas_a_consultar = [
    "informe_CEAL__Summary",
    "informe_CEAL__basecompleta",
    "informe_CEAL__df_porcentajes_niveles",
    "informe_CEAL__df_res_dimTE3",
    "informe_CEAL__df_resumen",
    "informe_CEAL__resultado",
    "informe_CEAL__top_glosas"
]

# Mapeo de nombres de tablas a nombres amigables para visualización
nombres_amigables = {
    "informe_CEAL__Summary": "Summary",
    "informe_CEAL__basecompleta": "Base Completa",
    "informe_CEAL__df_porcentajes_niveles": "Porcentajes Niveles",
    "informe_CEAL__df_res_dimTE3": "Res Dim TE3",
    "informe_CEAL__df_resumen": "Resumen",
    "informe_CEAL__resultado": "Resultado",
    "informe_CEAL__top_glosas": "Top Glosas"
}

# Título de la aplicación
st.title("Aplicación de Búsqueda por CUV")

# Cuadro de búsqueda para ingresar el CUV
cuv_valor = st.text_input("Ingresa el CUV que deseas buscar:", "")

# Botón para ejecutar la búsqueda
if st.button("Buscar"):
    if cuv_valor.strip() == "":
        st.warning("Por favor, ingresa un valor de CUV antes de continuar.")
    else:
        st.header(f"Resultados para CUV: {cuv_valor}")

        # Diccionario para almacenar los resultados de cada tabla
        resultados = {}

        for tabla in tablas_a_consultar:
            df = consultar_tabla(tabla, cuv_valor)
            nombre_amigable = nombres_amigables.get(tabla, tabla)
            resultados[nombre_amigable] = df

        # Visualizar los resultados
        for nombre, df in resultados.items():
            st.subheader(nombre)
            if not df.empty:
                st.dataframe(df)
            else:
                st.info(f"No se encontraron registros en '{nombre}' para el CUV proporcionado.")

        # (Opcional) Integración adicional con otras tablas si es necesario
        # Por ejemplo, si tienes una tabla de Recomendaciones relacionada con CIIU
        # Puedes realizar consultas adicionales aquí.

# Nota de seguridad
st.sidebar.info(
    """
    **Nota de Seguridad:**
    Las credenciales de la base de datos están codificadas directamente en el código.
    Se recomienda utilizar variables de entorno o servicios de gestión de secretos para manejar información sensible.
    """
)

