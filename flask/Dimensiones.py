import pandas as pd

# Ruta al archivo
file_path = '../RecomendacionesCEAL.xlsx'

# Cargar el archivo Excel
excel_data = pd.ExcelFile(file_path)

# Crear un diccionario para almacenar cada hoja como un DataFrame
dataframes = {
    "dimensiones": excel_data.parse('Dimensiones'),
    "preguntas": excel_data.parse('Preguntas'),
    "riesgos": excel_data.parse('Riesgos'),
    "intervenciones": excel_data.parse('Intervenciones')
}

# Combinar las tablas de 'dimensiones' y 'intervenciones' usando 'dimm' para facilitar consultas
dimensiones_intervenciones = pd.merge(
    dataframes['dimensiones'],
    dataframes['intervenciones'],
    on='id_dimension',
    how='left'
)


# Función para obtener y mostrar todas las intervenciones para un código de dimensión específica
def mostrar_intervenciones_por_dimm(codigo_dimm):
    # Filtrar por el código de dimensión 'dimm'
    resultado = dimensiones_intervenciones[dimensiones_intervenciones['dimm'] == codigo_dimm]

    # Verificar si hay resultados
    if resultado.empty:
        print(f"No se encontraron intervenciones para el código de dimensión '{codigo_dimm}'.")
    else:
        # Mostrar cada intervención en una fila separada
        print(f"Intervenciones para el código de dimensión '{codigo_dimm}':\n")
        print(resultado[['dimm', 'intervención']].to_string(index=False))


# Ejemplo de uso
# Cambia 'CT' por el código de dimensión que deseas consultar
# codigo_dimm = 'CT'
# mostrar_intervenciones_por_dimm(codigo_dimm)