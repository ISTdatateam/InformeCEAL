import pandas as pd

def obtener_dimm_por_dimension(nombre_dimension, df_dimensiones):
    nombre_dimension_normalizado = normalizar_texto(nombre_dimension)
    resultado = df_dimensiones[df_dimensiones['dimension_normalizada'] == nombre_dimension_normalizado]
    if not resultado.empty:
        return resultado.iloc[0]['dimm']
    else:
        st.warning(f"No se encontró el código 'dimm' para la dimensión '{nombre_dimension}'.")
        return None

def normalizar_texto(texto):
    import unicodedata
    if isinstance(texto, str):
        texto = texto.strip().lower()
        texto = ''.join(
            c for c in unicodedata.normalize('NFD', texto)
            if unicodedata.category(c) != 'Mn'
        )
        return texto
    else:
        return ''
