import streamlit as st
import pandas as pd
import numpy as np

# Título de la aplicación
st.title('Dashboard Interactivo con Barra Lateral')

# Barra lateral
st.sidebar.header('Configuración del Dashboard')

# Selección de datos
dataset = st.sidebar.selectbox('Selecciona un dataset:',
                               ['Dataset A', 'Dataset B', 'Dataset C'])

# Filtros de datos
st.sidebar.subheader('Filtros')
columna = st.sidebar.selectbox('Selecciona una columna:', ['Columna 1', 'Columna 2', 'Columna 3'])
valor_min = st.sidebar.slider('Valor mínimo:', 0, 100, 25)
valor_max = st.sidebar.slider('Valor máximo:', 0, 100, 75)

# Opciones de visualización
st.sidebar.subheader('Opciones de Visualización')
mostrar_grafico = st.sidebar.checkbox('Mostrar gráfico de líneas')
mostrar_tabla = st.sidebar.checkbox('Mostrar tabla de datos')

# Cargar datos (ejemplo simulado)
def cargar_datos(nombre):
    np.random.seed(0)
    data = {
        'Columna 1': np.random.randint(0, 100, 50),
        'Columna 2': np.random.randint(0, 100, 50),
        'Columna 3': np.random.randint(0, 100, 50),
    }
    df = pd.DataFrame(data)
    return df

df = cargar_datos(dataset)

# Aplicar filtros
df_filtrado = df[(df[columna] >= valor_min) & (df[columna] <= valor_max)]

# Mostrar tabla
if mostrar_tabla:
    st.subheader('Tabla de Datos Filtrados')
    st.write(df_filtrado)

# Mostrar gráfico
if mostrar_grafico:
    st.subheader('Gráfico de Líneas')
    st.line_chart(df_filtrado)

# Información adicional
st.sidebar.markdown('---')
st.sidebar.write('Aplicación creada con Streamlit.')