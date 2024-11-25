import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# Título de la aplicación
st.title("Explorando Streamlit 🚀")

# Sidebar para la navegación
st.sidebar.title("Opciones")
st.sidebar.write("Usa el menú para interactuar.")

# Entrada de texto
name = st.text_input("¿Cuál es tu nombre?", "Usuario")
st.write(f"¡Hola, {name}! Bienvenido a Streamlit.")

# Selector de opciones
opcion = st.sidebar.selectbox("Elige una funcionalidad:",
                              ["Introducción", "Gráfico", "Carga de datos"])

if opcion == "Introducción":
    st.header("Introducción a Streamlit")
    st.write("""
        **Streamlit** es un framework que te permite crear aplicaciones web
        de forma rápida y sencilla utilizando Python.

        Algunas funcionalidades:
        - Creación de gráficos interactivos.
        - Integración con datos y machine learning.
        - Generación de widgets como sliders, botones, y más.
    """)

elif opcion == "Gráfico":
    st.header("Gráfico Interactivo")
    st.write("Generando datos aleatorios para visualizar.")

    # Datos aleatorios
    data = np.random.randn(100, 2)
    df = pd.DataFrame(data, columns=["x", "y"])

    # Seleccionar tipo de gráfico
    tipo = st.selectbox("Selecciona el tipo de gráfico:", ["scatter", "line"])

    if tipo == "scatter":
        fig, ax = plt.subplots()
        ax.scatter(df["x"], df["y"])
        ax.set_title("Scatter Plot")
        st.pyplot(fig)
    elif tipo == "line":
        st.line_chart(df)

elif opcion == "Carga de datos":
    st.header("Carga de Datos")
    st.write("Sube un archivo CSV para analizarlo.")

    # Subir archivo
    file = st.file_uploader("Cargar archivo CSV", type=["csv"])

    if file:
        # Mostrar datos
        data = pd.read_csv(file)
        st.write("Vista previa de los datos:")
        st.dataframe(data)

        # Resumen estadístico
        st.write("Resumen estadístico:")
        st.write(data.describe())
