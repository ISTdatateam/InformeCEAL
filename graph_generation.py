# graph_generation.py

import matplotlib.pyplot as plt
import numpy as np
import streamlit as st

# Lazy Loading para gráficos
def generar_grafico_lazy(df, cuv):
    df_subset = df[df['CUV'] == cuv]
    if df_subset.empty:
        st.warning("No se encontraron datos para el CUV seleccionado.")
        return None

    fig = generar_grafico_principal(df_subset, cuv)
    return fig

def generar_grafico_principal(df, CUV):
    df_filtrado = df[df['CUV'] == CUV]
    if df_filtrado.empty:
        st.warning(f"No se encontraron datos para el CUV {CUV}.")
        return None

    try:
        df_pivot = df_filtrado.pivot(index="Dimensión", columns="Nivel", values="Porcentaje").fillna(0).iloc[::-1]
    except Exception as e:
        st.error(f"Error al pivotear los datos: {e}")
        return None

    fig, ax = plt.subplots(figsize=(12, 8))

    niveles = ["Bajo", "Medio", "Alto"]
    colores = {"Bajo": "green", "Medio": "orange", "Alto": "red"}
    posiciones = np.arange(len(df_pivot.index))
    ancho_barra = 0.2

    for i, nivel in enumerate(niveles):
        if nivel in df_pivot.columns:
            valores = df_pivot[nivel]
            ax.barh(posiciones + i * ancho_barra, valores, height=ancho_barra,
                    label=f"Riesgo {nivel.lower()} (%)", color=colores.get(nivel, 'grey'))
        else:
            st.warning(f"Nivel '{nivel}' no encontrado en las columnas de pivot.")

    ax.axvline(50, color="blue", linestyle="--", linewidth=1)
    ax.set_title(f"Porcentaje de trabajadores por nivel de riesgo - CUV {CUV}", pad=50)
    ax.set_xlabel("Porcentaje")
    ax.set_ylabel("Dimensiones")
    ax.set_xlim(0, 100)
    ax.set_yticks(posiciones + ancho_barra)
    ax.set_yticklabels(df_pivot.index, rotation=0, ha='right')

    fig.legend(title="Nivel de Riesgo", loc="upper center", bbox_to_anchor=(0.6, 0.96), ncol=3)
    plt.subplots_adjust(left=0.3, top=0.85)
    plt.tight_layout()

    return fig

def generar_graficos_por_te3(df, CUV):
    df_cuv = df[df['CUV'] == CUV]
    if df_cuv.empty:
        st.warning(f"No se encontraron datos para el CUV {CUV}.")
        return []

    if 'TE3' not in df_cuv.columns:
        st.warning(f"La columna 'TE3' no existe en el DataFrame para CUV {CUV}.")
        return []

    te3_values = df_cuv['TE3'].unique()

    if len(te3_values) <= 1:
        st.info(f"Solo hay un valor de TE3 para el CUV {CUV}. No se generarán gráficos adicionales.")
        return []

    figs_te3 = []

    for te3 in te3_values:
        df_te3 = df_cuv[df_cuv['TE3'] == te3]
        if df_te3.empty:
            continue
        try:
            df_pivot = df_te3.pivot(index="Dimensión", columns="Nivel", values="Porcentaje").fillna(0).iloc[::-1]
        except Exception as e:
            st.error(f"Error al pivotear los datos para TE3 {te3}: {e}")
            continue

        fig, ax = plt.subplots(figsize=(12, 8))

        niveles = ["Bajo", "Medio", "Alto"]
        colores = {"Bajo": "green", "Medio": "orange", "Alto": "red"}
        posiciones = np.arange(len(df_pivot.index))
        ancho_barra = 0.2

        for i, nivel in enumerate(niveles):
            if nivel in df_pivot.columns:
                valores = df_pivot[nivel]
                ax.barh(posiciones + i * ancho_barra, valores, height=ancho_barra,
                        label=f"Riesgo {nivel.lower()} (%)", color=colores.get(nivel, 'grey'))
            else:
                st.warning(f"Nivel '{nivel}' no encontrado en las columnas de pivot para TE3 {te3}.")

        ax.axvline(50, color="blue", linestyle="--", linewidth=1)
        ax.set_title(f"Porcentaje de trabajadores por nivel de riesgo - CUV {CUV}, TE3 {te3}", pad=50)
        ax.set_xlabel("Porcentaje")
        ax.set_ylabel("Dimensiones")
        ax.set_xlim(0, 100)
        ax.set_yticks(posiciones + ancho_barra)
        ax.set_yticklabels(df_pivot.index, rotation=0, ha='right')

        fig.legend(title="Nivel de Riesgo", loc="upper center", bbox_to_anchor=(0.5, 0.96), ncol=3)
        plt.tight_layout()

        figs_te3.append((fig, te3))

    return figs_te3
