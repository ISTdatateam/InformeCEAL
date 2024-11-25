import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

# Datos iniciales
data = {
    "Nombre": ["Juan", "María", "Carlos"],
    "Edad": [25, 30, 35],
    "Ciudad": ["Santiago", "Valparaíso", "Concepción"]
}
df = pd.DataFrame(data)

# Función para agregar nueva fila
def agregar_fila(dataframe):
    nueva_fila = {"Nombre": "", "Edad": 0, "Ciudad": ""}
    return dataframe.append(nueva_fila, ignore_index=True)

# Función para eliminar fila seleccionada
def eliminar_filas(dataframe, indices):
    return dataframe.drop(indices).reset_index(drop=True)

# Configuración de la tabla editable
gb = GridOptionsBuilder.from_dataframe(df)
gb.configure_default_column(editable=True, resizable=True)
gb.configure_selection(selection_mode="multiple", use_checkbox=True)
grid_options = gb.build()

# Tabla editable
response = AgGrid(
    df,
    gridOptions=grid_options,
    update_mode=GridUpdateMode.MANUAL,
    editable=True,
    fit_columns_on_grid_load=True,
    height=300,
    reload_data=False
)

# Obtener datos editados
df_actualizado = response["data"]
indices_seleccionados = response["selected_rows"]

# Botones para agregar/eliminar filas
col1, col2 = st.columns(2)

with col1:
    if st.button("Agregar fila"):
        df_actualizado = agregar_fila(df_actualizado)

with col2:
    if st.button("Eliminar filas seleccionadas"):
        if indices_seleccionados:  # Validación
            indices = [row['_selectedRowNodeInfo']['nodeRowIndex'] for row in indices_seleccionados]
            df_actualizado = eliminar_filas(df_actualizado, indices)
        else:
            st.warning("No se han seleccionado filas para eliminar.")

# Mostrar datos actualizados
st.subheader("Datos Actualizados")
st.dataframe(df_actualizado)

# Guardar resultados finales
st.download_button(
    label="Descargar datos",
    data=df_actualizado.to_csv(index=False),
    file_name="datos_actualizados.csv",
    mime="text/csv"
)
