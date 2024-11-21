import streamlit as st
import pyodbc
from datetime import datetime

# Configuración de la base de datos para SQL Server
server = '170.110.40.38'
database = 'ept_modprev'
username = 'usr_ept_modprev'
password = 'C(Q5N:6+5sIt'
driver = '{ODBC Driver 17 for SQL Server}'


def get_db_connection():
    return pyodbc.connect(
        f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}'
    )


st.title('Medidas Propuestas para Home Farm Rengo')

# Datos iniciales
dimensiones = [
    {
        'dimension': 'Carga de trabajo',
        'riesgo': '60.0% Riesgo Alto, 6 personas',
        'area': 'Administración',
        'preguntas': [
            '¿Su carga de trabajo se distribuye de manera desigual de modo que se le acumula el trabajo?',
            '¿Con qué frecuencia le falta tiempo para completar sus tareas?'
        ],
        'medidas': [
            '- Involucrar a los trabajadores en el rediseño de tareas, considerando la estacionalidad y cambios climáticos.',
            '- Ajustar tareas a la capacidad física y experiencia, revisando la carga estacionalmente.'
        ]
    }
]

# Inicializar el estado de sesión
if 'dimensiones' not in st.session_state:
    st.session_state.dimensiones = dimensiones

for idx, dim in enumerate(st.session_state.dimensiones):
    st.header(f"Dimensión en riesgo: **{dim['dimension']}**")
    st.subheader(f"{dim['riesgo']} en {dim['area']}")

    st.write("**Preguntas clave:**")
    for pregunta in dim['preguntas']:
        st.write(f"- {pregunta}")

    # Campo para la explicación
    dim['explicacion'] = st.text_area(f"Explicación para {dim['dimension']}", key=f"explicacion_{idx}")

    # Manejar las medidas propuestas
    if f"medidas_{idx}" not in st.session_state:
        st.session_state[f"medidas_{idx}"] = dim['medidas']

    st.write("**Medidas propuestas:**")
    medidas = st.session_state[f"medidas_{idx}"]
    for i in range(len(medidas)):
        medidas[i] = st.text_area(f"Medida {i + 1}", value=medidas[i], key=f"medida_{idx}_{i}")
        fecha = st.selectbox(
            f"Fecha monitoreo para medida {i + 1}",
            options=['01/03/2024 (Corto Plazo)', '01/06/2024 (Mediano Plazo)', '01/09/2024 (Largo Plazo)'],
            key=f"fecha_{idx}_{i}"
        )
        responsable = st.text_input(f"Responsable seguimiento para medida {i + 1}", key=f"responsable_{idx}_{i}")
        # Botón para eliminar la medida
        if st.button(f"Eliminar medida {i + 1}", key=f"eliminar_medida_{idx}_{i}"):
            medidas.pop(i)
            st.experimental_rerun()

    # Botón para agregar una nueva medida
    if st.button(f"Agregar nueva medida para {dim['dimension']}", key=f"agregar_medida_{idx}"):
        medidas.append('')
        st.experimental_rerun()

# Datos de usuario
st.write("## Datos de Usuario")
nombre = st.text_input('Nombre')
email = st.text_input('Email')

# Botón para enviar los datos
if st.button("Enviar datos"):
    try:
        # Establecer la conexión con SQL Server
        conn = get_db_connection()
        cursor = conn.cursor()

        # Insertar datos del usuario
        created_at = datetime.now()
        updated_at = created_at
        valores_usuario = (nombre, email, created_at, updated_at)
        consulta_usuario = 'INSERT INTO usuarios (nombre, email, created_at, updated_at) VALUES (?, ?, ?, ?)'
        cursor.execute(consulta_usuario, valores_usuario)

        # Obtener el ID del usuario insertado
        usuario_id = cursor.execute("SELECT SCOPE_IDENTITY()").fetchone()[0]

        # Insertar medidas propuestas
        for idx, dim in enumerate(st.session_state.dimensiones):
            dimension = dim['dimension']
            explicacion = dim['explicacion']
            for i, medida in enumerate(st.session_state[f"medidas_{idx}"]):
                fecha_monitoreo = st.session_state[f"fecha_{idx}_{i}"]
                responsable = st.session_state[f"responsable_{idx}_{i}"]
                valores_medida = (usuario_id, dimension, medida, fecha_monitoreo, responsable, explicacion)
                consulta_medida = '''
                INSERT INTO medidas (usuario_id, dimension, medida, fecha_monitoreo, responsable, explicacion)
                VALUES (?, ?, ?, ?, ?, ?)
                '''
                cursor.execute(consulta_medida, valores_medida)

        conn.commit()
        cursor.close()
        conn.close()
        st.success("Datos enviados exitosamente.")
    except pyodbc.Error as e:
        st.error(f"Error al insertar los datos en SQL Server: {e}")
