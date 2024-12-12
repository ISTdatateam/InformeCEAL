'''
folder_path = r'H:\Mi unidad\SM-CEAL\Reporteria masiva\tablas'
output_path = r'H:\Mi unidad\SM-CEAL\Reporteria masiva\salida_test.xlsx'
resultados_path = r'H:\Mi unidad\SM-CEAL\Reporteria masiva\database.xlsx'
output_archivos = r'H:\Mi unidad\SM-CEAL\Reporteria masiva'
'''


#resultados_path = r'H:\Mi unidad\SM-CEAL\database.xlsx'
output_path = r'H:\Mi unidad\SM-CEAL\salida_test.xlsx'


import pandas as pd
import logging
import pyodbc
import numpy as np


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



# Función para cargar los datos desde la tabla 'archivo_combinado'
def load_data_database():
    query = "SELECT * FROM informeCEAL_combinado WHERE RUT = '69060300-4';"
    connection = get_db_connection()

    if connection is not None:
        try:
            df = pd.read_sql(query, connection)
            return df
        except pd.io.sql.DatabaseError as e:
            st.error(f"Error al ejecutar la consulta SQL: {e}")
            return pd.DataFrame()  # Retorna un DataFrame vacío en caso de error
        finally:
            connection.close()
    else:
        return pd.DataFrame()  # Retorna un DataFrame vacío si la conexión falla


# Función para cargar los datos desde la tabla 'archivo_combinado'
def load_data_combinado():
    query = "SELECT * FROM informeCEAL_combinado"
    connection = get_db_connection()

    if connection is not None:
        try:
            df = pd.read_sql(query, connection)
            return df
        except pd.io.sql.DatabaseError as e:
            st.error(f"Error al ejecutar la consulta SQL: {e}")
            return pd.DataFrame()  # Retorna un DataFrame vacío en caso de error
        finally:
            connection.close()
    else:
        return pd.DataFrame()  # Retorna un DataFrame vacío si la conexión falla

# Cargar el DataFrame
df_res_com = load_data_database()


# Cargar el DataFrame
combined_df_base_completa = load_data_combinado()
combined_df_base_completa.rename(columns={'CdT': 'CDT_Glosa','DD1': 'Genero', 'DD2': 'Edad', 'TE1': 'CdT','TE1.1': 'TE1'}, inplace=True)
combined_df_base_completa['Genero'] = combined_df_base_completa['Genero'].replace({1: 'Hombre', 2: 'Mujer', 3: 'NcOtro', 4: 'NcOtro'})
print(combined_df_base_completa.columns.tolist())

# Función para comparar y modificar los campos
def compare_and_concat(row):
    cdt = row['CdT'].strip()
    cdt_glosa = row['CDT_Glosa'].strip()

    if cdt != cdt_glosa:
        return f"{cdt_glosa} - {cdt}"
    return cdt


# Aplicar la función a cada fila del DataFrame y eliminar la columna 'CDT_glosa'
combined_df_base_completa['CdT'] = combined_df_base_completa.apply(compare_and_concat, axis=1)
combined_df_base_completa = combined_df_base_completa.drop(columns=['CDT_Glosa'])

#Parche para considion actual de DB
for col in ['AL', 'HO']:
    if col in combined_df_base_completa.columns:
        combined_df_base_completa[col] = pd.to_numeric(combined_df_base_completa[col], errors='coerce').astype('Int64')

# Crear una nueva columna para los rangos de edad
bins = [18, 25, 36, 49, float('inf')]
labels = ['18 a 25', '26 a 36', '37 a 49', '50 o más']
combined_df_base_completa['Rango Edad'] = pd.cut(combined_df_base_completa['Edad'], bins=bins, labels=labels,
                                                 right=False)

##############

ceal = [
    {"Coddim": "CT", "Dimensión": "Carga de trabajo", "Codpreg": "QD1",
     "Pregunta": "¿Su carga de trabajo se distribuye de manera desigual de modo que se le acumula el trabajo?"},
    {"Coddim": "CT", "Dimensión": "Carga de trabajo", "Codpreg": "QD2",
     "Pregunta": "¿Con qué frecuencia le falta tiempo para completar sus tareas?"},
    {"Coddim": "CT", "Dimensión": "Carga de trabajo", "Codpreg": "QD3",
     "Pregunta": "¿Se retrasa en la entrega de su trabajo?"},
    {"Coddim": "EM", "Dimensión": "Exigencias emocionales", "Codpreg": "ED1",
     "Pregunta": "Su trabajo, ¿le coloca en situaciones emocionalmente perturbadoras?"},
    {"Coddim": "EM", "Dimensión": "Exigencias emocionales", "Codpreg": "ED2",
     "Pregunta": "Como parte de su trabajo, ¿tiene que lidiar con los problemas personales de usuarios o clientes?"},
    {"Coddim": "EM", "Dimensión": "Exigencias emocionales", "Codpreg": "HE2",
     "Pregunta": "Su trabajo, ¿le exige esconder sus emociones?"},
    {"Coddim": "DP", "Dimensión": "Desarrollo profesional", "Codpreg": "DP2",
     "Pregunta": "¿Tiene la posibilidad de adquirir nuevos conocimientos a través de su trabajo?"},
    {"Coddim": "DP", "Dimensión": "Desarrollo profesional", "Codpreg": "DP3",
     "Pregunta": "En su trabajo, ¿puede utilizar sus habilidades o experiencia?"},
    {"Coddim": "DP", "Dimensión": "Desarrollo profesional", "Codpreg": "DP4",
     "Pregunta": "Su trabajo, ¿le da la oportunidad de desarrollar sus habilidades?"},
    {"Coddim": "RC", "Dimensión": "Reconocimiento y claridad de rol", "Codpreg": "PR2",
     "Pregunta": "¿Recibe toda la información que necesita para hacer bien su trabajo?"},
    {"Coddim": "RC", "Dimensión": "Reconocimiento y claridad de rol", "Codpreg": "RE1",
     "Pregunta": "Su trabajo, ¿es reconocido y valorado por sus superiores?"},
    {"Coddim": "RC", "Dimensión": "Reconocimiento y claridad de rol", "Codpreg": "RE2",
     "Pregunta": "En su trabajo, ¿es respetado por sus superiores?"},
    {"Coddim": "RC", "Dimensión": "Reconocimiento y claridad de rol", "Codpreg": "RE3",
     "Pregunta": "En su trabajo, ¿es tratado de forma justa?"},
    {"Coddim": "RC", "Dimensión": "Reconocimiento y claridad de rol", "Codpreg": "MW1",
     "Pregunta": "Su trabajo, ¿tiene sentido para usted?"},
    {"Coddim": "RC", "Dimensión": "Reconocimiento y claridad de rol", "Codpreg": "CL1",
     "Pregunta": "Su trabajo, ¿tiene objetivos claros?"},
    {"Coddim": "RC", "Dimensión": "Reconocimiento y claridad de rol", "Codpreg": "CL2",
     "Pregunta": "En su trabajo, ¿sabe exactamente qué tareas son de su responsabilidad?"},
    {"Coddim": "RC", "Dimensión": "Reconocimiento y claridad de rol", "Codpreg": "CL3",
     "Pregunta": "¿Sabe exactamente lo que se espera de usted en el trabajo?"},
    {"Coddim": "CR", "Dimensión": "Conflicto de rol", "Codpreg": "CO2",
     "Pregunta": "En su trabajo, ¿se le exigen cosas contradictorias?"},
    {"Coddim": "CR", "Dimensión": "Conflicto de rol", "Codpreg": "CO3",
     "Pregunta": "¿Tiene que hacer tareas que usted cree que deberían hacerse de otra manera?"},
    {"Coddim": "CR", "Dimensión": "Conflicto de rol", "Codpreg": "IT1",
     "Pregunta": "¿Tiene que realizar tareas que le parecen innecesarias?"},
    {"Coddim": "QL", "Dimensión": "Calidad del liderazgo", "Codpreg": "QL3",
     "Pregunta": "Su superior inmediato, ¿planifica bien el trabajo?"},
    {"Coddim": "QL", "Dimensión": "Calidad del liderazgo", "Codpreg": "QL4",
     "Pregunta": "Su superior inmediato, ¿resuelve bien los conflictos?"},
    {"Coddim": "QL", "Dimensión": "Calidad del liderazgo", "Codpreg": "SS1",
     "Pregunta": "Si usted lo necesita, ¿con qué frecuencia su superior inmediato está dispuesto a escuchar sus problemas?"},
    {"Coddim": "QL", "Dimensión": "Calidad del liderazgo", "Codpreg": "SS2",
     "Pregunta": "Si usted lo necesita, ¿con qué frecuencia obtiene ayuda y apoyo de su superior inmediato?"},
    {"Coddim": "CM", "Dimensión": "Compañerismo", "Codpreg": "SC1",
     "Pregunta": "De ser necesario, ¿con qué frecuencia obtiene ayuda y apoyo de sus compañeros(as) de trabajo?"},
    {"Coddim": "CM", "Dimensión": "Compañerismo", "Codpreg": "SC2",
     "Pregunta": "De ser necesario, ¿con qué frecuencia sus compañeros(as) de trabajo están dispuestos(as) a escuchar problemas?"},
    {"Coddim": "CM", "Dimensión": "Compañerismo", "Codpreg": "SW1",
     "Pregunta": "¿Hay un buen ambiente entre usted y sus compañeros(as) de trabajo?"},
    {"Coddim": "CM", "Dimensión": "Compañerismo", "Codpreg": "SW3",
     "Pregunta": "En su trabajo, ¿usted siente que forma parte de un equipo?"},
    {"Coddim": "IT", "Dimensión": "Inseguridad en las condiciones de trabajo", "Codpreg": "IW1",
     "Pregunta": "¿Está preocupado(a) de que le cambien sus tareas laborales en contra de su voluntad?"},
    {"Coddim": "IT", "Dimensión": "Inseguridad en las condiciones de trabajo", "Codpreg": "IW2",
     "Pregunta": "¿Está preocupado(a) por si le trasladan a otro lugar de trabajo, obra, funciones, unidad, departamento o sección en contra de su voluntad?"},
    {"Coddim": "IT", "Dimensión": "Inseguridad en las condiciones de trabajo", "Codpreg": "IW3",
     "Pregunta": "¿Está preocupado(a) de que le cambien el horario (turnos, días de la semana, hora de entrada y salida) en contra de su voluntad?"},
    {"Coddim": "TV", "Dimensión": "Equilibrio trabajo y vida privada", "Codpreg": "WF2",
     "Pregunta": "¿Siente que su trabajo le consume demasiada ENERGÍA teniendo un efecto negativo en su vida privada?"},
    {"Coddim": "TV", "Dimensión": "Equilibrio trabajo y vida privada", "Codpreg": "WF3",
     "Pregunta": "¿Siente que su trabajo le consume demasiado TIEMPO teniendo un efecto negativo en su vida privada?"},
    {"Coddim": "TV", "Dimensión": "Equilibrio trabajo y vida privada", "Codpreg": "WF5",
     "Pregunta": "Las exigencias de su trabajo, ¿interfieren con su vida privada y familiar?"},
    {"Coddim": "CJ", "Dimensión": "Confianza y justicia organizacional", "Codpreg": "TE1",
     "Pregunta": "En general, ¿los trabajadores(as) en su organización confían entre sí?"},
    {"Coddim": "CJ", "Dimensión": "Confianza y justicia organizacional", "Codpreg": "TM1",
     "Pregunta": "¿Los gerentes o directivos confían en que los trabajadores(as) hacen bien su trabajo?"},
    {"Coddim": "CJ", "Dimensión": "Confianza y justicia organizacional", "Codpreg": "TM2",
     "Pregunta": "¿Los trabajadores(as) confían en la información que proviene de los gerentes, directivos o empleadores?"},
    {"Coddim": "CJ", "Dimensión": "Confianza y justicia organizacional", "Codpreg": "TM4",
     "Pregunta": "¿Los trabajadores(as) pueden expresar sus opiniones y sentimientos?"},
    {"Coddim": "CJ", "Dimensión": "Confianza y justicia organizacional", "Codpreg": "JU1",
     "Pregunta": "En su trabajo, ¿los conflictos se resuelven de manera justa?"},
    {"Coddim": "CJ", "Dimensión": "Confianza y justicia organizacional", "Codpreg": "JU2",
     "Pregunta": "¿Se valora a los trabajadores(as) cuando han hecho un buen trabajo?"},
    {"Coddim": "CJ", "Dimensión": "Confianza y justicia organizacional", "Codpreg": "JU4",
     "Pregunta": "¿Se distribuye el trabajo de manera justa?"},
    {"Coddim": "VU", "Dimensión": "Vulnerabilidad", "Codpreg": "VU1",
     "Pregunta": "¿Tiene miedo a pedir mejores condiciones de trabajo?"},
    {"Coddim": "VU", "Dimensión": "Vulnerabilidad", "Codpreg": "VU2",
     "Pregunta": "¿Se siente indefenso(a) ante el trato injusto de sus superiores?"},
    {"Coddim": "VU", "Dimensión": "Vulnerabilidad", "Codpreg": "VU3",
     "Pregunta": "¿Tiene miedo de que lo(la) despidan si no hace lo que le piden?"},
    {"Coddim": "VU", "Dimensión": "Vulnerabilidad", "Codpreg": "VU4",
     "Pregunta": "¿Considera que sus superiores lo(la) tratan de forma discriminatoria o injusta?"},
    {"Coddim": "VU", "Dimensión": "Vulnerabilidad", "Codpreg": "VU5",
     "Pregunta": "¿Considera que lo(la) tratan de forma autoritaria o violenta?"},
    {"Coddim": "VU", "Dimensión": "Vulnerabilidad", "Codpreg": "VU6",
     "Pregunta": "¿Lo(la) hacen sentir que usted puede ser fácilmente reemplazado(a)?"},
    {"Coddim": "VA", "Dimensión": "Violencia y acoso", "Codpreg": "CQ1",
     "Pregunta": "En su trabajo, durante los últimos 12 meses, ¿ha estado involucrado(a) en disputas o conflictos?"},
    {"Coddim": "VA", "Dimensión": "Violencia y acoso", "Codpreg": "UT1",
     "Pregunta": "En su trabajo, durante los últimos 12 meses, ¿ha estado expuesto(a) a bromas desagradables?"},
    {"Coddim": "VA", "Dimensión": "Violencia y acoso", "Codpreg": "HSM1",
     "Pregunta": "En los últimos 12 meses, ¿ha estado expuesto(a) a acoso relacionado al trabajo por correo electrónico, mensajes de texto y/o en las redes sociales (por ejemplo, Facebook, Instagram, Twitter)?"},
    {"Coddim": "VA", "Dimensión": "Violencia y acoso", "Codpreg": "SH1",
     "Pregunta": "En su trabajo, durante los últimos 12 meses, ¿ha estado expuesta(o) a acoso sexual?"},
    {"Coddim": "VA", "Dimensión": "Violencia y acoso", "Codpreg": "PV1",
     "Pregunta": "En su trabajo, en los últimos 12 meses, ¿ha estado expuesta(o) a violencia física?"},
    {"Coddim": "VA", "Dimensión": "Violencia y acoso", "Codpreg": "AL",
     "Pregunta": "En su trabajo, en los últimos 12 meses, ¿ha estado expuesto(a) a bullying o acoso?"},
    {"Coddim": "VA", "Dimensión": "Violencia y acoso", "Codpreg": "HO",
     "Pregunta": "¿Con qué frecuencia se siente intimidado(a), colocado(a) en ridículo o injustamente criticado(a), frente a otros por sus compañeros(as) de trabajo o su superior?"},
    {"Coddim": "GHQ", "Dimensión": "Cuestionario de salud general", "Codpreg": "GHQ1",
     "Pregunta": "¿Ha podido concentrarse bien en lo que hace?"},
    {"Coddim": "GHQ", "Dimensión": "Cuestionario de salud general", "Codpreg": "GHQ2",
     "Pregunta": "¿Sus preocupaciones le han hecho perder mucho sueño?"},
    {"Coddim": "GHQ", "Dimensión": "Cuestionario de salud general", "Codpreg": "GHQ3",
     "Pregunta": "¿Ha sentido que está jugando un papel útil en la vida?"},
    {"Coddim": "GHQ", "Dimensión": "Cuestionario de salud general", "Codpreg": "GHQ4",
     "Pregunta": "¿Se ha sentido capaz de tomar decisiones?"},
    {"Coddim": "GHQ", "Dimensión": "Cuestionario de salud general", "Codpreg": "GHQ5",
     "Pregunta": "¿Se ha sentido constantemente agobiado(a) y en tensión?"},
    {"Coddim": "GHQ", "Dimensión": "Cuestionario de salud general", "Codpreg": "GHQ6",
     "Pregunta": "¿Ha sentido que no puede superar sus dificultades?"},
    {"Coddim": "GHQ", "Dimensión": "Cuestionario de salud general", "Codpreg": "GHQ7",
     "Pregunta": "¿Ha sido capaz de disfrutar sus actividades normales de cada día?"},
    {"Coddim": "GHQ", "Dimensión": "Cuestionario de salud general", "Codpreg": "GHQ8",
     "Pregunta": "¿Ha sido capaz de hacer frente a sus problemas?"},
    {"Coddim": "GHQ", "Dimensión": "Cuestionario de salud general", "Codpreg": "GHQ9",
     "Pregunta": "¿Se ha sentido poco feliz y deprimido(a)?"},
    {"Coddim": "GHQ", "Dimensión": "Cuestionario de salud general", "Codpreg": "GHQ10",
     "Pregunta": "¿Ha perdido confianza en sí mismo?"},
    {"Coddim": "GHQ", "Dimensión": "Cuestionario de salud general", "Codpreg": "GHQ11",
     "Pregunta": "¿Ha pensado que usted es una persona que no vale para nada?"},
    {"Coddim": "GHQ", "Dimensión": "Cuestionario de salud general", "Codpreg": "GHQ12",
     "Pregunta": "¿Se siente razonablemente feliz considerando todas las circunstancias?"}
]

df_ceal = pd.DataFrame(ceal)

# Arreglo con la información de la tabla
risk_intervals = [
    {"Dimensión": "Carga de trabajo", "Nivel de riesgo bajo": (0, 1), "Nivel de riesgo medio": (2, 4),
     "Nivel de riesgo alto": (5, 12)},
    {"Dimensión": "Exigencias emocionales", "Nivel de riesgo bajo": (0, 1), "Nivel de riesgo medio": (2, 5),
     "Nivel de riesgo alto": (6, 12)},
    {"Dimensión": "Desarrollo profesional", "Nivel de riesgo bajo": (0, 1), "Nivel de riesgo medio": (2, 5),
     "Nivel de riesgo alto": (6, 12)},
    {"Dimensión": "Reconocimiento y claridad de rol", "Nivel de riesgo bajo": (0, 4), "Nivel de riesgo medio": (5, 9),
     "Nivel de riesgo alto": (10, 32)},
    {"Dimensión": "Conflicto de rol", "Nivel de riesgo bajo": (0, 2), "Nivel de riesgo medio": (3, 5),
     "Nivel de riesgo alto": (6, 12)},
    {"Dimensión": "Calidad del liderazgo", "Nivel de riesgo bajo": (0, 2), "Nivel de riesgo medio": (3, 7),
     "Nivel de riesgo alto": (8, 16)},
    {"Dimensión": "Compañerismo", "Nivel de riesgo bajo": (0, 0), "Nivel de riesgo medio": (1, 4),
     "Nivel de riesgo alto": (5, 16)},
    {"Dimensión": "Inseguridad en las condiciones de trabajo", "Nivel de riesgo bajo": (0, 2),
     "Nivel de riesgo medio": (3, 5), "Nivel de riesgo alto": (6, 12)},
    {"Dimensión": "Equilibrio trabajo y vida privada", "Nivel de riesgo bajo": (0, 2), "Nivel de riesgo medio": (3, 5),
     "Nivel de riesgo alto": (6, 12)},
    {"Dimensión": "Confianza y justicia organizacional", "Nivel de riesgo bajo": (0, 7),
     "Nivel de riesgo medio": (8, 12), "Nivel de riesgo alto": (13, 28)},
    {"Dimensión": "Violencia y acoso", "Nivel de riesgo bajo": (0, 0), "Nivel de riesgo medio": (1, 14),
     "Nivel de riesgo alto": (15, 28)},
    {"Dimensión": "Vulnerabilidad", "Nivel de riesgo bajo": (1, 6), "Nivel de riesgo medio": (7, 11),
     "Nivel de riesgo alto": (12, 24)}
]

# Crear el DataFrame a partir del arreglo
df_risk_intervals = pd.DataFrame(risk_intervals)

############



# Definir la función para asignar puntaje
def calcular_puntaje(row):
    if row['Dimensión'] == 'Salud mental (GHQ)':
        return None
    if row['Nivel'] == 'Bajo' and row['Porcentaje'] >= 50:
        return -2
    elif row['Nivel'] == 'Medio' and row['Porcentaje'] >= 50:
        return 1
    elif row['Nivel'] == 'Alto' and row['Porcentaje'] >= 50:
        return 2
    else:
        return None


# Calcular el Riesgo basado en el Puntaje
def calcular_riesgo(puntaje):
    if puntaje <= 1:
        return 'Riesgo bajo'
    elif 2 <= puntaje <= 12:
        return 'Riesgo medio'
    else:
        return 'Riesgo alto'


# Crear un diccionario donde las claves sean los Coddim y los valores sean listas de Codpreg
coddim_to_codpreg = df_ceal.groupby('Coddim')['Codpreg'].apply(list).to_dict()

# Crear las columnas de subtotales en df_base_completa
for coddim, codpreg_list in coddim_to_codpreg.items():
    # Sumar las columnas correspondientes a cada Codpreg y crear una nueva columna con el subtotal
    combined_df_base_completa[coddim] = combined_df_base_completa[codpreg_list].sum(axis=1)

# Crear un diccionario que mapee los Coddim a los nombres de la dimensión
coddim_to_dimension = df_ceal[['Coddim', 'Dimensión']].drop_duplicates().set_index('Coddim')['Dimensión'].to_dict()


# Función para determinar el nivel de riesgo basado en el puntaje
def determinar_nivel_riesgo(dim, puntaje):
    fila = df_risk_intervals[df_risk_intervals['Dimensión'] == dim]
    if fila.empty:
        return 'dimensión no encontrada'
    fila = fila.iloc[0]
    if fila['Nivel de riesgo bajo'][0] <= puntaje <= fila['Nivel de riesgo bajo'][1]:
        return 'Bajo'
    elif fila['Nivel de riesgo medio'][0] <= puntaje <= fila['Nivel de riesgo medio'][1]:
        return 'Medio'
    elif fila['Nivel de riesgo alto'][0] <= puntaje <= fila['Nivel de riesgo alto'][1]:
        return 'Alto'
    else:
        return 'fuera de rango'


# Aplicar la función para determinar el nivel de riesgo a cada dimensión
for coddim, dimension in coddim_to_dimension.items():
    column_name = f'{coddim}_riesgo'
    if dimension in df_risk_intervals['Dimensión'].values:
        combined_df_base_completa[column_name] = combined_df_base_completa[coddim].apply(
            lambda x: determinar_nivel_riesgo(dimension, x))
    else:
        print(f"La dimensión '{dimension}' no está presente en df_risk_intervals")

print("listo")

combined_df_base_complet2 = combined_df_base_completa


def calcular_porcentaje_respuestas_nuevas_columnas(df, nuevas_columnas):
    resultados = []
    total_respuestas = df.shape[0]
    mapping01 = {0: "No expuesto", 1: "Expuesto"}

    for columna in nuevas_columnas:
        for valor in [0, 1]:
            conteo = df[columna].value_counts().get(valor, 0)
            porcentaje = round((conteo / total_respuestas) * 100, 2)
            resultados.append({
                'CUV': df['CUV'].unique()[0],
                'CdT': df['CdT'].unique()[0],
                'Codpreg': columna,
                'Valor': valor,
                'Exposición': mapping01[valor],
                'Porcentaje': porcentaje,
                'Respuestas': conteo
            })
    return pd.DataFrame(resultados)


# Crear nuevas columnas basadas en la condición de ser diferente de 0
columnas_originales = ['CQ1', 'UT1', 'HSM1', 'SH1', 'PV1', 'AL', 'HO']
nuevas_columnas = []
for columna in columnas_originales:
    nueva_columna = columna + '_01'
    combined_df_base_complet2[nueva_columna] = combined_df_base_complet2[columna].apply(lambda x: 1 if x != 0 else 0)
    nuevas_columnas.append(nueva_columna)

# Crear una nueva columna que cuente cuántos 1 hay en las nuevas columnas
combined_df_base_complet2['Expo_total'] = combined_df_base_complet2[nuevas_columnas].sum(axis=1)
combined_df_base_complet2['Expo_total_01'] = combined_df_base_complet2['Expo_total'].apply(lambda x: 1 if x != 0 else 0)
nuevas_columnas.append('Expo_total_01')

df_exposicionviolencia = combined_df_base_complet2.groupby('CUV')[nuevas_columnas].agg(['mean'])
df_exposicionviolencia.columns = ['_'.join(col).strip() for col in df_exposicionviolencia.columns.values]
df_exposicionviolencia = df_exposicionviolencia.reset_index()


# Calcular porcentajes para las nuevas columnas
df_porcentajes = calcular_porcentaje_respuestas_nuevas_columnas(combined_df_base_complet2, nuevas_columnas)

# df_porcentajes
df_resultados_porcentaje_nuevas = pd.concat(
    [calcular_porcentaje_respuestas_nuevas_columnas(grupo, nuevas_columnas)
     for _, grupo in combined_df_base_complet2.groupby('CUV')],
    ignore_index=True
)

violencia = [
    {"Codpreg": "CQ1_01", "Temática": "Disputas o conflictos"},
    {"Codpreg": "UT1_01", "Temática": "Bromas desagradables"},
    {"Codpreg": "HSM1_01", "Temática": "Acoso virtual"},
    {"Codpreg": "SH1_01", "Temática": "Acoso sexual"},
    {"Codpreg": "PV1_01", "Temática": "Violencia física"},
    {"Codpreg": "AL_01", "Temática": "Bullying o acoso"},
    {"Codpreg": "HO_01", "Temática": "Humillaciones"},
    {"Codpreg": "Expo_total_01", "Temática": "Exposicion a violencia"}
]
# Crear el DataFrame a partir del arreglo
df_violencia = pd.DataFrame(violencia)

# Realizar un merge con el DataFrame de resultados de porcentajes
df_resultados_porcentaje_nuevas = df_resultados_porcentaje_nuevas.merge(df_violencia, on='Codpreg', how='left')

combined_df_base_complet3 = combined_df_base_complet2


def calcular_porcentaje_respuestas_nuevas_columnasp(df, nuevas_columnasn, nuevas_columnasd):
    resultados2 = []
    mapping01 = {0: "No Proteccion", 1: "Proteccion"}

    for coln, cold in zip(nuevas_columnasn, nuevas_columnasd):
        for valor in [0, 1]:
            conteon = df[coln].value_counts().get(valor, 0)
            conteod = df[cold].value_counts().get(1, 0)
            porcentaje = round((conteon / conteod) * 100, 2) if conteod != 0 else 0

            resultados2.append({
                'CUV': df['CUV'].unique()[0],
                'CdT': df['CdT'].unique()[0],
                'Codpreg': coln,
                'Exposición': mapping01[valor],  # Assuming you want to use "No proteccion" for Exposición
                'Valor': valor,
                'Porcentaje': porcentaje,
                'Respuestan': conteon,
                'Respuestad': conteod
            })

    return pd.DataFrame(resultados2)


# Definir las columnas originales
columnas_originalesp = ['SS1', 'SS2', 'SC1', 'SC2', 'SW1', 'SW3']
nuevas_columnasn = []
nuevas_columnasd = []

# Crear nuevas columnas 'n' y 'd' basadas en las reglas especificadas
for columna in columnas_originalesp:
    nueva_columnan = columna + '_n'
    nueva_columnad = columna + '_d'

    # Aplicar reglas para nueva_columnan
    combined_df_base_complet3[nueva_columnan] = combined_df_base_complet3[columna].apply(
        lambda x: 1 if x in [0, 1] else (np.nan if x == 5 else 0))
    nuevas_columnasn.append(nueva_columnan)

    # Aplicar reglas para nueva_columnad
    combined_df_base_complet3[nueva_columnad] = combined_df_base_complet3[columna].apply(
        lambda x: 1 if x in [0, 1, 2, 3, 4] else 0)
    nuevas_columnasd.append(nueva_columnad)

# Calcular porcentajes para las nuevas columnas
df_resultados_porcentaje_nuevas2 = pd.concat(
    [calcular_porcentaje_respuestas_nuevas_columnasp(grupo, nuevas_columnasn, nuevas_columnasd)
     for _, grupo in combined_df_base_complet3.groupby('CUV')],
    ignore_index=True
)

protectores = [
    {"Codpreg": "SS1_n", "Temática": "superior1"},
    {"Codpreg": "SS2_n", "Temática": "superior2"},
    {"Codpreg": "SC1_n", "Temática": "compañeros1"},
    {"Codpreg": "SC2_n", "Temática": "compañeros2"},
    {"Codpreg": "SW1_n", "Temática": "oficina1"},
    {"Codpreg": "SW3_n", "Temática": "oficina2"}
]

# Crear el DataFrame a partir del arreglo
df_protectores = pd.DataFrame(protectores)

# Realizar un merge con el DataFrame de resultados de porcentajes
df_resultados_porcentaje_nuevas2 = df_resultados_porcentaje_nuevas2.merge(df_protectores, on='Codpreg', how='left')

# Verificar el resultado
#print(df_resultados_porcentaje_nuevas2)


def calcular_porcentaje_respuestas(df, coddim_to_dimension, df_risk_intervals):
    resultados = []
    nivel_mapping = {'Bajo': 1, 'Medio': 2, 'Alto': 3}

    for coddim, dimension in coddim_to_dimension.items():
        if dimension in df_risk_intervals['Dimensión'].values:
            columna_riesgo = f'{coddim}_riesgo'
            total_respuestas = df.shape[0]
            for nivel in ['Bajo', 'Medio', 'Alto']:
                conteo = df[columna_riesgo].value_counts().get(nivel, 0)
                porcentaje = round((conteo / total_respuestas) * 100, 2)
                resultados.append({
                    'CUV': df['CUV'].unique()[0],
                    'CdT': df['CdT'].unique()[0],
                    'Dimensión': dimension,
                    'Nivel': nivel,
                    'Nivel_n': nivel_mapping[nivel],
                    'Porcentaje': porcentaje,
                    'Respuestas': conteo
                })
    return pd.DataFrame(resultados)


# Aplicar la función a cada grupo de CdT y concatenar los resultados
df_resultados_porcentaje = pd.concat([calcular_porcentaje_respuestas(grupo, coddim_to_dimension, df_risk_intervals)
                                      for _, grupo in combined_df_base_complet3.groupby('CUV')], ignore_index=True)

df_resultados_porcentaje['Puntaje'] = df_resultados_porcentaje.apply(calcular_puntaje, axis=1)

# Crear un DataFrame con los puntajes máximos por 'CdT' y 'Dimensión'
df_puntajes_max = df_resultados_porcentaje.dropna(subset=['Puntaje'])
df_puntajes_max = (df_puntajes_max.groupby(['CUV', 'Dimensión'])
                   .apply(lambda x: x.loc[x['Puntaje'].idxmax()])
                   .reset_index(drop=True))

# Hacer un merge con el DataFrame original utilizando 'CdT', 'Dimensión', 'Nivel'
df_resultados_porcentaje = df_resultados_porcentaje.merge(
    df_puntajes_max[['CUV', 'CdT', 'Dimensión', 'Nivel', 'Nivel_n', 'Porcentaje', 'Respuestas', 'Puntaje']],
    on=['CUV', 'Dimensión', 'Nivel'],
    suffixes=('', '_max'),
    how='left'
)

# Actualizar la columna 'Puntaje' con los valores de 'Puntaje_max' y eliminar 'Puntaje_max'
df_resultados_porcentaje['Puntaje'] = df_resultados_porcentaje['Puntaje_max']
df_resultados_porcentaje.drop(columns=['CdT_max', 'Puntaje_max', 'Nivel_n_max', 'Porcentaje_max', 'Respuestas_max'],
                              inplace=True)

# Asegurarse de que las columnas 'CUV' y 'CdT' sean del tipo string
df_resultados_porcentaje['CUV'] = df_resultados_porcentaje['CUV'].astype(str)
df_resultados_porcentaje['CdT'] = df_resultados_porcentaje['CdT'].astype(str)

# Agrupar por archivo y CdT para calcular Puntaje total y número de evaluaciones
summary_df = df_resultados_porcentaje.groupby(['CUV', 'CdT']).agg(
    Puntaje=('Puntaje', 'sum')
).reset_index()

summary_df['Riesgo'] = summary_df['Puntaje'].apply(calcular_riesgo)

# Reordenar las columnas en el orden deseado
summary_df = summary_df[['CUV', 'CdT', 'Puntaje', 'Riesgo']]

print("Listo")




############

def obtener_porcentaje_niveles(df, coddim_to_dimension, df_intervalos_riesgo):
    resultados = []
    nivel_mapping = {'Bajo': 1, 'Medio': 2, 'Alto': 3}

    for coddim, dimension in coddim_to_dimension.items():
        if dimension in df_intervalos_riesgo['Dimensión'].values:
            columna_riesgo = f'{coddim}_riesgo'
            total_respuestas = df.shape[0]
            for nivel in ['Bajo', 'Medio', 'Alto']:
                conteo = df[columna_riesgo].value_counts().get(nivel, 0)
                porcentaje = round((conteo / total_respuestas) * 100, 2)
                resultados.append({
                    'CUV': df['CUV'].unique()[0],
                    'CdT': df['CdT'].unique()[0],
                    'TE3': df['TE3'].unique()[0],  # Agregar TE3
                    'Dimensión': dimension,
                    'Nivel': nivel,
                    'Nivel_n': nivel_mapping[nivel],
                    'Porcentaje': porcentaje,
                    'Respuestas': conteo
                })
    return pd.DataFrame(resultados)


# Aplicar la función a cada grupo de CUV y TE3, y concatenar los resultados
df_porcentajes_niveles = pd.concat([
    obtener_porcentaje_niveles(grupo, coddim_to_dimension, df_risk_intervals)
    for _, grupo in combined_df_base_complet3.groupby(['CUV', 'TE3'])
], ignore_index=True)

df_porcentajes_niveles['Puntaje'] = df_porcentajes_niveles.apply(calcular_puntaje, axis=1)

# Crear un DataFrame con los puntajes máximos por 'CUV', 'TE3' y 'Dimensión'
df_max_puntajes = df_porcentajes_niveles.dropna(subset=['Puntaje'])
df_max_puntajes = (df_max_puntajes.groupby(['CUV', 'TE3', 'Dimensión'])
                   .apply(lambda x: x.loc[x['Puntaje'].idxmax()])
                   .reset_index(drop=True))

# Hacer un merge con el DataFrame original utilizando 'CUV', 'TE3', 'Dimensión' y 'Nivel'
df_porcentajes_niveles = df_porcentajes_niveles.merge(
    df_max_puntajes[['CUV', 'TE3', 'CdT', 'Dimensión', 'Nivel', 'Nivel_n', 'Porcentaje', 'Respuestas', 'Puntaje']],
    on=['CUV', 'TE3', 'Dimensión', 'Nivel'],
    suffixes=('', '_max'),
    how='left'
)

# Actualizar la columna 'Puntaje' con los valores de 'Puntaje_max' y eliminar 'Puntaje_max'
df_porcentajes_niveles['Puntaje'] = df_porcentajes_niveles['Puntaje_max']
df_porcentajes_niveles.drop(columns=['CdT_max', 'Puntaje_max', 'Nivel_n_max', 'Porcentaje_max', 'Respuestas_max'],
                            inplace=True)

# Asegurarse de que las columnas 'CUV', 'CdT' y 'TE3' sean del tipo string
df_porcentajes_niveles['CUV'] = df_porcentajes_niveles['CUV'].astype(str)
df_porcentajes_niveles['CdT'] = df_porcentajes_niveles['CdT'].astype(str)
df_porcentajes_niveles['TE3'] = df_porcentajes_niveles['TE3'].astype(str)

# Agrupar por archivo, CdT y TE3 para calcular Puntaje total y número de evaluaciones
df_resumen = df_porcentajes_niveles.groupby(['CUV', 'CdT', 'TE3']).agg(
    Puntaje=('Puntaje', 'sum')
).reset_index()

df_resumen['Riesgo'] = df_resumen['Puntaje'].apply(calcular_riesgo)

# Reordenar las columnas en el orden deseado
df_resumen = df_resumen[['CUV', 'CdT', 'TE3', 'Puntaje', 'Riesgo']]

#df_res_dimTE3 = df_porcentajes_niveles[df_porcentajes_niveles['Puntaje'] >= 1]
df_res_dimTE3 = df_porcentajes_niveles

############


columns_to_keep = ['CUV', 'CdT', 'QD1', 'QD2', 'QD3', 'ED1', 'ED2', 'HE2', 'DP2', 'DP3', 'DP4', 'MW1', 'PR2', 'RE1',
                   'RE2', 'RE3', 'CL1', 'CL2', 'CL3', 'CO2', 'CO3',
                   'IT1', 'QL3', 'QL4', 'SS1', 'SS2', 'SC2', 'SC1', 'SW1', 'SW3', 'IW1',
                   'IW2', 'IW3', 'WF2', 'WF3', 'WF5', 'TE1', 'TM1', 'TM2', 'TM4', 'JU1',
                   'JU2', 'JU4', 'VU1', 'VU2', 'VU3', 'VU4', 'VU5', 'VU6', 'CQ1', 'UT1',
                   'HSM1', 'SH1', 'PV1', 'AL', 'HO', 'GHQ1', 'GHQ2', 'GHQ3', 'GHQ4', 'GHQ5',
                   'GHQ6', 'GHQ7', 'GHQ8', 'GHQ9', 'GHQ10', 'GHQ11', 'GHQ12']

# Crear un nuevo DataFrame con solo las columnas seleccionadas
new_df = combined_df_base_complet3[columns_to_keep]

# Obtener las columnas a analizar
columns = new_df.columns[2:]

# Resultado final
result = []

# Recorrer cada grupo de CUV y CdT
for (CUV, CdT), group in new_df.groupby(['CUV', 'CdT']):
    # Recorrer cada columna
    for col in columns:
        # Contar la frecuencia de cada valor en la columna
        freq = group[col].value_counts().sort_index()
        # Añadir los resultados a la lista final
        for value, count in freq.items():
            result.append([CUV, CdT, col, value, count])

# Convertir la lista de resultados a un DataFrame
result_df = pd.DataFrame(result, columns=['CUV', 'CdT', 'Codpreg', 'valor', 'frec'])

# Reorganizar el DataFrame para que el formato sea correcto
result_df = result_df[['CUV', 'CdT', 'Codpreg', 'valor', 'frec']]

# Agregar columnas adicionales desde df_ceal
result_df = result_df.merge(df_ceal, on='Codpreg', how='left')

# Agregar columnas adicionales desde df_ceal
df_resultados_porcentaje = df_resultados_porcentaje.merge(df_ceal[['Dimensión', 'Coddim']], on='Dimensión', how='left')
df_resultados_porcentaje = df_resultados_porcentaje.drop_duplicates(subset=None)

# Leer 'resultados.xlsx'
#df_resultados = pd.read_excel(resultados_path, sheet_name='Datos', usecols=['CUV', 'Folio'])
#df_res_com = pd.read_excel(resultados_path, sheet_name='Datos')

df_resultados = df_res_com[['CUV', 'Folio']]


# Convertir 'CUV' a int64 en df_resultados
df_resultados['CUV'] = df_resultados['CUV'].astype('int64')

# Lista de DataFrames y sus nombres
dataframes = [combined_df_base_complet3, result_df, df_resultados_porcentaje, df_resultados_porcentaje_nuevas, df_resultados_porcentaje_nuevas2, df_exposicionviolencia,  summary_df]
nombres = ['combined_df_base_complet3', 'result_df', 'df_resultados_porcentaje', 'df_resultados_porcentaje_nuevas', 'df_resultados_porcentaje_nuevas2', 'df_exposicionviolencia','summary_df']

# Función para convertir 'CUV' a int64 y manejar errores
def convert_cuv_to_int64(df, nombre_df):
    if 'CUV' in df.columns:
        df['CUV'] = pd.to_numeric(df['CUV'], errors='coerce')
        if df['CUV'].isnull().any():
            print(f"Valores no numéricos encontrados en 'CUV' de {nombre_df}.")
            df = df.dropna(subset=['CUV'])
        df['CUV'] = df['CUV'].astype('int64')
    return df

# Aplicar la función a cada DataFrame
for df, nombre in zip(dataframes, nombres):
    df = convert_cuv_to_int64(df, nombre)

# Realizar el merge
def merge_if_cuv_exists(df):
    if 'CUV' in df.columns:
        return df.merge(df_resultados, on='CUV', how='left')
    else:
        return df

# Aplicar el merge a cada DataFrame
combined_df_base_complet3 = merge_if_cuv_exists(combined_df_base_complet3)
result_df = merge_if_cuv_exists(result_df)
df_resultados_porcentaje = merge_if_cuv_exists(df_resultados_porcentaje)
df_resultados_porcentaje_nuevas = merge_if_cuv_exists(df_resultados_porcentaje_nuevas)
df_resultados_porcentaje_nuevas2 = merge_if_cuv_exists(df_resultados_porcentaje_nuevas2)
df_exposicionviolencia= merge_if_cuv_exists(df_exposicionviolencia)
summary_df = merge_if_cuv_exists(summary_df)


df_resultados_porcentaje['Descripción'] = df_resultados_porcentaje.apply(
    lambda row: f"{row['Dimensión']} ({row['Porcentaje']}% Riesgo {row['Nivel']}, {row['Respuestas']} personas)"
    if row['Puntaje'] in [1, 2] else "",
    axis=1
)

df_porcentajes_niveles['Descripción'] = df_porcentajes_niveles.apply(
    lambda row: f"{row['Dimensión']} ({row['Porcentaje']}% Riesgo {row['Nivel']}, {row['Respuestas']} personas)"
    if row['Puntaje'] in [1, 2] else "",
    axis=1
)

# Calcular la columna 'Factor' como la multiplicación de 'valor' por 'frec'
result_df['Factor'] = result_df['valor'] * result_df['frec']
# Agrupar por 'CUV', 'Dimensión', y 'Pregunta' y sumar los valores de 'Factor' en cada grupo
result_df = result_df.groupby(['CUV', 'Dimensión', 'Pregunta'], as_index=False)['Factor'].sum()
# Para cada combinación de 'CUV' y 'Dimensión', seleccionar la pregunta con el mayor 'Factor'
top_glosas = result_df.sort_values(by='Factor', ascending=False).groupby(['CUV', 'Dimensión']).head(2)


#print(top_glosas)


# Escribir en 'combined_output.xlsx'
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    combined_df_base_complet3.to_excel(writer, sheet_name='basecompleta', index=False)
    result_df.to_excel(writer, sheet_name='recuentopreguntas', index=False)
    top_glosas.to_excel(writer, sheet_name='top_glosas', index=False)
    df_resultados_porcentaje.to_excel(writer, sheet_name='resultado', index=False)
    df_resultados_porcentaje_nuevas.to_excel(writer, sheet_name='violencia', index=False)
    df_resultados_porcentaje_nuevas2.to_excel(writer, sheet_name='protectores', index=False)
    df_exposicionviolencia.to_excel(writer, sheet_name='expoviolencia', index=False)
    summary_df.to_excel(writer, sheet_name='Summary', index=False)
    df_porcentajes_niveles.to_excel(writer, sheet_name='df_porcentajes_niveles', index=False)
    df_res_dimTE3.to_excel(writer, sheet_name='df_res_dimTE3', index=False)
    df_resumen.to_excel(writer, sheet_name='df_resumen', index=False)

print("exito")