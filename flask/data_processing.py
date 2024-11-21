# data_processing.py

import pandas as pd
import numpy as np
import logging

def calculate_scores(df, df_ceal, df_risk_intervals):
    logging.info("Calculando puntajes por dimensión...")

    # Normalizar los nombres de las columnas y Codpreg
    df.columns = df.columns.str.strip().str.upper()
    df_ceal['Codpreg'] = df_ceal['Codpreg'].str.strip().str.upper()
    df_ceal['Coddim'] = df_ceal['Coddim'].str.strip().str.upper()

    # Crear los diccionarios de mapeo
    coddim_to_codpreg = df_ceal.groupby('Coddim')['Codpreg'].apply(list).to_dict()
    coddim_to_dimension = df_ceal[['Coddim', 'Dimensión']].drop_duplicates().set_index('Coddim')['Dimensión'].to_dict()

    for coddim, codpreg_list in coddim_to_codpreg.items():
        dimension = coddim_to_dimension.get(coddim)
        if not dimension:
            logging.warning(f"No se encontró la dimensión para el Coddim '{coddim}'. Skipping.")
            continue
        if dimension not in df_risk_intervals['Dimensión'].values:
            logging.info(f"Dimensión '{dimension}' no está en los intervalos de riesgo. Skipping coddim '{coddim}'.")
            continue
        # Verificar que las columnas existen en df
        missing_columns = [col for col in codpreg_list if col not in df.columns]
        if missing_columns:
            logging.warning(f"Las columnas {missing_columns} para el Coddim '{coddim}' no se encontraron en df. Skipping.")
            continue
        logging.info(f"Procesando Coddim '{coddim}' con dimensión '{dimension}'.")
        # Convertir a numérico y sumar
        df[codpreg_list] = df[codpreg_list].apply(pd.to_numeric, errors='coerce')
        df[coddim] = df[codpreg_list].sum(axis=1)
        # Calcular nivel de riesgo
        column_name = f'{coddim}_RIESGO'
        df[column_name] = df[coddim].apply(lambda x: determine_risk_level(dimension, x, df_risk_intervals))

    logging.info("Puntajes y niveles de riesgo calculados.")
    return df



def determine_risk_level(dim, score, df_risk_intervals):
    fila = df_risk_intervals[df_risk_intervals['Dimensión'] == dim]
    if fila.empty:
        # Omitir dimensiones no encontradas
        return None

    fila = fila.iloc[0]
    if fila['Nivel de riesgo bajo'][0] <= score <= fila['Nivel de riesgo bajo'][1]:
        return 'Bajo'
    elif fila['Nivel de riesgo medio'][0] <= score <= fila['Nivel de riesgo medio'][1]:
        return 'Medio'
    elif fila['Nivel de riesgo alto'][0] <= score <= fila['Nivel de riesgo alto'][1]:
        return 'Alto'
    else:
        return 'Fuera de rango'


def calcular_puntaje(row):
    nivel = row['Nivel']
    if nivel == 'Alto':
        puntaje = 3
    elif nivel == 'Medio':
        puntaje = 2
    elif nivel == 'Bajo':
        puntaje = 1
    else:
        puntaje = 0  # O el valor que consideres apropiado
        # Imprimir el puntaje calculado para depuración
        print(f"Fila {row.name}, Nivel: {nivel}, Puntaje: {puntaje}")
    return puntaje  # Devuelve un valor escalar


def create_summary(df, df_ceal, df_risk_intervals):
    logging.info("Creando resumen de puntajes y niveles de riesgo...")

    df_resultados_porcentaje = calculate_percentage_responses(df, df_ceal, df_risk_intervals)

    # Aplicar la función 'calcular_puntaje'
    df_resultados_porcentaje['Puntaje'] = df_resultados_porcentaje.apply(calcular_puntaje, axis=1)

    # Crear el resumen
    resumen = df_resultados_porcentaje.groupby('CUV').agg({'Puntaje': 'mean'}).reset_index()
    resumen['Riesgo'] = resumen['Puntaje'].apply(calculate_total_risk)

    logging.info("Resumen creado exitosamente.")
    return resumen, df_resultados_porcentaje

def calculate_total_risk(total_score):
    """
    Calcula el nivel de riesgo total basado en el puntaje acumulado.

    Parámetros:
    total_score (int): Puntaje total acumulado.

    Retorna:
    str: Nivel de riesgo total ('Riesgo bajo', 'Riesgo medio', 'Riesgo alto').
    """
    if total_score <= 1:
        return 'Riesgo bajo'
    elif 2 <= total_score <= 12:
        return 'Riesgo medio'
    else:
        return 'Riesgo alto'


def calculate_percentage_responses(df, df_ceal, df_risk_intervals):
    logging.info("Calculando porcentajes de respuestas por nivel de riesgo...")

    # Crear lista de dimensiones y columnas de riesgo
    dimensiones_riesgo = [col for col in df.columns if col.endswith('_RIESGO')]
    print("dimensiones_riesgo:", dimensiones_riesgo)

    coddim_to_dimension = {}
    for col in dimensiones_riesgo:
        coddim = col.replace('_RIESGO', '')
        dimension = df_ceal.loc[df_ceal['Coddim'] == coddim, 'Dimensión'].iloc[0]
        coddim_to_dimension[coddim] = dimension
    print("coddim_to_dimension:", coddim_to_dimension)

    resultados = []
    nivel_mapping = {'Bajo': 1, 'Medio': 2, 'Alto': 3}

    for coddim, dimension in coddim_to_dimension.items():
        columna_riesgo = f'{coddim}_RIESGO'
        if columna_riesgo not in df.columns:
            print(f"Columna '{columna_riesgo}' no encontrada en df.")
            continue
        # Filtrar solo las filas donde el nivel de riesgo no es None
        df_filtrado = df[df[columna_riesgo].notnull()]
        total_respuestas = df_filtrado.shape[0]
        print(f"Procesando dimensión '{dimension}' con {total_respuestas} respuestas.")
        for nivel in ['Bajo', 'Medio', 'Alto']:
            conteo = df_filtrado[columna_riesgo].value_counts().get(nivel, 0)
            porcentaje = round((conteo / total_respuestas) * 100, 2) if total_respuestas > 0 else 0
            resultados.append({
                'CUV': df['CUV'].unique()[0],
                'Dimensión': dimension,
                'Nivel': nivel,
                'Nivel_n': nivel_mapping[nivel],
                'Porcentaje': porcentaje,
                'Respuestas': conteo
            })

    # After the loop, print the resultados
    print("Resultados:")
    print(resultados)

    df_porcentajes = pd.DataFrame(resultados)
    logging.info("Porcentajes calculados.")
    return df_porcentajes


# Datos estáticos

ceal = [
    {"Coddim": "CT", "Dimensión": "Carga de trabajo", "Codpreg": "QD1", "Pregunta": "¿Su carga de trabajo se distribuye de manera desigual de modo que se le acumula el trabajo?"},
    {"Coddim": "CT", "Dimensión": "Carga de trabajo", "Codpreg": "QD2", "Pregunta": "¿Tiene suficiente tiempo para completar todo su trabajo?"},
    {"Coddim": "CT", "Dimensión": "Carga de trabajo", "Codpreg": "QD3", "Pregunta": "¿Tiene que trabajar muy rápido?"},
    {"Coddim": "EE", "Dimensión": "Exigencias emocionales", "Codpreg": "ED1", "Pregunta": "¿Su trabajo le exige estar en contacto con personas que sufren o están en situaciones problemáticas?"},
    {"Coddim": "EE", "Dimensión": "Exigencias emocionales", "Codpreg": "ED2", "Pregunta": "¿Su trabajo le exige enfrentarse a situaciones difíciles?"},
    {"Coddim": "EE", "Dimensión": "Exigencias emocionales", "Codpreg": "HE2", "Pregunta": "¿En su trabajo tiene que ocultar sus sentimientos?"},
    {"Coddim": "DP", "Dimensión": "Desarrollo profesional", "Codpreg": "DP2", "Pregunta": "¿Su trabajo le da la posibilidad de aprender cosas nuevas?"},
    {"Coddim": "DP", "Dimensión": "Desarrollo profesional", "Codpreg": "DP3", "Pregunta": "¿Tiene un trabajo variado?"},
    {"Coddim": "DP", "Dimensión": "Desarrollo profesional", "Codpreg": "DP4", "Pregunta": "Su trabajo, ¿le da la oportunidad de desarrollar sus habilidades?"},
    {"Coddim": "RC", "Dimensión": "Reconocimiento y claridad de rol", "Codpreg": "PR2", "Pregunta": "¿Recibe toda la información que necesita para hacer bien su trabajo?"},
    {"Coddim": "RC", "Dimensión": "Reconocimiento y claridad de rol", "Codpreg": "RE1", "Pregunta": "Su trabajo, ¿es reconocido y valorado por sus superiores?"},
    {"Coddim": "RC", "Dimensión": "Reconocimiento y claridad de rol", "Codpreg": "RE2", "Pregunta": "En su trabajo, ¿es respetado por sus superiores?"},
    {"Coddim": "RC", "Dimensión": "Reconocimiento y claridad de rol", "Codpreg": "RE3", "Pregunta": "En su trabajo, ¿es tratado de forma justa?"},
    {"Coddim": "RC", "Dimensión": "Reconocimiento y claridad de rol", "Codpreg": "MW1", "Pregunta": "Su trabajo, ¿tiene sentido para usted?"},
    {"Coddim": "RC", "Dimensión": "Reconocimiento y claridad de rol", "Codpreg": "CL1", "Pregunta": "Su trabajo, ¿tiene objetivos claros?"},
    {"Coddim": "RC", "Dimensión": "Reconocimiento y claridad de rol", "Codpreg": "CL2", "Pregunta": "En su trabajo, ¿sabe exactamente qué tareas son de su responsabilidad?"},
    {"Coddim": "RC", "Dimensión": "Reconocimiento y claridad de rol", "Codpreg": "CL3", "Pregunta": "¿Sabe exactamente lo que se espera de usted en el trabajo?"},
    {"Coddim": "CR", "Dimensión": "Conflicto de rol", "Codpreg": "CO2", "Pregunta": "En su trabajo, ¿se le exigen cosas contradictorias?"},
    {"Coddim": "CR", "Dimensión": "Conflicto de rol", "Codpreg": "CO3", "Pregunta": "¿Tiene que hacer tareas que usted cree que deberían hacerse de otra manera?"},
    {"Coddim": "CR", "Dimensión": "Conflicto de rol", "Codpreg": "IT1", "Pregunta": "¿Tiene que realizar tareas que le parecen innecesarias?"},
    {"Coddim": "QL", "Dimensión": "Calidad del liderazgo", "Codpreg": "QL3", "Pregunta": "Su superior inmediato, ¿planifica bien el trabajo?"},
    {"Coddim": "QL", "Dimensión": "Calidad del liderazgo", "Codpreg": "QL4", "Pregunta": "Su superior inmediato, ¿resuelve bien los conflictos?"},
    {"Coddim": "QL", "Dimensión": "Calidad del liderazgo", "Codpreg": "SS1", "Pregunta": "Si usted lo necesita, ¿con qué frecuencia su superior inmediato está dispuesto a escuchar sus problemas?"},
    {"Coddim": "QL", "Dimensión": "Calidad del liderazgo", "Codpreg": "SS2", "Pregunta": "Si usted lo necesita, ¿con qué frecuencia obtiene ayuda y apoyo de su superior inmediato?"},
    {"Coddim": "CM", "Dimensión": "Compañerismo", "Codpreg": "SC1", "Pregunta": "De ser necesario, ¿con qué frecuencia obtiene ayuda y apoyo de sus compañeros(as) de trabajo?"},
    {"Coddim": "CM", "Dimensión": "Compañerismo", "Codpreg": "SC2", "Pregunta": "De ser necesario, ¿con qué frecuencia sus compañeros(as) de trabajo están dispuestos(as) a escuchar problemas?"},
    {"Coddim": "CM", "Dimensión": "Compañerismo", "Codpreg": "SW1", "Pregunta": "¿Hay un buen ambiente entre usted y sus compañeros(as) de trabajo?"},
    {"Coddim": "CM", "Dimensión": "Compañerismo", "Codpreg": "SW3", "Pregunta": "En su trabajo, ¿usted siente que forma parte de un equipo?"},
    {"Coddim": "IT", "Dimensión": "Inseguridad en las condiciones de trabajo", "Codpreg": "IW1", "Pregunta": "¿Está preocupado(a) de que le cambien sus tareas laborales en contra de su voluntad?"},
    {"Coddim": "IT", "Dimensión": "Inseguridad en las condiciones de trabajo", "Codpreg": "IW2", "Pregunta": "¿Está preocupado(a) por si le trasladan a otro lugar de trabajo, obra, funciones, unidad, departamento o sección en contra de su voluntad?"},
    {"Coddim": "IT", "Dimensión": "Inseguridad en las condiciones de trabajo", "Codpreg": "IW3", "Pregunta": "¿Está preocupado(a) de que le cambien el horario (turnos, días de la semana, hora de entrada y salida) en contra de su voluntad?"},
    {"Coddim": "TV", "Dimensión": "Equilibrio trabajo y vida privada", "Codpreg": "WF2", "Pregunta": "¿Siente que su trabajo le consume demasiada ENERGÍA teniendo un efecto negativo en su vida privada?"},
    {"Coddim": "TV", "Dimensión": "Equilibrio trabajo y vida privada", "Codpreg": "WF3", "Pregunta": "¿Siente que su trabajo le consume demasiado TIEMPO teniendo un efecto negativo en su vida privada?"},
    {"Coddim": "TV", "Dimensión": "Equilibrio trabajo y vida privada", "Codpreg": "WF5", "Pregunta": "Las exigencias de su trabajo, ¿interfieren con su vida privada y familiar?"},
    {"Coddim": "CJ", "Dimensión": "Confianza y justicia organizacional", "Codpreg": "TE1", "Pregunta": "En general, ¿los trabajadores(as) en su organización confían entre sí?"},
    {"Coddim": "CJ", "Dimensión": "Confianza y justicia organizacional", "Codpreg": "TM1", "Pregunta": "¿Los gerentes o directivos confían en que los trabajadores(as) hacen bien su trabajo?"},
    {"Coddim": "CJ", "Dimensión": "Confianza y justicia organizacional", "Codpreg": "TM2", "Pregunta": "¿Los trabajadores(as) confían en la información que proviene de los gerentes, directivos o empleadores?"},
    {"Coddim": "CJ", "Dimensión": "Confianza y justicia organizacional", "Codpreg": "TM4", "Pregunta": "¿Los trabajadores(as) pueden expresar sus opiniones y sentimientos?"},
    {"Coddim": "CJ", "Dimensión": "Confianza y justicia organizacional", "Codpreg": "JU1", "Pregunta": "En su trabajo, ¿los conflictos se resuelven de manera justa?"},
    {"Coddim": "CJ", "Dimensión": "Confianza y justicia organizacional", "Codpreg": "JU2", "Pregunta": "¿Se valora a los trabajadores(as) cuando han hecho un buen trabajo?"},
    {"Coddim": "CJ", "Dimensión": "Confianza y justicia organizacional", "Codpreg": "JU4", "Pregunta": "¿Se distribuye el trabajo de manera justa?"},
    {"Coddim": "VU", "Dimensión": "Vulnerabilidad", "Codpreg": "VU1", "Pregunta": "¿Tiene miedo a pedir mejores condiciones de trabajo?"},
    {"Coddim": "VU", "Dimensión": "Vulnerabilidad", "Codpreg": "VU2", "Pregunta": "¿Se siente indefenso(a) ante el trato injusto de sus superiores?"},
    {"Coddim": "VU", "Dimensión": "Vulnerabilidad", "Codpreg": "VU3", "Pregunta": "¿Tiene miedo de que lo(la) despidan si no hace lo que le piden?"},
    {"Coddim": "VU", "Dimensión": "Vulnerabilidad", "Codpreg": "VU4", "Pregunta": "¿Considera que sus superiores lo(la) tratan de forma discriminatoria o injusta?"},
    {"Coddim": "VU", "Dimensión": "Vulnerabilidad", "Codpreg": "VU5", "Pregunta": "¿Considera que lo(la) tratan de forma autoritaria o violenta?"},
    {"Coddim": "VU", "Dimensión": "Vulnerabilidad", "Codpreg": "VU6", "Pregunta": "¿Lo(la) hacen sentir que usted puede ser fácilmente reemplazado(a)?"},
    {"Coddim": "VA", "Dimensión": "Violencia y acoso", "Codpreg": "CQ1", "Pregunta": "En su trabajo, durante los últimos 12 meses, ¿ha estado involucrado(a) en disputas o conflictos?"},
    {"Coddim": "VA", "Dimensión": "Violencia y acoso", "Codpreg": "UT1", "Pregunta": "En su trabajo, durante los últimos 12 meses, ¿ha estado expuesto(a) a bromas desagradables?"},
    {"Coddim": "VA", "Dimensión": "Violencia y acoso", "Codpreg": "HSM1", "Pregunta": "En los últimos 12 meses, ¿ha estado expuesto(a) a acoso relacionado al trabajo por correo electrónico, mensajes de texto y/o en las redes sociales (por ejemplo, Facebook, Instagram, Twitter)?"},
    {"Coddim": "VA", "Dimensión": "Violencia y acoso", "Codpreg": "SH1", "Pregunta": "En su trabajo, durante los últimos 12 meses, ¿ha estado expuesta(o) a acoso sexual?"},
    {"Coddim": "VA", "Dimensión": "Violencia y acoso", "Codpreg": "PV1", "Pregunta": "En su trabajo, en los últimos 12 meses, ¿ha estado expuesta(o) a violencia física?"},
    {"Coddim": "VA", "Dimensión": "Violencia y acoso", "Codpreg": "AL", "Pregunta": "En su trabajo, en los últimos 12 meses, ¿ha estado expuesto(a) a bullying o acoso?"},
    {"Coddim": "VA", "Dimensión": "Violencia y acoso", "Codpreg": "HO", "Pregunta": "¿Con qué frecuencia se siente intimidado(a), colocado(a) en ridículo o injustamente criticado(a), frente a otros por sus compañeros(as) de trabajo o su superior?"},
    {"Coddim": "GHQ", "Dimensión": "Cuestionario de salud general", "Codpreg": "GHQ1", "Pregunta": "¿Ha podido concentrarse bien en lo que hace?"},
    {"Coddim": "GHQ", "Dimensión": "Cuestionario de salud general", "Codpreg": "GHQ2", "Pregunta": "¿Sus preocupaciones le han hecho perder mucho sueño?"},
    {"Coddim": "GHQ", "Dimensión": "Cuestionario de salud general", "Codpreg": "GHQ3", "Pregunta": "¿Ha sentido que está jugando un papel útil en la vida?"},
    {"Coddim": "GHQ", "Dimensión": "Cuestionario de salud general", "Codpreg": "GHQ4", "Pregunta": "¿Se ha sentido capaz de tomar decisiones?"},
    {"Coddim": "GHQ", "Dimensión": "Cuestionario de salud general", "Codpreg": "GHQ5", "Pregunta": "¿Se ha sentido constantemente agobiado(a) y en tensión?"},
    {"Coddim": "GHQ", "Dimensión": "Cuestionario de salud general", "Codpreg": "GHQ6", "Pregunta": "¿Ha sentido que no puede superar sus dificultades?"},
    {"Coddim": "GHQ", "Dimensión": "Cuestionario de salud general", "Codpreg": "GHQ7", "Pregunta": "¿Ha sido capaz de disfrutar sus actividades normales de cada día?"},
    {"Coddim": "GHQ", "Dimensión": "Cuestionario de salud general", "Codpreg": "GHQ8", "Pregunta": "¿Ha sido capaz de hacer frente a sus problemas?"},
    {"Coddim": "GHQ", "Dimensión": "Cuestionario de salud general", "Codpreg": "GHQ9", "Pregunta": "¿Se ha sentido poco feliz y deprimido(a)?"},
    {"Coddim": "GHQ", "Dimensión": "Cuestionario de salud general", "Codpreg": "GHQ10", "Pregunta": "¿Ha perdido confianza en sí mismo?"},
    {"Coddim": "GHQ", "Dimensión": "Cuestionario de salud general", "Codpreg": "GHQ11", "Pregunta": "¿Ha pensado que usted es una persona que no vale para nada?"},
    {"Coddim": "GHQ", "Dimensión": "Cuestionario de salud general", "Codpreg": "GHQ12", "Pregunta": "¿Se siente razonablemente feliz considerando todas las circunstancias?"}
]

df_ceal = pd.DataFrame(ceal)

risk_intervals = [
    {"Dimensión": "Carga de trabajo", "Nivel de riesgo bajo": (0, 1), "Nivel de riesgo medio": (2, 4), "Nivel de riesgo alto": (5, 12)},
    {"Dimensión": "Exigencias emocionales", "Nivel de riesgo bajo": (0, 1), "Nivel de riesgo medio": (2, 5), "Nivel de riesgo alto": (6, 12)},
    {"Dimensión": "Desarrollo profesional", "Nivel de riesgo bajo": (0, 1), "Nivel de riesgo medio": (2, 5), "Nivel de riesgo alto": (6, 12)},
    {"Dimensión": "Reconocimiento y claridad de rol", "Nivel de riesgo bajo": (0, 4), "Nivel de riesgo medio": (5, 9), "Nivel de riesgo alto": (10, 32)},
    {"Dimensión": "Conflicto de rol", "Nivel de riesgo bajo": (0, 2), "Nivel de riesgo medio": (3, 5), "Nivel de riesgo alto": (6, 12)},
    {"Dimensión": "Calidad del liderazgo", "Nivel de riesgo bajo": (0, 2), "Nivel de riesgo medio": (3, 7), "Nivel de riesgo alto": (8, 16)},
    {"Dimensión": "Compañerismo", "Nivel de riesgo bajo": (0, 0), "Nivel de riesgo medio": (1, 4), "Nivel de riesgo alto": (5, 16)},
    {"Dimensión": "Inseguridad en las condiciones de trabajo", "Nivel de riesgo bajo": (0, 2), "Nivel de riesgo medio": (3, 5), "Nivel de riesgo alto": (6, 12)},
    {"Dimensión": "Equilibrio trabajo y vida privada", "Nivel de riesgo bajo": (0, 2), "Nivel de riesgo medio": (3, 5), "Nivel de riesgo alto": (6, 12)},
    {"Dimensión": "Confianza y justicia organizacional", "Nivel de riesgo bajo": (0, 7), "Nivel de riesgo medio": (8, 12), "Nivel de riesgo alto": (13, 28)},
    {"Dimensión": "Violencia y acoso", "Nivel de riesgo bajo": (0, 0), "Nivel de riesgo medio": (1, 14), "Nivel de riesgo alto": (15, 28)},
    {"Dimensión": "Vulnerabilidad", "Nivel de riesgo bajo": (1, 6), "Nivel de riesgo medio": (7, 11), "Nivel de riesgo alto": (12, 24)}
]

df_risk_intervals = pd.DataFrame(risk_intervals)
