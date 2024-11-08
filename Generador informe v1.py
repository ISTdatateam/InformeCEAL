import pypandoc
import matplotlib.pyplot as plt
import io


# Función para generar el contenido en formato Markdown
def generar_contenido_markdown(datos, medidas):
    contenido = f"""
# INFORME DE IMPLEMENTACIÓN

## PROTOCOLO DE VIGILANCIA DE RIESGOS PSICOSOCIALES EN EL TRABAJO

**Razón Social**: {datos['razon_social']}  
**RUT**: {datos['rut']}  
**Nombre del centro de trabajo**: {datos['nombre_centro']}  
**Dirección**: {datos['direccion']}  
**Dotación**: {datos['dotacion']}  
**Porcentaje de participación**: {datos['participacion']}%  
**Fecha constitución del comité**: {datos['fecha_comite']}  
**Período de aplicación**: {datos['periodo_aplicacion']}  
**Estado de Riesgo**: {datos['estado_riesgo']}  

## RESULTADOS CUESTIONARIO CEAL-SM SUSESO
Datos del cuestionario con resultados dummy.

## LISTADO DE MEDIDAS
"""
    for medida in medidas:
        contenido += f"""
**MEDIDA**: {medida['medida']}  
**Dimensión(es)**: {medida['dimensiones']}  
**Origen del Riesgo**: {medida['origen_riesgo']}  
**Alcance (GES)**: {medida['alcance']}  
**Plazo**: {medida['plazo']}  
**Responsable**: {medida['responsable']}  

---
"""
    return contenido


# Función para guardar el gráfico como una imagen temporal
def guardar_grafico():
    plt.figure(figsize=(6, 4))
    plt.bar(["Dimensión A", "Dimensión B", "Dimensión C"], [70, 85, 60])
    plt.title("Resultados de Dimensiones")
    plt.xlabel("Dimensiones")
    plt.ylabel("Puntaje")

    # Guardar el gráfico en un archivo temporal
    plt.savefig("grafico_resultado.png", format="png")
    plt.close()


# Datos de ejemplo
datos_encabezado = {
    "razon_social": "Summit Motors S.A.",
    "rut": "76.027.979-K",
    "nombre_centro": "Antofagasta Ónix",
    "direccion": "Ónix N°35, Antofagasta",
    "dotacion": 24,
    "participacion": 79,
    "fecha_comite": "12-06-2024",
    "periodo_aplicacion": "08-07-2024 al 07-08-2024",
    "estado_riesgo": "Medio"
}

medidas = [
    {
        "medida": "Revisar y definir puestos de trabajo clave para el desarrollo de la operación y resguardar su adecuado desenvolvimiento en el lugar.",
        "dimensiones": "Carga de trabajo",
        "origen_riesgo": "Debido a la falta de puesto de trabajo clave (junior) los vendedores realizan labores extras.",
        "alcance": "Ventas",
        "plazo": "Corto plazo (180 días)",
        "responsable": "Gerencia de área"
    },
    # Puedes añadir más medidas aquí
]

# Generar contenido en Markdown
contenido_markdown = generar_contenido_markdown(datos_encabezado, medidas)

# Guardar el gráfico en la carpeta actual
guardar_grafico()

# Agregar el gráfico al documento
contenido_markdown += "\n\n![Gráfico de Resultados](grafico_resultado.png)\n"

# Convertir el contenido Markdown a Word
output_path = "informe_generado.docx"
pypandoc.convert_text(contenido_markdown, 'docx', format='md', outputfile=output_path)

print("Informe generado con éxito en formato Word.")
