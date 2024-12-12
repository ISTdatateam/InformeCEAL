import streamlit as st
from datetime import date

# Definir las secciones por código (ya existente)
secciones_por_codigo = {
    "1234": [
        "Información Personal",
        "Detalles del Producto",
        "Confirmación"
    ],
    "2355": [
        "Información de Empresa",
        "Confirmación"
    ],
    "3456": [
        "Información Personal",
        "Detalles del Producto",
        "Información Adicional",
        "Confirmación"
    ],
    "4567": [
        "Información de Proyecto",
        "Detalles del Proyecto",
        "Confirmación"
    ]
}


# Inicialización de st.session_state (ejemplo)
def inicializar_session_state():
    default_values = {
        "nombre": "",
        "apellido": "",
        "edad": 25,
        "genero": "Masculino",
        "producto": "",
        "cantidad": 1,
        "categoria": "Electrónica",
        "fecha_entrega": date.today(),
        "empresa_nombre": "",
        "empresa_industria": "",
        "info_extra": "",
        "proyecto_nombre": "",
        "proyecto_descripcion": "",
        "proyecto_detalle": ""
    }
    for key, default in default_values.items():
        if key not in st.session_state:
            st.session_state[key] = default


inicializar_session_state()

st.title("Formulario Multisección Lateral Dinámica")

# Barra Lateral
st.sidebar.header("Menú de Navegación")

# Barra de Búsqueda en el Sidebar
st.sidebar.subheader("Buscar Código")
opciones_busqueda = list(secciones_por_codigo.keys())
codigo_seleccionado = st.sidebar.selectbox("Selecciona un código:", opciones_busqueda)

# Contenido Dinámico Basado en la Selección
st.sidebar.markdown("---")
st.sidebar.subheader("Detalles del Código")

detalles_codigo = {
    "1234": {
        "Descripción": "Este código corresponde al producto XYZ.",
        "Precio": "$100",
        "Disponibilidad": "En stock"
    },
    "2355": {
        "Descripción": "Este código corresponde al servicio ABC.",
        "Precio": "$200",
        "Disponibilidad": "Bajo pedido"
    },
    "3456": {
        "Descripción": "Este código corresponde al producto LMN.",
        "Precio": "$150",
        "Disponibilidad": "En stock"
    },
    "4567": {
        "Descripción": "Este código corresponde al proyecto DEF.",
        "Precio": "$300",
        "Disponibilidad": "En desarrollo"
    }
}

if codigo_seleccionado in detalles_codigo:
    st.sidebar.write(f"### Código {codigo_seleccionado}")
    st.sidebar.write(f"**Descripción:** {detalles_codigo[codigo_seleccionado]['Descripción']}")
    st.sidebar.write(f"**Precio:** {detalles_codigo[codigo_seleccionado]['Precio']}")
    st.sidebar.write(f"**Disponibilidad:** {detalles_codigo[codigo_seleccionado]['Disponibilidad']}")
else:
    st.sidebar.write("No hay detalles disponibles para este código.")

st.sidebar.markdown("---")
st.sidebar.write("© 2024 Tu Empresa")

# Obtiene las secciones actuales según el código
secciones_actuales = secciones_por_codigo.get(codigo_seleccionado, ["Confirmación"])

# Mover la navegación de secciones a la barra lateral
seccion_formulario = st.sidebar.radio("Navega entre las secciones del formulario:", secciones_actuales)

# Barra de Progreso en la barra lateral
total_secciones = len(secciones_actuales)
indice_seccion = secciones_actuales.index(seccion_formulario) + 1
st.sidebar.progress(indice_seccion / total_secciones)


# Definición de funciones para las secciones (ejemplo simplificado)
def informacion_personal():
    st.header("Información Personal")
    with st.form("form_informacion_personal"):
        nombre = st.text_input("Nombre", value=st.session_state["nombre"])
        apellido = st.text_input("Apellido", value=st.session_state["apellido"])
        edad = st.number_input("Edad", min_value=0, max_value=120, value=st.session_state["edad"])
        genero = st.selectbox("Género", options=["Masculino", "Femenino", "Otro"],
                              index=["Masculino", "Femenino", "Otro"].index(st.session_state["genero"]))
        submitted = st.form_submit_button("Guardar Información Personal")
        if submitted:
            if not nombre.strip() or not apellido.strip():
                st.error("Por favor, completa todos los campos obligatorios.")
            else:
                st.session_state["nombre"] = nombre
                st.session_state["apellido"] = apellido
                st.session_state["edad"] = edad
                st.session_state["genero"] = genero
                st.success("Información personal guardada.")


def detalles_producto():
    st.header("Detalles del Producto")
    with st.form("form_detalles_producto"):
        producto = st.text_input("Nombre del Producto", value=st.session_state["producto"])
        cantidad = st.number_input("Cantidad", min_value=1, max_value=1000, value=st.session_state["cantidad"])
        categoria = st.selectbox("Categoría", options=["Electrónica", "Ropa", "Alimentos", "Otros"],
                                 index=["Electrónica", "Ropa", "Alimentos", "Otros"].index(
                                     st.session_state["categoria"]))
        fecha_entrega = st.date_input("Fecha de Entrega", value=st.session_state["fecha_entrega"])
        submitted = st.form_submit_button("Guardar Detalles del Producto")
        if submitted:
            if not producto.strip():
                st.error("Por favor, ingresa el nombre del producto.")
            else:
                st.session_state["producto"] = producto
                st.session_state["cantidad"] = cantidad
                st.session_state["categoria"] = categoria
                st.session_state["fecha_entrega"] = fecha_entrega
                st.success("Detalles del producto guardados.")


def informacion_empresa():
    st.header("Información de Empresa")
    with st.form("form_informacion_empresa"):
        empresa_nombre = st.text_input("Nombre de la Empresa", value=st.session_state["empresa_nombre"])
        empresa_industria = st.text_input("Industria", value=st.session_state["empresa_industria"])
        submitted = st.form_submit_button("Guardar Información de Empresa")
        if submitted:
            if not empresa_nombre.strip() or not empresa_industria.strip():
                st.error("Por favor, completa todos los campos obligatorios.")
            else:
                st.session_state["empresa_nombre"] = empresa_nombre
                st.session_state["empresa_industria"] = empresa_industria
                st.success("Información de empresa guardada.")


def informacion_adicional():
    st.header("Información Adicional")
    with st.form("form_informacion_adicional"):
        info_extra = st.text_area("Información Adicional", value=st.session_state["info_extra"])
        submitted = st.form_submit_button("Guardar Información Adicional")
        if submitted:
            st.session_state["info_extra"] = info_extra
            st.success("Información adicional guardada.")


def informacion_proyecto():
    st.header("Información de Proyecto")
    with st.form("form_informacion_proyecto"):
        proyecto_nombre = st.text_input("Nombre del Proyecto", value=st.session_state["proyecto_nombre"])
        proyecto_descripcion = st.text_area("Descripción del Proyecto", value=st.session_state["proyecto_descripcion"])
        submitted = st.form_submit_button("Guardar Información de Proyecto")
        if submitted:
            if not proyecto_nombre.strip() or not proyecto_descripcion.strip():
                st.error("Por favor, completa todos los campos obligatorios.")
            else:
                st.session_state["proyecto_nombre"] = proyecto_nombre
                st.session_state["proyecto_descripcion"] = proyecto_descripcion
                st.success("Información de proyecto guardada.")


def detalles_proyecto():
    st.header("Detalles del Proyecto")
    with st.form("form_detalles_proyecto"):
        proyecto_detalle = st.text_input("Detalle del Proyecto", value=st.session_state["proyecto_detalle"])
        submitted = st.form_submit_button("Guardar Detalles del Proyecto")
        if submitted:
            st.session_state["proyecto_detalle"] = proyecto_detalle
            st.success("Detalles del proyecto guardados.")


def confirmacion():
    st.header("Confirmación")
    st.subheader("Revisa tus datos antes de enviar:")

    # Mostrar datos según lo llenado
    if st.session_state["nombre"]:
        st.write("**Información Personal:**")
        st.write(f"Nombre: {st.session_state['nombre']}")
        st.write(f"Apellido: {st.session_state['apellido']}")
        st.write(f"Edad: {st.session_state['edad']}")
        st.write(f"Género: {st.session_state['genero']}")
        st.write("---")

    if st.session_state["empresa_nombre"]:
        st.write("**Información de Empresa:**")
        st.write(f"Nombre de la Empresa: {st.session_state['empresa_nombre']}")
        st.write(f"Industria: {st.session_state['empresa_industria']}")
        st.write("---")

    if st.session_state["proyecto_nombre"]:
        st.write("**Información de Proyecto:**")
        st.write(f"Nombre del Proyecto: {st.session_state['proyecto_nombre']}")
        st.write(f"Descripción del Proyecto: {st.session_state['proyecto_descripcion']}")
        st.write(f"Detalle del Proyecto: {st.session_state['proyecto_detalle']}")
        st.write("---")

    if st.session_state["producto"]:
        st.write("**Detalles del Producto:**")
        st.write(f"Producto: {st.session_state['producto']}")
        st.write(f"Cantidad: {st.session_state['cantidad']}")
        st.write(f"Categoría: {st.session_state['categoria']}")
        st.write(f"Fecha de Entrega: {st.session_state['fecha_entrega']}")
        st.write("---")

    if st.session_state["info_extra"]:
        st.write("**Información Adicional:**")
        st.write(st.session_state["info_extra"])
        st.write("---")

    if st.button("Enviar"):
        # Validaciones aquí si se requiere
        st.success("¡Formulario enviado exitosamente!")


# Mapeo de secciones a funciones
secciones_funciones = {
    "Información Personal": informacion_personal,
    "Detalles del Producto": detalles_producto,
    "Información de Empresa": informacion_empresa,
    "Información Adicional": informacion_adicional,
    "Información de Proyecto": informacion_proyecto,
    "Detalles del Proyecto": detalles_proyecto,
    "Confirmación": confirmacion
}

# Renderizar la sección seleccionada en el main area
if seccion_formulario in secciones_funciones:
    secciones_funciones[seccion_formulario]()
else:
    st.error("Sección no encontrada.")
