from flask import Flask, render_template, request, redirect, url_for
import pyodbc
from datetime import datetime

app = Flask(__name__)

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


@app.route('/')
def formulario():
    return render_template('formulario.html')


@app.route('/enviar', methods=['POST'])
def enviar():
    if request.method == 'POST':
        nombre = request.form['nombre']
        email = request.form['email']

        try:
            # Establecer la conexión con SQL Server
            conn = get_db_connection()
            cursor = conn.cursor()

            # Definir los campos de destino en la tabla de SQL Server
            campos_destino = ['nombre', 'email', 'created_at', 'updated_at']

            # Preparar valores con timestamp
            created_at = datetime.now()
            updated_at = created_at
            valores = (nombre, email, created_at, updated_at)

            # Construir consulta SQL con placeholders
            placeholders = ', '.join('?' * len(campos_destino))
            consulta_insercion = f'INSERT INTO usuarios ({", ".join(campos_destino)}) VALUES ({placeholders})'

            # Imprimir consulta y valores para depuración
            print(f"Consulta: {consulta_insercion}")
            print(f"Valores a insertar: {valores}")

            # Ejecutar la consulta
            cursor.execute(consulta_insercion, valores)
            conn.commit()

            # Cerrar cursor y conexión
            cursor.close()
            conn.close()

            return redirect(url_for('formulario'))

        except pyodbc.Error as e:
            print(f"Error al insertar los datos en SQL Server: {e}")
            return f"Error al conectar con la base de datos: {e}"


if __name__ == '__main__':
    app.run(debug=True)
