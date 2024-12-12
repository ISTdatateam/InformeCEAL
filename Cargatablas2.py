import pandas as pd
import pyodbc

server = '170.110.40.38'
database = 'ept_modprev'
username = 'usr_ept_modprev'
password = 'C(Q5N:6+5sIt'
driver = '{ODBC Driver 17 for SQL Server}'


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
        print(f"Error al conectar a la base de datos: {e}")
        return None


def drop_table(cursor, table_name):
    try:
        cursor.execute(f"IF OBJECT_ID('{table_name}', 'U') IS NOT NULL DROP TABLE {table_name};")
        cursor.commit()
        print(f"Tabla {table_name} eliminada exitosamente (si existía).")
    except pyodbc.Error as e:
        print(f"Error al eliminar la tabla {table_name}: {e}")


def create_table(cursor, table_name, df):
    # Mapeamos todo a NVARCHAR(MAX)
    columns = []
    for column_name in df.columns:
        column_name_clean = column_name.replace(' ', '_').replace('-', '_')
        columns.append(f"[{column_name_clean}] NVARCHAR(MAX)")

    columns_def = ",\n    ".join(columns)
    create_table_sql = f"""
    CREATE TABLE {table_name} (
        {columns_def}
    );
    """
    try:
        cursor.execute(create_table_sql)
        cursor.commit()
        print(f"Tabla {table_name} creada exitosamente.")
    except pyodbc.Error as e:
        print(f"Error al crear la tabla {table_name}: {e}")


def insert_data_in_chunks(cursor, table_name, df, chunk_size=10000):
    df = df.applymap(lambda x: str(x) if pd.notnull(x) else None)
    df.columns = [col.replace(' ', '_').replace('-', '_') for col in df.columns]
    columns = ", ".join([f"[{col}]" for col in df.columns])
    placeholders = ", ".join(["?" for _ in df.columns])
    insert_sql = f"INSERT INTO {table_name} ({columns}) VALUES ({placeholders})"

    # Convertir el DataFrame a una lista de tuplas
    data = [tuple(row) for row in df.itertuples(index=False, name=None)]

    try:
        cursor.fast_executemany = True
        # Insertar en lotes
        for i in range(0, len(data), chunk_size):
            chunk = data[i:i + chunk_size]
            cursor.executemany(insert_sql, chunk)
            cursor.commit()
            print(f"Insertadas {len(chunk)} filas (total {i + len(chunk)}).")
        print(f"Datos insertados en la tabla {table_name} exitosamente.")
    except pyodbc.Error as e:
        print(f"Error al insertar datos en la tabla {table_name}: {e}")


def load_excel_to_table(cursor, excel_path, table_name):
    try:
        print(f"\nLeyendo el archivo: {excel_path}")
        df = pd.read_excel(excel_path)  # Lee la primera hoja por defecto
        print("Archivo leído exitosamente.")
    except Exception as e:
        print(f"Error al leer el archivo {excel_path}: {e}")
        return

    # Elimina la tabla si existe, para asegurar nuevo esquema
    drop_table(cursor, table_name)

    # Crea la tabla siempre (NVARCHAR(MAX))
    create_table(cursor, table_name, df)

    # Inserta datos en lotes
    print(f"Insertando datos en {table_name}...")
    insert_data_in_chunks(cursor, table_name, df)
    print(f"Finalizada la carga para {table_name}.")


def main():
    connection = get_db_connection()
    if connection is None:
        return
    cursor = connection.cursor()

    files_and_tables = [
        ('Recomendaciones.xlsx', 'informe_CEAL__rec'),
        ('ciiu.xlsx', 'informe_CEAL__ciiu'),
        ('resultados.xlsx', 'informe_CEAL__fileresultados')
    ]

    for excel_file, table_name in files_and_tables:
        load_excel_to_table(cursor, excel_file, table_name)

    cursor.close()
    connection.close()
    print("\nProceso completado.")


if __name__ == "__main__":
    main()
