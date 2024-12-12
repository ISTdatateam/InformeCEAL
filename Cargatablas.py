import pandas as pd
import pyodbc
import os

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


def table_exists(cursor, table_name):
    check_sql = f"SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '{table_name}'"
    cursor.execute(check_sql)
    return cursor.fetchone() is not None


def drop_table(cursor, table_name):
    try:
        cursor.execute(f"IF OBJECT_ID('{table_name}', 'U') IS NOT NULL DROP TABLE {table_name};")
        cursor.commit()
        print(f"Tabla {table_name} eliminada exitosamente (si existía).")
    except pyodbc.Error as e:
        print(f"Error al eliminar la tabla {table_name}: {e}")


def create_table(cursor, table_name, df):
    # Mapear todo a NVARCHAR(MAX)
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


def insert_data(cursor, table_name, df):
    # Convertir todo a string
    df = df.applymap(lambda x: str(x) if pd.notnull(x) else None)
    df.columns = [col.replace(' ', '_').replace('-', '_') for col in df.columns]
    columns = ", ".join([f"[{col}]" for col in df.columns])
    placeholders = ", ".join(["?" for _ in df.columns])
    insert_sql = f"INSERT INTO {table_name} ({columns}) VALUES ({placeholders})"

    try:
        data = [tuple(row) for row in df.itertuples(index=False, name=None)]
        cursor.fast_executemany = True
        cursor.executemany(insert_sql, data)
        cursor.commit()
        print(f"Datos insertados en la tabla {table_name} exitosamente.")
    except pyodbc.Error as e:
        print(f"Error al insertar datos en la tabla {table_name}: {e}")


def main():
    excel_path = r'H:\Mi unidad\SM-CEAL\salida_test.xlsx'
    sheet_names = [
        'basecompleta',
        'Summary',
        'resultado',
        'df_porcentajes_niveles',
        'df_res_dimTE3',
        'df_resumen',
        'top_glosas'
    ]

    try:
        print("Leyendo el archivo Excel...")
        excel_data = pd.read_excel(excel_path, sheet_name=sheet_names)
        print("Archivo Excel leído exitosamente.")
    except Exception as e:
        print(f"Error al leer el archivo Excel: {e}")
        excel_data = {}

    connection = get_db_connection()
    if connection is None:
        return

    cursor = connection.cursor()

    for sheet_name, df in excel_data.items():
        print(f"\nProcesando hoja: {sheet_name}")
        table_name = f"informe_CEAL__{sheet_name}"

        # Elimina la tabla si existe, para asegurar nuevo esquema
        drop_table(cursor, table_name)

        # Crea la tabla siempre (NVARCHAR(MAX))
        create_table(cursor, table_name, df)

        # Inserta datos
        print(f"Insertando datos en {table_name}...")
        insert_data(cursor, table_name, df)
        print(f"Finalizada la carga para {table_name}.")

    cursor.close()
    connection.close()
    print("\nProceso completado.")


if __name__ == "__main__":
    main()
