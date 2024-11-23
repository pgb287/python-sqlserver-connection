import pyodbc
import csv

# Configura los parámetros de conexión
server = 'LAPTOP-PCS5TK5G'  
database = 'BDE'  # Nombre de la base de datos
#username = 'TU_USUARIO'  # Usuario de SQL Server
#password = 'TU_CONTRASEÑA'  # Contraseña del usuario

# Cadena de conexión. Si es autenticacion de windows se usa Trusted_Connection
connection_string = (
    f"DRIVER={{ODBC Driver 17 for SQL Server}};"
    f"SERVER={server};"
    f"DATABASE={database};"
    f"Trusted_Connection=yes;"
    #f"UID={username};"
    #f"PWD={password}"
)

try:
    connection = pyodbc.connect(connection_string)
    cursor = connection.cursor()

    with open('usuarios.csv', mode='r', encoding='utf-8') as file:
        csv_reader = csv.DictReader(file, delimiter=';')  # Usamos DictReader para trabajar con los nombres de las columnas
        
        for row in csv_reader:
            # row['nombre_de_la_columna_en_el_csv']
            titulo = row['titulo']
            nombre = row['nombre']
            apellido = row['apellido']
            telefono = row['telefono']
            
            # Consulta parametrizada para insertar datos en la base de datos
            query = "INSERT INTO articulos.autores (TituloId,NombreAutor,ApellidosAutor,TelefonoAutor) VALUES (?, ?, ?, ?)"
            
            # Ejecutar la consulta, query y tupla con las variables
            cursor.execute(query, (titulo, nombre, apellido, telefono))
        
        # Confirmar la transacción. Conviene realizar al final del recorrido
        connection.commit()

    # Cerrar el cursor y la conexión
    cursor.close()
    connection.close()
    print("Se realizó el INSERT en la BD exitosamente")

except pyodbc.Error as e:
    print("Error al realizar el insert:", e)
