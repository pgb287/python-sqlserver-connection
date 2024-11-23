import pyodbc
import csv

# Configura los parámetros de conexión
server = 'LAPTOP-PCS5TK5G'  # Puede ser 'localhost' o el nombre del servidor LAPTOP-PCS5TK5G
database = 'BDE'  # Nombre de la base de datos
#username = 'TU_USUARIO'  # Usuario de SQL Server
#password = 'TU_CONTRASEÑA'  # Contraseña del usuario

# Cadena de conexión
connection_string = (
    f"DRIVER={{ODBC Driver 17 for SQL Server}};"
    f"SERVER={server};"
    f"DATABASE={database};"
    f"Trusted_Connection=yes;"
    #f"UID={username};"
    #f"PWD={password}"
)

# Consulta
# try:
#     connection = pyodbc.connect(connection_string)
#     cursor = connection.cursor()

#     query = "select * from articulos.autores"
#     cursor.execute(query)
    
#     # Obtener todas las filas del resultado
#     rows = cursor.fetchall()
    
#     # Iterar a través de los resultados y mostrarlos
#     for row in rows:        
#         print(row)

#     # Cierra el cursor
#     cursor.close()

#     #print("Conexión exitosa a SQL Server")
# except pyodbc.Error as e:
#     print("Error en la conexión:", e)

# Datos a insertar
titulo = '10'
nombre = 'Ignacio'
apellido = 'Bustos'
telefono = '3886857111'

# Consulta de inserción

# comentar varias lineas control K control C
# descomentar varias lineas control K control U

# Insertar un registro
# try:
#     connection = pyodbc.connect(connection_string)
#     cursor = connection.cursor()
#     query = f"INSERT INTO articulos.autores (TituloId,NombreAutor,ApellidosAutor,TelefonoAutor) VALUES (?,?,?,?)"
#     cursor.execute(query,(titulo,nombre,apellido,telefono))  
#     connection.commit()  # Confirma la transacción
#     print("Inserción exitosa")
#     cursor.close()
# except pyodbc.Error as e:
#     print("Error al realizar el insert:", e)

try:
    connection = pyodbc.connect(connection_string)
    cursor = connection.cursor()

    with open('usuarios.csv', mode='r', encoding='utf-8') as file:
        csv_reader = csv.DictReader(file, delimiter=';')  # Usamos DictReader para trabajar con los nombres de las columnas
        
        for row in csv_reader:
            # Asumiendo que los nombres de las columnas del CSV coinciden con los de la base de datos
            titulo = row['titulo']
            nombre = row['nombre']
            apellido = row['apellido']
            telefono = row['telefono']
            
            # Consulta parametrizada para insertar datos en la base de datos
            query = "INSERT INTO articulos.autores (TituloId,NombreAutor,ApellidosAutor,TelefonoAutor) VALUES (?, ?, ?, ?)"
            
            # Ejecutar la consulta
            cursor.execute(query, (titulo, nombre, apellido, telefono))
        
        # Confirmar la transacción
        connection.commit()

    # Cerrar el cursor y la conexión
    cursor.close()
    connection.close()

except pyodbc.Error as e:
    print("Error al realizar el insert:", e)
