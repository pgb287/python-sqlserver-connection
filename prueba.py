import pyodbc
import csv

# Configura los parámetros de conexión
#server = 'vegadpi1\\sqlserver2012'
server = 'LAPTOP-PCS5TK5G'
database = 'sicvaf'
username = 'bpablo'
password = 'bpablo'


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

    cod = 'A'
    padron = '12285'
    next_code = 1
    valor_unitario = {"A":833.27,"B":640.83, "C":506.98, "D":316.46,"E":175.69}
    e1cat = 'C'                
    #------------------------------------------------------------------------                        
    # Insertar datos en tabla TE1R1
    query1 = "INSERT INTO TE1R1 (DCodDep,T0Padron,E1Cod,E11A,E11B,E11C,E11D,E11E,E12A,E12B,E12C,E12D,E12E,E13A,E13B,E13C,E13D,E13E,E14A,E14B,E14C,E14D,E14E,E15A,E15B,E15C,E15D,E15E,E16A,E16B,E16C,E16D,E16E,E17A,E17B,E17C,E17D,E17E,E18A,E18B,E18C,E18D,E18E,E19A,E19B,E19C,E19D,E19E,E110A,E110B,E110C,E110D,E110E,E111A,E111B,E111C,E111D,E111E,E112A,E112B,E112C,E112D,E112E,E113A,E113B,E113C,E113D,E113E,E114A,E114B,E114C,E114D,E114E,E115A,E115B,E115C,E115D,E115E,E116A,E116B,E116C,E116D,E116E,E117A,E117B,E117C,E117D,E117E,E118A,E118B,E118C,E118D,E118E,E1TipoE,E1Valor1,E1VUR3) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
    
    # Ejecutar la consulta
    cursor.execute(query1, (cod, padron, next_code, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, e1cat, 0, valor_unitario[e1cat]))
    #---------------------------------------------------------------------------

    query = ""



    # Confirmar la transacción. Conviene realizar al final del recorrido
    #connection.commit()


    # Cerrar el cursor y la conexión
    cursor.close()
    connection.close()
#    print("Se realizó el INSERT en la BD exitosamente")

except pyodbc.Error as e:
    print("Error al realizar el insert:", e)
