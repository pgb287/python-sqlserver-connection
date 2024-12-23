import pyodbc
import csv
from datetime import datetime

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

# Padrones no incorporados porque tienen algun rubro del E1 o E2 mal cargado, es decir no coinciden los codigos maximos
padrones_no_incorporados = []
# Contador de mejoras cargadas
contador_e1 = 0
contador_e2 = 0

datenow = datetime.now()
fecha_carga = datetime.strptime(datenow, "%Y-%m-%d %H:%M:%S.%f")                


try:
    connection = pyodbc.connect(connection_string)
    cursor = connection.cursor()

    with open('datos.csv', mode='r', encoding='utf-8') as file:
        csv_reader = csv.DictReader(file, delimiter=',')  # Usamos DictReader para trabajar con los nombres de las columnas       

        # Recorre el CSV con las mejoras a incorporar        
        for row in csv_reader:                             
                        
            # Asigna los datos del CSV a variables
            cod = row['cod']   
            padron = row['padron']        
            cu = float(row['cu'])
            sc = float(row['sc'])
            pi = float(row['pi'])
            ga = float(row['ga'])
            pa = float(row['pa'])
            e1cat = row['e1cat']
            e1est = row['e1est']
            e2cat = row['e2cat']
            e2est = row['e2est']
            ph = row['ph']

            bande1 = 0
            bande2 = 0
            
            if (cu + sc + pi) > 0 :

                #**********************************************CARGA DE E1**********************************************************
                #-------------------------------------------------------------------------------------------------------------------
                # Selecciona el codigo mayor de las 6 tablas, es decir devuelve una tabla con 6 registros (tabla/valor)
                queryE1 = f"""SELECT 'TAB006' AS Tabla, MAX(T01CodE1) AS Max FROM TAB006 WHERE DCodDep = '{cod}' AND T0Padron = '{padron}' UNION ALL
                SELECT 'TE1R1' AS Tabla, MAX(E1Cod) AS Max FROM TE1R1 WHERE DCodDep = '{cod}' and T0Padron = '{padron}' UNION ALL
                SELECT 'TE1R2' AS Tabla, MAX(E1R2Cod) AS Max FROM TE1R2 WHERE DCodDep = '{cod}' and T0Padron = '{padron}' UNION ALL
                SELECT 'TE1R4' AS Tabla, MAX(E1R4Cod) AS Max FROM TE1R4 WHERE DCodDep = '{cod}' and T0Padron = '{padron}' UNION ALL
                SELECT 'TE1R5' AS Tabla, MAX(E1R5Cod) AS Max FROM TE1R5 WHERE DCodDep = '{cod}' and T0Padron = '{padron}' UNION ALL
                SELECT 'TE1R6' AS Tabla, MAX(E1R6Cod) AS Max FROM TE1R6 WHERE DCodDep = '{cod}' and T0Padron = '{padron}'"""

                cursor.execute(queryE1)
                rE1 = cursor.fetchall()
                
                # Recorre y compara los codigos de las 6 tablas. Si son iguales continua el proceso (quiere decir que no hay errores en los e1 cargados), caso contrario se separa el padron para su posterior analisis y continua con otro registro 
                if all(row[1] == rE1[0][1] for row in rE1):
                    # Todos los valores son iguales o bien no hay E1 cargado (baldio)
                    
                    # Toma el codigo y le suma uno para asignar el nuevo codigo de E1
                    next_code = 1 if rE1[0][1] == None else rE1[0][1]+1
                    # Valor basico unitario para cada categoria                
                    valor_unitario = {"A":833.27, "B":640.83, "C":506.98, "D":316.46, "E":175.69}
                    # Coeficiente de antiguedad (para antiguedad de 1 año) segun la categoria
                    coef_antiguedad = {"A":1, "B":1, "C":1, "D":0.99, "E":0.98}
                    # Antiguedad unica al ser incorporacion masiva
                    antiguedad = 1
                    # Valor basico unitario para piletas segun la categoria, en el caso de vivienda categoria C, D y E se asigna pileta tipo C
                    tipo_pileta = {'A':['A',290.49], 'B':['B',174.17], 'C':['C',95.51], 'D':['C',95.51], 'E':['C',95.51]}
                    
                                    
                    #------------------------------------------------------------------------                        
                    # Insertar datos en tabla TE1R1
                    queryTE1R1 = """INSERT INTO TE1R1 (DCodDep,T0Padron,E1Cod,E11A,E11B,E11C,E11D,E11E,E12A,E12B,E12C,E12D,E12E,E13A,E13B,E13C,E13D,E13E,E14A,E14B,E14C,E14D,E14E,E15A,E15B,E15C,E15D,E15E,E16A,E16B,E16C,E16D,E16E,E17A,E17B,E17C,E17D,E17E,E18A,E18B,E18C,E18D,E18E,E19A,E19B,E19C,E19D,E19E,E110A,E110B,E110C,E110D,E110E,E111A,E111B,E111C,E111D,E111E,E112A,E112B,E112C,E112D,E112E,E113A,E113B,E113C,E113D,E113E,E114A,E114B,E114C,E114D,E114E,E115A,E115B,E115C,E115D,E115E,E116A,E116B,E116C,E116D,E116E,E117A,E117B,E117C,E117D,E117E,E118A,E118B,E118C,E118D,E118E,E1TipoE,E1Valor1,E1VUR3) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"""
                    
                    # Ejecutar la consulta
                    cursor.execute(queryTE1R1, (cod, padron, next_code, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, e1cat, 0, valor_unitario[e1cat]))
                    #---------------------------------------------------------------------------


                    #------------------------------------------------------------------------                        
                    # Insertar datos en tabla TE1R2
                    pi_tipo = '' if pi == 0 else tipo_pileta[e1cat][0]

                    queryTE1R2 = "INSERT INTO TE1R2 (DCodDep,T0Padron,E1R2Cod,E1R2a,E1R2b,E1R2c,E1R2d,E1R2e,E1R2f,E1R2g,E1R2h,E1R2i,E1R2ii,E1R2j,E1R2k,E1R2l,E1R2m,E1Const2,E1Const3,E1Const4,E1R2Est,E1Construc) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
                    
                    # Ejecutar la consulta
                    cursor.execute(queryTE1R2, (cod, padron, next_code, e1est, antiguedad, cu, sc, 0, 0, 0, 0, pi, pi_tipo, 0, 0, 0, 0, 'N', '', '', '', 'N'))
                    #---------------------------------------------------------------------------


                    #------------------------------------------------------------------------                        
                    # Insertar datos en tabla TE1R4
                    valor_cu = coef_antiguedad[e1cat] * valor_unitario[e1cat] * cu
                    valor_sc = coef_antiguedad[e1cat] * round((valor_unitario[e1cat]/2),2) * sc
                    valor_cu_sc = valor_cu + valor_sc

                    queryTE1R4 = "INSERT INTO TE1R4 (DCodDep,T0Padron,E1R4cod,E1R4coef,E1R4VUa,E1R4VUb,E1R4valA,E1R4valB,E1R4valC,E1R4valD,E1R4valE,E1R4Tot) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)"
                    
                    # Ejecutar la consulta
                    cursor.execute(queryTE1R4, (cod, padron, next_code, coef_antiguedad[e1cat], valor_unitario[e1cat], round((valor_unitario[e1cat]/2),2), round(valor_cu,2), round(valor_sc,2), 0, 0, 0, round(valor_cu_sc,2)))                
                    #---------------------------------------------------------------------------

                    
                    #------------------------------------------------------------------------                        
                    # Insertar datos en tabla TE1R5
                    queryTE1R5 = "INSERT INTO TE1R5 (DCodDep,T0Padron,E1R5Cod,E1R5VUa,E1R5VUb,E1R5Vala,E1R5Valb,E1R5Valc,E1R5vald,E1R5Tot) VALUES (?,?,?,?,?,?,?,?,?,?)"
                    
                    # Ejecutar la consulta
                    cursor.execute(queryTE1R5, (cod, padron, next_code, 0, 0, 0, 0, 0, 0, 0))
                    #---------------------------------------------------------------------------


                    #---------------------------------------------------------------------------                        
                    # Insertar datos en tabla TE1R6
                    valor_pileta = pi * coef_antiguedad[e1cat] * tipo_pileta[e1cat][1]

                    queryTE1R6 = "INSERT INTO TE1R6 (DCodDep,T0Padron,E1R6Cod,E1R6Vala,E1R6Valb,E1R6Valc,E1R6Vald,E1R6Ve1,E1R6Ve2,E1R6Valf,E1R6Tot) VALUES (?,?,?,?,?,?,?,?,?,?,?)"
                    
                    # Ejecutar la consulta
                    cursor.execute(queryTE1R6, (cod, padron, next_code, 0, 0, valor_pileta, 0, 0, 0, 0, valor_pileta))
                    #---------------------------------------------------------------------------


                    #---------------------------------------------------------------------------                        
                    # Insertar datos en tabla TAB006
                    leyenda_e1 = 'Incorporación automática'
                    #fecha_carga = '2024-11-21 00:00:00.000'
                    #fecha_carga_dt = datetime.strptime(fecha_carga, "%Y-%m-%d %H:%M:%S.%f")                
                    valor_total_e1 = valor_cu_sc + valor_pileta
                    
                    queryTAB006 = "INSERT INTO TAB006 (DCodDep,T0Padron,T01CodE1,T01Desc,T01Fecha,T01Estado,T01ValTot,T01fcarga,T01fmodi,T01fbaja,T01user1,T01user2,T01user3) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)"
                    
                    # Ejecutar la consulta
                    cursor.execute(queryTAB006, (cod, padron, next_code, leyenda_e1, fecha_carga, 'A', valor_total_e1, fecha_carga_dt, '', '', 'dbo','',''))
                    #---------------------------------------------------------------------------
                    
                    connection.commit()
                    # Contador de mejoras cargadas, es decir padrones actualizados
                    contador_e1 = contador_e1 + 1
                    # Bandera que indica si se cargó un E1
                    bande1 = 1
                
                else: #Los codigos no son todos iguales, no se incorporan las mejoras del padron
                    padrones_no_incorporados.append(padron)                    


            if (ga + pa) > 0:
                
                #**********************************************CARGA DE E2**********************************************************
                #-------------------------------------------------------------------------------------------------------------------
                # Selecciona el codigo mayor de las 5 tablas, es decir devuelve una tabla con 5 registros (tabla/valor)
                queryE2 = f"""SELECT 'TAB007' AS Tabla, MAX(T02CodE2) AS Max FROM TAB007 WHERE DCodDep = '{cod}' AND T0Padron = '{padron}' UNION ALL
                SELECT 'TE2R1' AS Tabla, MAX(E2Cod) AS Max FROM TE2R1 WHERE DCodDep = '{cod}' and T0Padron = '{padron}' UNION ALL
                SELECT 'TE2R2A' AS Tabla, MAX(E2R2Cod) AS Max FROM TE2R2A WHERE DCodDep = '{cod}' and T0Padron = '{padron}' UNION ALL
                SELECT 'TE2R4' AS Tabla, MAX(E2R4cod) AS Max FROM TE2R4 WHERE DCodDep = '{cod}' and T0Padron = '{padron}' UNION ALL
                SELECT 'TE2R5' AS Tabla, MAX(E2R5Cod) AS Max FROM TE2R5 WHERE DCodDep = '{cod}' and T0Padron = '{padron}'"""

                cursor.execute(queryE2)
                rE2 = cursor.fetchall()
                                                
                # Recorre y compara los codigos de las 5 tablas. Si son iguales continua el proceso (quiere decir que no hay errores en los e2 cargados), caso contrario se separa el padron para su posterior analisis y continua con otro registro 
                if all(row[1] == rE2[0][1] for row in rE2):
                    # Todos los valores son iguales o bien no hay E2 cargado
                    
                    # Toma el codigo y le suma uno para asignar el nuevo codigo de E2
                    next_code_e2 = 1 if rE2[0][1] == None else rE2[0][1]+1
                    # Valor basico unitario para cada categoria                
                    valor_unitario_e2 = {"A":335.67, "B":278.36, "C":211.21, "D":69.68}
                    # Coeficiente de antiguedad (para antiguedad de 1 año) segun la categoria
                    coef_antiguedad_e2 = {"A":1, "B":1, "C":1, "D":0.99}
                    # Antiguedad unica al ser incorporacion masiva
                    antiguedad = 1
                    # Valor basico unitario para pavimento segun la categoria, en el caso de vivienda categoria C, D y E se asigna pileta tipo C
                    tipo_pavimento = {'A':['A',37.41], 'B':['B',28.12], 'C':['B',28.12], 'D':['B',28.12]}

                    #------------------------------------------------------------------------                        
                    # Insertar datos en tabla TE2R1
                    queryTE2R1 = """INSERT INTO TE2R1 (DCodDep,T0Padron,E2Cod,E21A,E21B,E21C,E21D,E22A,E22B,E22C,E22D,E23A,E23B,E23C,E23D,E24A,E24B,E24C,E24D,E25A,E25B,E25C,E25D,E26A,E26B,E26C,E26D,E27A,E27B,E27C,E27D,E28A,E28B,E28C,E28D,E29A,E29B,E29C,E29D,E210A,E210B,E210C,E210D,E211A,E211B,E211C,E211D,E212A,E212B,E212C,E212D,E213A,E213B,E213C,E213D,E214A,E214B,E214C,E214D,E215A,E215B,E215C,E215D,E216A,E216B,E216C,E216D,E217A,E217B,E217C,E217D,E2TipoE,E2Valor1,E2VUR3) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"""
                    
                    # Ejecutar la consulta
                    cursor.execute(queryTE2R1, (cod, padron, next_code_e2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', '', ''))
                    #---------------------------------------------------------------------------


                    #------------------------------------------------------------------------                        
                    # Insertar datos en tabla TE2R2A                 
                    queryTE2R2A = "INSERT INTO TE2R2A (DCodDep,T0Padron,E2R2Cod,E2R2a,E2R2b,E2R2c,E2R2d,E2R2e,E2R2fa,E2R2fb,E2R2g,E2R2h,E2R2i,E2R2j,E2Construc,E2Const2,E2Est) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
                    
                    # Ejecutar la consulta
                    cursor.execute(queryTE2R2A, (cod, padron, next_code_e2, '', '', '', '', '', '', '', '', '', '', '', '', '', ''))
                    #---------------------------------------------------------------------------


                    #------------------------------------------------------------------------                        
                    # Insertar datos en tabla TE2R2A                 
                    queryTE2R4 = "INSERT INTO TE2R4 (DCodDep,T0Padron,E2R4cod,E2R4coef,E2R4VUa,E2R4VUb,E2R4VUc,E2R4VUd,E2R4valA,E2R4valB,E2R4valC,E2R4valD,E2R4Tot) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)"
                    
                    # Ejecutar la consulta
                    cursor.execute(queryTE2R4, (cod, padron, next_code_e2, '', '', '', '', '', '', '', '', '', ''))
                    #---------------------------------------------------------------------------


                    #------------------------------------------------------------------------                        
                    # Insertar datos en tabla TE2R5                 
                    queryTE2R5 = "INSERT INTO TE2R5 (DCodDep,T0Padron,E2R5Cod,E2R5VUa,E2R5VUb,E2R5VUc,E2R5VUd1,E2R5VUd2,E2R5Vala,E2R5Valb,E2R5Valc,E2R5vad1,E2R5vad2,E2R5Tot) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
                    
                    # Ejecutar la consulta
                    cursor.execute(queryTE2R5, (cod, padron, next_code_e2, '', '', '', '', '', '', '', '', '', '',''))
                    #---------------------------------------------------------------------------


                    #---------------------------------------------------------------------------                        
                    # Insertar datos en tabla TAB007
                    # leyenda_e1 = 'Incorporación automática'
                    # fecha_carga = '2024-11-21 00:00:00.000'
                    # fecha_carga_dt = datetime.strptime(fecha_carga, "%Y-%m-%d %H:%M:%S.%f")                
                    # valor_total_e1 = valor_cu_sc + valor_pileta
                    
                    queryTAB007 = "INSERT INTO TAB007 (DCodDep,T0Padron,T02CodE2,T02Estado,T02Valtot,T02Desc,T02Fecha,T02user1,T02user2,T02user3,T02fcarga,T02fmodi,T02baja) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)"
                    
                    # Ejecutar la consulta
                    cursor.execute(queryTAB007, (cod, padron, next_code_e2, '', '', '', '', '', '', '', '', '', ''))
                    #---------------------------------------------------------------------------

                    connection.commit()
                    # Contador de mejoras cargadas, es decir padrones actualizados
                    contador_e2 = contador_e2 + 1
                    # Bandera que indica si se cargó un E2
                    bande2 = 1
                
                else: #Los codigos no son todos iguales, no se incorporan las mejoras del padron
                    padrones_no_incorporados.append(padron)                    

            
            #------------------------------------------------------------------------                        
            # Actualizar datos en tabla TAB00 (U PRINCIPAL)

            # Primero traigo algunos datos del padron 
            queryTraer = "SELECT T0Observ,T0VTFiscal FROM TAB00 WHERE DCodDep = ? AND T0Padron = ?"
            cursor.execute(queryTraer, (cod, padron))
            rTraer = cursor.fetchone()
            # rTraer[0] rTraer[1]
            leyenda_u = rTraer[0] + ' - Se actualiza VF por incorporación de mejoras 2024 (automatica)'            

            
            queryTAB00 = "UPDATE TAB00 SET T0Tipo = ?,T0Observ = ?,T0FVigen = ?,T0User2 = ?,T0FModif = ?,T0VFiscal = ?,T0VF08 = ?,T0VTFiscal = ?,T0VMjra = ? WHERE DCodDep = ? AND T0Padron = ?"
            
            # Ejecutar la consulta
            cursor.execute(queryTAB00, ('E', leyenda_u, '20250101 00:00:00', 'dbo', fecha_carga, continuar aquiiiiii  ))
            #---------------------------------------------------------------------------




        # Confirmar la transacción. Conviene realizar al final del recorrido
        connection.commit()

    # Cerrar el cursor y la conexión
    cursor.close()
    connection.close()

    print(f'Total de E1 cargados: {contador_e1}')
    print(f'Total de E2 cargados: {contador_e2}')
    print(f'Total de padrones no cargados: {padrones_no_incorporados}')

except pyodbc.Error as e:
    print("Error al realizar el insert:", e)



#TODA LA CARGA DEL E1 ESTA FINALIZADAAAA
#SEGUIR CON LA CARGA DEL E2 (CREAR LAS TABLAS) Y LUEGO ACTUALIZAR FORMULARIO U