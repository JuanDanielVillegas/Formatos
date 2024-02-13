import pyodbc

server_name = 'INTRAFIGO_SVR'
database_name = 'fiscalizacion'
username = 'sa'
password = 'Systemas07'
conn_str = f'DRIVER={{SQL Server}};SERVER={server_name};DATABASE={database_name};UID={username};PWD={password}'

nom = input('Nombre: ')
apat = input('Apellido paterno: ')
amat = input('Apellido materno: ')

try:
    connection = pyodbc.connect(conn_str)
    cursor = connection.cursor()

    sql_query = """
        SELECT
            emp.str_empleado AS 'Numero',
            titulo.str_titulo AS titulo,
            CONCAT(emp.str_nombre, ' ', emp.str_apat, ' ', emp.str_amat) AS 'Nombre', 
            emp.str_curp AS Curp,
            emp.str_rfc AS RFC,
            puesto.str_puesto AS 'Puesto',
            depto.str_depto AS Departamento,
			emp.str_email AS correo,
            CONCAT(sup.str_nombre, ' ', sup.str_apat, ' ', sup.str_amat) AS 'Nombre del superior'
        FROM 
            dd_empleados AS emp
            INNER JOIN dd_empleados AS sup ON emp.c_superior = sup.c_empleado
            INNER JOIN mm_departamento AS depto ON depto.c_depto = emp.c_depto
            INNER JOIN mm_ciudad AS ciudad ON ciudad.c_ciudad = emp.c_ciudad
            INNER JOIN mm_puesto AS puesto ON puesto.c_puesto = emp.c_puesto
            INNER JOIN mm_titulo AS titulo ON titulo.c_titulo = emp.c_titulo
        WHERE 
            emp.c_ciudad = 1 AND emp.bol_vige = 1 
            AND emp.str_nombre  COLLATE Latin1_General_100_CI_AI LIKE ? COLLATE Latin1_General_100_CI_AI
            AND emp.str_apat  COLLATE Latin1_General_100_CI_AI LIKE ? COLLATE Latin1_General_100_CI_AI
            AND emp.str_amat  COLLATE Latin1_General_100_CI_AI LIKE ? COLLATE Latin1_General_100_CI_AI;
    """
    cursor.execute(sql_query, ('%' + nom + '%', '%' + apat + '%', '%' +amat + '%'))

    rows = cursor.fetchall()

    

    
    depto = rows[0].Departamento
    nombre = rows[0].Nombre
    puesto = rows[0].Puesto
    rfc = rows[0].RFC
    curp = rows[0].Curp
    numero_emp = rows[0].Numero
    correo = rows[0].correo.replace('BUZON.', '')
    titulo = rows[0].titulo


    print(depto+'| '+nombre+'| '+puesto+'| '+rfc+'| '+curp+'| '+ numero_emp+'| '+correo+'| '+titulo.upper())


    #for row in rows:
    #    print(row)

except pyodbc.Error as e:
    print(f'Error de conexión: {e}')

finally:
    if connection:
        connection.close()