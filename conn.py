
server_name = 'INTRAFIGO_SVR'
database_name = 'fiscalizacion'
username = 'sa'
password = 'Systemas07'
conn_str = f'DRIVER={{SQL Server}};SERVER={server_name};DATABASE={database_name};UID={username};PWD={password}'

query = """
    SELECT
        emp.str_empleado AS 'Numero',
        titulo.str_titulo AS titulo,
        CONCAT(emp.str_nombre, ' ', emp.str_apat, ' ', emp.str_amat) AS 'Nombre', 
        emp.str_curp AS Curp,
        emp.str_rfc AS RFC,
        puesto.str_puesto AS 'Puesto',
        depto.str_depto AS Departamento,
        emp.str_email AS correo,
        emp.str_telefono AS ext,
        str_ip AS ip,
        CONCAT(sup.str_nombre, ' ', sup.str_apat, ' ', sup.str_amat) AS 'Superior'
    FROM 
        dd_empleados AS emp
        INNER JOIN dd_empleados AS sup ON emp.c_superior = sup.c_empleado
        INNER JOIN mm_departamento AS depto ON depto.c_depto = emp.c_depto
        INNER JOIN mm_ciudad AS ciudad ON ciudad.c_ciudad = emp.c_ciudad
        INNER JOIN mm_puesto AS puesto ON puesto.c_puesto = emp.c_puesto
        INNER JOIN mm_titulo AS titulo ON titulo.c_titulo = emp.c_titulo
        INNER JOIN dd_cuentas AS cuentas ON cuentas.c_empleado = emp.c_empleado
    WHERE 
        emp.c_ciudad = 1 AND emp.bol_vige = 1 
        AND emp.str_nombre  COLLATE Latin1_General_100_CI_AI LIKE ? COLLATE Latin1_General_100_CI_AI
        AND emp.str_apat  COLLATE Latin1_General_100_CI_AI LIKE ? COLLATE Latin1_General_100_CI_AI
        AND emp.str_amat  COLLATE Latin1_General_100_CI_AI LIKE ? COLLATE Latin1_General_100_CI_AI
    ORDER BY 
	CASE WHEN str_ip IS NULL OR str_ip = ' ' THEN 1 ELSE 0 END
"""