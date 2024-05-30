import pyodbc
import os
import subprocess
import pathlib

from cells import *
from functions import *
from forms import *
from conn import *

def find_emp(nombre): 

    try:
        connection = pyodbc.connect(conn_str)
        cursor = connection.cursor()

        param = '%' + nombre + '%'

        print(param)

        cursor.execute(query, (param))

        rows = cursor.fetchall()

        print(rows[0].Departamento.title())

        alert = '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!'

        num_value = rows[0].Numero if rows[0].Numero is not None else alert
        curp = rows[0].Curp if rows[0].Curp is not None else alert
        rfc = rows[0].RFC if rows[0].RFC is not None else alert
        ip_value = str(rows[0].ip) if rows[0].ip is not None else 'ip_not_found'
        depto_value = fix_y(rows[0].Departamento.title()) if rows[0].Departamento is not None else 'depto_not_found'
        title_value = rows[0].titulo.upper() if rows[0].titulo is not None else 'C.'
        email_value = rows[0].correo.replace('BUZON.', '').lower() if rows[0].correo is not None else alert
        ext = str(rows[0].ext) if rows[0].ext is not None else alert
        puesto = rows[0].Puesto.title() if rows[0].Puesto is not None else alert

        # PENDIENTE
        # Alertar sobre tipos de datos nonetype 
        # Transformar a string


        emp = {
            'A8' : get_date(),
            'A14' : depto_value,
            'A18' : rows[0].Nombre.title(),
            'A19' : puesto,
            'A20' : rfc,
            'C20' : curp,
            'A21' : num_value,
            'A22' : email_value,
            'A24': '(614) 429-33-00 Ext. ' + ext,
            'A25': rows[0].Superior.title(),
            'A34': ip_value,
            'A81' : title_value + ' ' + rows[0].Nombre.title()  
        }

        return emp

    except pyodbc.Error as e:
        print(f'Error de conexión: {e}')

    finally:
        if connection:
            connection.close()



def mark_cells(opt):

    checked = 'C41'

    if opt == 'Alta':
        checked = 'C41'
    if opt == 'Baja':
        checked ='E41'
    if opt == 'Cambio':
        checked = 'E43'
    if opt == 'Reactivación':
        checked = 'G41'

    return checked

def mark_suiefi(opt, cedula, st, s):
    suiefi = False
    x = False

    if opt == 'Alta':
        suiefi = 'C47'
    if opt == 'Baja':
        suiefi ='E47'
    if opt == 'Cambio':
        suiefi = 'C49'

    if cedula == 1:
        cedula = 'B54'
    if st == 1:
        st ='B57'
    if st == 2:
        st = 'D57'
    if s == 1:
        s = 'B59'


    return suiefi, cedula, st, s


def select_emp(nombre):
    emp = find_emp(nombre)
    global cells
    cells |= emp
    os.system('cls')
    print('Numero de empleado \t Nombre')
    print(emp['A21'],'\t\t', emp['A18'])
    name = emp['A21'] + '                                                 '+ emp['A18']
    
    return emp


def start(emp, opt1, opt2, cedula, st, s, checkboxes):

    path = './plantillas_formatos/Formatos/' + emp['A18'].strip()
    
    if not(os.path.exists(path) and os.path.isdir(path)):
        os.mkdir(path)

    checked = mark_cells(opt1)
    suiefi = mark_suiefi(opt2, cedula, st, s)

    fill_templates(emp['A18'], checked, suiefi, cedula, st, s, checkboxes)
    command = 'explorer ' + '"' + str(pathlib.Path(__file__).parent.absolute()) + '\plantillas_formatos\Formatos\\' + emp['A18'].strip() + '"'
    
    subprocess.Popen(command)

    
    






    