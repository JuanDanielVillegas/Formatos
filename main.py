import pyodbc
import os
import subprocess
import pathlib


from cells import *
from functions import *
from forms import *
from conn import *

def find_emp(nom, apat, amat): 

    try:
        connection = pyodbc.connect(conn_str)
        cursor = connection.cursor()

        cursor.execute(query, ('%' + nom + '%', '%' + apat + '%', '%' +amat + '%'))

        rows = cursor.fetchall()

        ip_value = str(rows[0].ip) if rows[0].ip is not None else 'ip_not_found'

        emp = {
            'A8' : get_date(),
            'A14' : rows[0].Departamento.title(),
            'A18' : rows[0].Nombre.title(),
            'A19' : rows[0].Puesto.title(),
            'A20' : rows[0].RFC,
            'C20' : rows[0].Curp,
            'A21' : rows[0].Numero,
            'A22' : rows[0].correo.replace('BUZON.', '').lower(),
            'A24': '(614) 429-33-00 Ext. ' + rows[0].ext,
            'A25': rows[0].Superior.title(),
            'A34': ip_value,
            'A81' : rows[0].titulo.upper() + ' ' + rows[0].Nombre.title()  
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

    # print('(1). Envío de Celula')
    # print('Control y Seguimiento de Propuestas')
    # print('(2). S=Usuario')
    # print('(3). T=Transferencia')
    # print('(4). Seguimiento Revisiones S=Usuario')
    # print('(5). Omitir')


    if cedula == 1:
        cedula = 'B54'
    if st == 1:
        st ='B57'
    if st == 2:
        st = 'D57'
    if s == 1:
        s = 'B59'


    return suiefi, cedula, st, s


def select_emp(nom, apat, amat):
    emp = find_emp(nom, apat, amat)
    global cells
    cells |= emp
    os.system('cls')
    print('Numero de empleado \t Nombre')
    print(emp['A21'],'\t\t', emp['A18'])
    name = emp['A21'] + '                                                 '+ emp['A18']
    
    return emp


def start(emp, opt1, opt2, cedula, st, s, checkboxes):

    path = './plantillas_formatos/Formatos/' + emp['A18']
    
    if not(os.path.exists(path) and os.path.isdir(path)):
        os.mkdir(path)

    checked = mark_cells(opt1)
    suiefi = mark_suiefi(opt2, cedula, st, s)

    fill_templates(emp['A18'], checked, suiefi, cedula, st, s, checkboxes)
    command = 'explorer ' + '"' + str(pathlib.Path(__file__).parent.absolute()) + '\plantillas_formatos\Formatos\\' + emp['A18'] + '"'
    
    subprocess.Popen(command)

    
    






    