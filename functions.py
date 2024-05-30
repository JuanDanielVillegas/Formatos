from datetime import datetime

def get_month(n):
    months = [
        'Enero',
        'Febrero',
        'Marzo',
        'Abril',
        'Mayo',
        'Junio',
        'Julio',
        'Agosto',
        'Septiembre',
        'Octubre',
        'Noviembre',
        'Diciembre'
    ]

    return months[n-1]

def fix_y(depto):
  
    if(depto in ['Programación Y Seguimiento', 'Gabinete Y Dictámenes']):
        array = depto.split()
        return ' '.join([array[0], array[1].lower(), array[2]])
    else:
        return depto

def get_date():
    now = datetime.now()
    return str(now.day) + ' de ' + get_month(now.month) + ' ' + str(now.year)

def get_folder_date():
    date = datetime.now()
    return str(date.day) + '-' + str(date.month) + '-' + str(date.year)


def fill_template_tags():

    cells = ['A','B','C','D','E','F','G']

    #t = [f'A{i+1}' for i in range(81)]
    t = []

    for c in cells:
        for i in range(81):
            t.append(f'{c}{i+1}')

    return t


