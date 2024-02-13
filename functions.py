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


def get_date():
    now = datetime.now()
    return str(now.day) + ' de ' + get_month(now.month) + ' ' + str(now.year)


def fill_template_tags():

    cells = ['A','B','C','D','E','F','G']

    #t = [f'A{i+1}' for i in range(81)]
    t = []

    for c in cells:
        for i in range(81):
            t.append(f'{c}{i+1}')

    return t


