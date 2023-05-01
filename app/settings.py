import os


# ПЕРЕМЕННАЯ УКАЗЫВАЩАЯ НА ПУТЬ ПРОГРАММЫ
APP_PATH = os.path.dirname(os.path.abspath(__file__))
# APP_PATH = ''APP_PATH.split('\\')[:-1]

FILES_PATH = {'DATA_FOLDER_PATH': None, 'SENSOR_DATA_PATH': None}

try:
    with open(f'{APP_PATH}\\settings.txt') as file:
        for line in file.readlines():
            if ':=' not in line:
                continue
            key, value = line.split(':=')
            if key.strip() in FILES_PATH.keys():
                FILES_PATH[key.strip()] = value.strip()
except FileNotFoundError:
    print('Ошибка: Файл settings.txt не найден. Проверьте что он существует.')
    input('Нажмите enter для выхода из программы...')


DATA_FOLDER_PATH = FILES_PATH['DATA_FOLDER_PATH']
SENSOR_DATA_PATH = FILES_PATH['SENSOR_DATA_PATH']
