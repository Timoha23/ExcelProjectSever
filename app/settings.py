import os


FILES_PATH = {'DATA_FOLDER_PATH': None, 'SENSOR_DATA_PATH': None}
with open('settings.txt') as file:
    for line in file.readlines():
        if ':=' not in line:
            continue
        key, value = line.split(':=')
        if key.strip() in FILES_PATH.keys():
            FILES_PATH[key.strip()] = value.strip()


DATA_FOLDER_PATH = FILES_PATH['DATA_FOLDER_PATH']
SENSOR_DATA_PATH = FILES_PATH['SENSOR_DATA_PATH']

# ПЕРЕМЕННАЯ УКАЗЫВАЩАЯ НА ПУТЬ ПРОГРАММЫ
APP_PATH = os.path.dirname(os.path.abspath(__file__))
