import os
from pathlib import Path

# ПЕРЕМЕННАЯ УКАЗЫВАЩАЯ НА ПУТЬ ПРОГРАММЫ
APP_PATH = os.path.dirname(os.path.abspath(__file__))


def get_paths() -> dict[str]:
    """
    Получаем пути к файлу с датчиками и с ГК.
    """
    paths = {'DATA_FOLDER_PATH': None, 'SENSOR_DATA_PATH': None}
    settings_txt_path = str(Path(APP_PATH).parent) + '\\settings.txt'
    with open(settings_txt_path, encoding='utf-8') as file:
        for line in file.readlines():
            if ':=' not in line:
                continue
            key, value = line.split(':=')
            if key.strip() in paths.keys():
                paths[key.strip()] = value.strip()

    return paths
