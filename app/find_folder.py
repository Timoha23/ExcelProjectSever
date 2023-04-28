import os

import settings


def find_file(gk_name: str) -> str | None:
    """
    Функция-поиск файла с КГ
    """

    path = settings.DATA_FOLDER_PATH

    for rootdir, dirs, files in os.walk(path):
        for file in files:
            if file.split('.')[0].lower() == gk_name.lower():
                return f'{rootdir}\\{file}'

    return None
