import os

from settings import get_paths


def find_file(gk_name: str, data_folder_path: str) -> str | None:
    """
    Функция-поиск файла с КГ
    """

    path = data_folder_path

    for rootdir, dirs, files in os.walk(path):
        for file in files:
            if file.split('.')[0].lower() == gk_name.lower():
                return f'{rootdir}\\{file}'

    return None
