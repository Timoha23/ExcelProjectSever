import os
from datetime import datetime

import openpyxl

from settings import APP_PATH


class Backup:
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.name_folder_date = str(datetime.utcnow().date())
        self.backups_path = f'{APP_PATH}\\backups'
        self.path_backup_file = None
        self.first_date_path = False

    def create(self, gk_name: str) -> None:
        if not os.path.exists(self.backups_path):
            os.makedirs(self.backups_path)

        self.path_folder_date = f'{self.backups_path}\\{self.name_folder_date}'
        if not os.path.exists(self.path_folder_date):
            self.first_date_path = True
            os.makedirs(self.path_folder_date)
        full_file_name = gk_name + '_' + str(datetime.utcnow().time().strftime('%H_%M_%S')) + '.xlsx'
        wb = openpyxl.load_workbook(self.file_path)
        self.path_backup_file = f'{self.path_folder_date}\\{full_file_name}'
        wb.save(self.path_backup_file)
        return self.path_backup_file

    def delete(self):
        if self.path_backup_file is not None:
            if self.first_date_path is True:
                os.remove(self.path_folder_date)
            else:
                os.remove(self.path_backup_file)



# if not os.path.exists(f'backups/{name_folder_date}'):
#     os.makedirs(name_folder_date)
