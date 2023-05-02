import logging
import os
import time
import traceback
from pathlib import Path

from add_row import (add_new_row, find_cell_ts, find_header_in_sheet,
                     get_header_lengths_and_input_row, make_merged_cells,
                     make_style_for_new_row, put_data_to_excel,
                     unmerge_all_cells_after_header)
from backup import Backup
from find_folder import find_file
from sensor_data import get_sensor_data
from settings import APP_PATH, get_paths

# создаем директорию с логом

log_path = str(Path(APP_PATH).parent) + '\\logs'
backups_path = str(Path(APP_PATH).parent) + '\\backups'

if not os.path.exists(log_path):
    os.makedirs(log_path)


logging.basicConfig(
    level=logging.INFO,
    filename=f'{log_path}\\log.log',
    format='%(asctime)s:%(levelname)s - %(message)s',
    datefmt='%d/%m/%Y %H:%M:%S',
    encoding='utf-8'
)


def main():
    steps = 12
    count_steps = 1
    sensor_number = input('Введите номер сенсора. Пример("n232"): ')
    logging.info(f'Начали работу. Датчик {sensor_number}')
    print(f'[{count_steps}/{steps}] Начинаем работу')
    count_steps += 1

    # проверяем пути
    try:
        paths = get_paths()
        for key, value in paths.items():
            if value is None:
                return (f'Ошибка. Проверьте путь для {key} в файле settings.')
        data_folder_path, sensor_data_path = paths['DATA_FOLDER_PATH'], paths['SENSOR_DATA_PATH']
        print(f'[{count_steps}/{steps}] Получили пути к файлам')
        count_steps += 1
    except FileNotFoundError:
        print('Ошибка: Файл settings.txt не найден. Проверьте что он существует.')
        return
    except Exception as ex:
        logging.critical(f'Неизвестная ошибка: {ex}. Функция: get_paths(). Traceback: {traceback.format_exc()}')
        print(f'Неизвестная ошибка: {ex}. Функция: get_paths().')
        return

    # сбор иноформации с датчиков
    try:
        sensor_data = get_sensor_data(sensor_number, sensor_data_path)
        if sensor_data.get('error'):
            logging.error(f'Ошибка: {sensor_data["error"]}')
            print('Ошибка:', sensor_data['error'])
            return
        sensor_gk_name = sensor_data['gk_name']
        sensor_ts_number = sensor_data['ts_number']
        print(f'[{count_steps}/{steps}] Получили данные с датчиков')
        count_steps += 1
    except Exception as ex:
        logging.critical(f'Неизвестная ошибка: {ex}. Функция: get_sensor_data(). Traceback: {traceback.format_exc()}')
        print(f'Неизвестная ошибка: {ex}. Функция: get_sensor_data().')
        return

    # поиск файла с нужным ГК
    try:
        file_path = find_file(sensor_gk_name, data_folder_path)
        if file_path is None:
            logging.error(f'Ошибка: Файл с именем {sensor_gk_name} не найден.')
            print(f'Ошибка: Файл с именем {sensor_gk_name} не найден.')
            return
        print(f'[{count_steps}/{steps}] Нашли файл: {file_path}')
        count_steps += 1
    except Exception as ex:
        logging.critical(f'Неизвестная ошибка: {ex}. Функция: find_file(). Traceback: {traceback.format_exc()}')
        print(f'Неизвестная ошибка: {ex}. Функция: find_file()')
        return

    # создаем бекап файла
    try:
        backup = Backup(file_path, backups_path)
        backup_path = backup.create(sensor_gk_name)
        print(f'[{count_steps}/{steps}] Создали backup. Путь: {backup_path}')
    except Exception as ex:
        logging.critical(f'Неизвестная ошибка: {ex}. Функция: find_file(). Traceback: {traceback.format_exc()}')
        print(f'Неизвестная ошибка: {ex}. Функция: find_file()')
        return

    # ищет в файле лист с ТС
    try:
        cell_ts_dict = find_cell_ts(file_path, sensor_ts_number)
        if cell_ts_dict.get('error'):
            logging.error(f'Ошибка: {cell_ts_dict["error"]}')
            print('Ошибка:', cell_ts_dict['error'])
            backup.delete()
            return
        wb = cell_ts_dict['wb']
        wsheet = cell_ts_dict['wsheet']
        cell_ts = cell_ts_dict['cell_ts']
        print(f'[{count_steps}/{steps}] Нашли ячейку с ТС')
        count_steps += 1
    except Exception as ex:
        logging.critical(f'Неизвестная ошибка: {ex}. Функция: find_cell_ts(). Traceback: {traceback.format_exc()}')
        print(f'Неизвестная ошибка: {ex}. Функция: find_cell_ts()')
        backup.delete()
        return

    # ищем хедер на странице
    try:
        header_columns = find_header_in_sheet(wsheet)
        if header_columns.get('error'):
            logging.error(f'Ошибка: {header_columns["error"]}')
            print('Ошибка:', header_columns['error'])
            backup.delete()
            return
        print(f'[{count_steps}/{steps}] Нашли хедер')
        count_steps += 1
    except Exception as ex:
        logging.critical(f'Неизвестная ошибка: {ex}. Функция find_header_in_sheet(). Traceback: {traceback.format_exc()}')
        print(f'Неизвестная ошибка: {ex}. Функция find_header_in_sheet()')
        backup.delete()
        return

    # 
    try:
        input_row, after_row, temperatures_length, is_first_input = (
            get_header_lengths_and_input_row(wsheet, cell_ts, header_columns)
        )
    except Exception as ex:
        logging.critical(f'Неизвестная ошибка: {ex}. Функция get_header_lengths_and_input_row(). Traceback: {traceback.format_exc()}')
        print(f'Неизвестная ошибка: {ex}. Функция get_header_lengths_and_input_row()')
        backup.delete()
        return

    # разъединяем все ячейки после хедера
    try:
        merged_cells, last_header_row = unmerge_all_cells_after_header(wsheet, after_row)
        print(f'[{count_steps}/{steps}] Разъединили ячейки')
        count_steps += 1
    except Exception as ex:
        logging.critical(f'Неизвестная ошибка: {ex}. Функция unmerge_all_cells_after_header(). Traceback: {traceback.format_exc()}')
        print(f'Неизвестная ошибка: {ex}. Функция unmerge_all_cells_after_header()')
        backup.delete()
        return

    # вставляем новую строку
    try:
        new_row = add_new_row(wsheet, cell_ts, input_row)
        if new_row is None:
            logging.error('Ошибка: Не удалось добавить новую строку')
            print('Ошибка: Не удалось добавить новую строку')
            backup.delete()
            return
        print(f'[{count_steps}/{steps}] Вставили новую строку')
        count_steps += 1
    except Exception as ex:
        logging.critical(f'Неизвестная ошибка: {ex}. Функция add_new_row(). Traceback: {traceback.format_exc()}')
        print(f'Неизвестная ошибка: {ex}. Функция add_new_row()')
        backup.delete()
        return

    # добавляем стили к новой строке (стили берутся на основе предыдущей строки)
    try:
        make_style_for_new_row(wsheet, new_row)
        print(f'[{count_steps}/{steps}] Применили стили к новой строке')
        count_steps += 1
    except Exception as ex:
        logging.critical(f'Неизвестная ошибка: {ex}. Функция make_style_for_new_row(). Traceback: {traceback.format_exc()}')        
        print(f'Неизвестная ошибка: {ex}. Функция make_style_for_new_row()')
        backup.delete()
        return

    # соединяем все ячейки обратно
    try:
        make_merged_cells(wsheet, merged_cells, cell_ts, last_header_row,
                          is_first_input)
        print(f'[{count_steps}/{steps}] Соединили ячейки обратно')
        count_steps += 1
    except Exception as ex:
        logging.critical(f'Неизвестная ошибка: {ex}. Функция make_merged_cells(). Traceback: {traceback.format_exc()}')
        print(f'Неизвестная ошибка: {ex}. Функция make_merged_cells()')
        backup.delete()
        return

    # вставляем данные
    try:
        put_data = put_data_to_excel(wsheet, input_row, header_columns,
                                     sensor_data, temperatures_length, cell_ts)
        if put_data is not None:
            logging.error(f'Ошибка: {put_data["error"]}')
            print('Ошибка:', put_data['error'])
            backup.delete()
            return
        print(f'[{count_steps}/{steps}] Вставили данные в строку')
        count_steps += 1
    except Exception as ex:
        print(traceback.format_exc())
        logging.critical(f'Неизвестная ошибка: {ex}. Функция put_data_to_excel(). Traceback: {traceback.format_exc()}')
        print(f'Неизвестная ошибка: {ex}. Функция put_data_to_excel()')
        backup.delete()
        return

    # сохраняем файл
    try:
        wb.save(file_path)
    except PermissionError:
        logging.error('Ошибка: Закройте файл, в который идет сохранение и попробуй снова')
        print('Ошибка: Закройте файл, в который идет сохранение и попробуй снова')
        backup.delete()
        return
    except Exception as ex:
        logging.critical(f'Неизвестная ошибка: {ex}. Функция wb.save(). Traceback: {traceback.format_exc()}')
        print(f'Неизвестная ошибка: {ex}. wb.save()')
        backup.delete()
        return

    logging.info(f'Работа завершена успешно. Данные добавлены. '
                 f'Файл сохранен. Путь к файлу: {file_path}')
    print(f'[{count_steps}/{steps}] Работа завершена. Данные добавлены. '
          f'Файл сохранен. Путь к файлу: {file_path}')

    count_steps += 1

    return True


if __name__ == '__main__':

    result = main()
    if result is True:
        time.sleep(10)
    elif result is None:
        input('Нажмите Enter для выхода...')
