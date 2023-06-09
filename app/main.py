import logging
import os
import traceback
from pathlib import Path

from actions import (add_new_row, find_cell_ts, find_header_in_sheet,
                     get_header_lengths_and_input_row, make_merged_cells,
                     make_style_for_new_row, put_data_to_excel,
                     unmerge_all_cells_after_header,
                     change_formuls_for_avg_temp,
                     change_formuls_for_actual_depth)
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


def run(sensor: str):
    steps = 14
    count_steps = 1
    sensor_number = sensor
    logging.info(f'Начали работу. Датчик {sensor_number}')
    print(f'[{count_steps}/{steps}] Начинаем работу')
    count_steps += 1

    # проверяем пути
    try:
        paths = get_paths()
        for key, value in paths.items():
            if value is None:
                return (f'Ошибка. Проверьте путь для {key} в файле settings.')
        data_folder_path, sensor_data_path = (
            paths['DATA_FOLDER_PATH'], paths['SENSOR_DATA_PATH']
        )
        print(f'[{count_steps}/{steps}] Получили пути к файлам')
        count_steps += 1
    except FileNotFoundError:
        logging.error('Ошибка: Файл settings.txt не найден.')
        print('Ошибка: Файл settings.txt не найден. Проверьте что он '
              'существует.')
        return
    except Exception as ex:
        logging.critical(f'Неизвестная ошибка: {ex}. Функция: get_paths().'
                         f' Traceback: {traceback.format_exc()}')
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
        logging.critical(f'Неизвестная ошибка: {ex}. Функция: '
                         f'get_sensor_data(). Traceback: '
                         f'{traceback.format_exc()}')
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
        logging.critical(f'Неизвестная ошибка: {ex}. Функция: find_file().'
                         f' Traceback: {traceback.format_exc()}')
        print(f'Неизвестная ошибка: {ex}. Функция: find_file()')
        return

    # создаем бекап файла
    try:
        backup = Backup(file_path, backups_path)
        backup_path = backup.create(sensor_gk_name)
        print(f'[{count_steps}/{steps}] Создали backup. Путь: {backup_path}')

    except Exception as ex:
        logging.critical(f'Неизвестная ошибка: {ex}. Функция: find_file().'
                         f' Traceback: {traceback.format_exc()}')
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
        logging.critical(f'Неизвестная ошибка: {ex}. Функция: find_cell_ts().'
                         f' Traceback: {traceback.format_exc()}')
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
        logging.critical(f'Неизвестная ошибка: {ex}. Функция'
                         f' find_header_in_sheet(). '
                         f'Traceback: {traceback.format_exc()}')
        print(f'Неизвестная ошибка: {ex}. Функция find_header_in_sheet()')
        backup.delete()
        return

    # s
    try:
        input_row, after_row, temperatures_length, is_first_input = (
            get_header_lengths_and_input_row(wsheet, cell_ts, header_columns)
        )
    except Exception as ex:
        logging.critical(f'Неизвестная ошибка: {ex}. Функция'
                         f' get_header_lengths_and_input_row(). '
                         f'Traceback: {traceback.format_exc()}')
        print(f'Неизвестная ошибка: {ex}. Функция'
              f' get_header_lengths_and_input_row()')
        backup.delete()
        return

    # разъединяем все ячейки после хедера
    try:
        merged_cells, last_header_row = unmerge_all_cells_after_header(
            wsheet, after_row
        )
        print(f'[{count_steps}/{steps}] Разъединили ячейки')
        count_steps += 1
    except Exception as ex:
        logging.critical(f'Неизвестная ошибка: {ex}. Функция'
                         f' unmerge_all_cells_after_header().'
                         f' Traceback: {traceback.format_exc()}')
        print(f'Неизвестная ошибка: {ex}. Функция'
              f' unmerge_all_cells_after_header()')
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
        logging.critical(f'Неизвестная ошибка: {ex}. Функция add_new_row().'
                         f' Traceback: {traceback.format_exc()}')
        print(f'Неизвестная ошибка: {ex}. Функция add_new_row()')
        backup.delete()
        return

    # сдвигаем формулы для avg_temp
    try:
        change_formuls_for_avg_temp(wsheet, input_row, header_columns)
        print(f'[{count_steps}/{steps}] Сдвинули формулы для ср. темп.')
        count_steps += 1
    except Exception as ex:
        logging.critical(f'Неизвестная ошибка: {ex}. '
                         f'change_formuls_for_avg_temp().'
                         f' Traceback: {traceback.format_exc()}')
        print(f'Неизвестная ошибка: {ex}. '
              f'Функция change_formuls_for_avg_temp()')
        backup.delete()
        return

    # сдвигаем формулы для фактической глубины скважины
    try:
        change_formuls_for_actual_depth(wsheet, input_row, header_columns)
        print(f'[{count_steps}/{steps}] Сдвинули формулы для фактической'
              f' глубины скважины')
        count_steps += 1
    except Exception as ex:
        logging.critical(f'Неизвестная ошибка: {ex}. '
                         f'change_formuls_for_actual_depth().'
                         f' Traceback: {traceback.format_exc()}')
        print(f'Неизвестная ошибка: {ex}. '
              f'Функция change_formuls_for_actual_depth()')
        backup.delete()
        return

    # добавляем стили к новой строке
    # (стили берутся на основе предыдущей строки)
    try:
        make_style_for_new_row(wsheet, new_row)
        print(f'[{count_steps}/{steps}] Применили стили к новой строке')
        count_steps += 1
    except Exception as ex:
        logging.critical(f'Неизвестная ошибка: {ex}. Функция'
                         f' make_style_for_new_row().'
                         f' Traceback: {traceback.format_exc()}')
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
        logging.critical(f'Неизвестная ошибка: {ex}. Функция '
                         f'make_merged_cells(). '
                         f'Traceback: {traceback.format_exc()}')
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
        logging.critical(f'Неизвестная ошибка: {ex}. Функция '
                         f'put_data_to_excel(). '
                         f'Traceback: {traceback.format_exc()}')
        print(f'Неизвестная ошибка: {ex}. Функция put_data_to_excel()')
        backup.delete()
        return

    # сохраняем файл
    try:
        wb.save(file_path)
    except PermissionError:
        logging.error('Ошибка: Закройте файл, в который идет сохранение и'
                      ' попробуй снова')
        print('Ошибка: Закройте файл, в который идет сохранение и попробуй'
              ' снова')
        backup.delete()
        return
    except Exception as ex:
        logging.critical(f'Неизвестная ошибка: {ex}. Функция wb.save().'
                         f' Traceback: {traceback.format_exc()}')
        print(f'Неизвестная ошибка: {ex}. wb.save()')
        backup.delete()
        return

    logging.info(f'Работа завершена успешно. Данные добавлены. '
                 f'Файл сохранен. Путь к файлу: {file_path}')
    print(f'[{count_steps}/{steps}] Работа завершена. Данные добавлены. '
          f'Файл сохранен. Путь к файлу: {file_path}')

    count_steps += 1

    return True


def main():
    sensor_numbers = input(
        'Введите номер сенсора (сенсоров). Пример(n232, n0, n562): '
    )
    sensors = sensor_numbers.replace(' ', '').split(',')
    errors = 0
    success = 0
    errors_sensors = []
    maximum = len(sensors)

    for sensor in sensors:
        res = run(sensor)
        if res is True:
            success += 1
            print(f'Датчик {sensor} успешно внесен')
            print(f'Успешно добавлено: {success}/{maximum}. '
                  f'Ошибочных датчиков: {errors}/{maximum}')
        else:
            errors += 1
            errors_sensors.append(sensor)
            print(f'Датчик {sensor} не внесен')
    print(f'Работа завершена. Успешно добавлено: {success}/{maximum}.'
          f' Ошибочных датчиков: {errors}/{maximum}.'
          f' Датчики с ошибками: {errors_sensors}')


if __name__ == '__main__':
    result = main()
    input('Нажмите Enter для выхода...')
