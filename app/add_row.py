import datetime
import re

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet.merge import MergedCellRange
from openpyxl.cell import Cell
from openpyxl.styles import Alignment

from validators import HeaderDataModel
from cutypes import FindCellTS, HeaderColumns, SensorData


# данная переменная будет из find_folder
# excel_file = 'C:\\Users\\Timoha\\Desktop\\SeverWork\\ExcelProjS\\for_tests.xlsx'

COUNT_TEMPERATURE_COLUMN = 20


def find_cell_ts(excel_file: str, ts_number: str) -> FindCellTS:
    """
    Ищем координаты объединненой ячейки с номером ТС
    """

    wb = openpyxl.load_workbook(excel_file)

    result = {'error': 'ТС не найден'}

    # Флаг нашли ТС или нет
    cell_found = False

    # перебор страниц ексель книги
    for wsheet in wb.worksheets:
        # перебираем строки
        for cell in wsheet['A']:
            if cell.value is not None:
                if cell.value.lower() == ts_number.lower():
                    result = {'cell_ts': cell, 'wsheet': wsheet, 'wb': wb}
                    cell_found = True
                    break
        if cell_found:
            break
    return result


def find_header_in_sheet(wsheet: Worksheet) -> HeaderColumns:
    """
    Находим headedr файла wb страницы wsheet
    """

    target_header_data = {
        'Дата\nпроведения\nизмерений': 'date',
        'Цикл': 'cycle', 
        # 'Фактическая\nглубина\nскважины, м':
        # 'actual_depth',
        'Глубина измерения, м': 'temperatures',
        'Высота надземной части скважины, м': 'height',
        'Глубина скв-ны с учётом надземной части, м': 'depth',
    }

    header_col_dict = {}
    target_row = None

    # ищем строку с header
    for row in wsheet.rows:
        if row[0].value == 'Номер ТС':
            target_row = row
    if target_row:
        for col in target_row:
            if col.value in target_header_data.keys():
                header_col_dict[target_header_data[col.value]] = col
        try:
            HeaderDataModel(**header_col_dict)
        except Exception as ex:
            field_error = str(ex).split('\n')[1]
            return {'error': f'Ошибка. Поле хедера для {field_error} не найдено.'}
        return header_col_dict

    else:
        return {'error': 'Header не найден. В excel файле.'}


def unmerge_all_cells_after_header(
        wsheet: Worksheet,
        header_date: str) -> tuple[list[MergedCellRange], int]:
    """
    Разъединяем все ячейки снизу хедера
    """

    # переменная отвечающая за то, с какой строки разъединять,
    # 2 - дефолтная длина хедера за исключением первой строки
    after_row = header_date.row + 2 
    merged_cells = list(wsheet.merged_cells)
    for merge in merged_cells:
        if merge.start_cell.row > after_row:
            wsheet.unmerge_cells(str(merge))
    return (merged_cells, after_row)


def add_new_row(wsheet: Worksheet, cell_ts: Cell) -> int | None:
    """
    Добавляем новую строку в excel файл
    """

    ts_row = cell_ts.row
    max_row = wsheet.max_row
    value = None

    while value is None and value != 'EndTable':
        ts_row += 1
        value = wsheet[f'A{ts_row}'].value
        if ts_row > max_row:
            return None

    wsheet.insert_rows(ts_row)
    return ts_row


def make_style_for_new_row(wsheet: Worksheet, input_row: int) -> None:
    """
    Переносим стили с верхней строки на текующую
    """

    for cell in wsheet[input_row]:
        cell_colum_letter = cell.column_letter
        cell_row = cell.row

        old_cell = wsheet[f'{cell_colum_letter}{cell_row-1}']
        if old_cell.has_style:
            cell._style = old_cell._style


def make_merged_cells(
        wsheet: Worksheet,
        merged_cells: list[MergedCellRange],
        cell_ts: Cell,
        header_row: int) -> None:
    """
    Соединяем ячейки обартно
    """

    # объединяем ячейки
    for merged_cell in merged_cells:
        if merged_cell.start_cell.row == cell_ts.row:
            start_merge_column = merged_cell.top[0][1]
            start_merge_row = merged_cell.top[0][0]
            end_merge_column = merged_cell.bottom[0][1]
            end_merge_row = merged_cell.bottom[0][0] + 1
            wsheet.merge_cells(
                start_row=start_merge_row,
                start_column=start_merge_column,
                end_row=end_merge_row,
                end_column=end_merge_column
            )
        elif merged_cell.start_cell.row > cell_ts.row:
            start_merge_column = merged_cell.top[0][1]
            start_merge_row = merged_cell.top[0][0] + 1
            end_merge_column = merged_cell.bottom[0][1]
            end_merge_row = merged_cell.bottom[0][0] + 1
            wsheet.merge_cells(
                start_row=start_merge_row,
                start_column=start_merge_column,
                end_row=end_merge_row,
                end_column=end_merge_column
            )
        elif merged_cell.start_cell.row < cell_ts.row:
            wsheet.merge_cells(str(merged_cell))

        # текст по центру
        merged_cells = list(wsheet.merged_cells)
        for merged_cell in merged_cells:
            if merged_cell.start_cell.row > header_row:
                start_cell = merged_cell.start_cell
                start_cell.alignment = Alignment(vertical='center',
                                                 horizontal='center')


def put_data_to_excel(
        wsheet: Worksheet,
        input_row: int,
        header_in_cells: HeaderColumns,
        data: SensorData) -> dict | None:
    """
    Вставляем данные в ексель файл
    """

    # имена колонок
    date_column = header_in_cells['date'].column_letter
    cycle_column = header_in_cells['cycle'].column_letter
    height_column = header_in_cells['height'].column_letter
    depth_column = header_in_cells['depth'].column_letter
    temperatures_column = header_in_cells['temperatures'].column
    
    date = data['date']
    height = data['height']
    depth = data['depth']
    temperatures = data['temperatures']
    
    # вставляем дату
    old_cell_date = wsheet[f'{date_column}{input_row-1}']
    cell_date = wsheet[f'{date_column}{input_row}']
    # put_date = put_date.strftime("%m/%d/%Y")
    cell_date.value = date
    if old_cell_date.has_style:
        cell_date._style = old_cell_date._style

    # вставляем цикл
    old_cell_cycle = wsheet[f'{cycle_column}{input_row-1}']
    old_cell_cycle_num = old_cell_cycle.value
    if old_cell_cycle_num is None:
        return {'error': 'Добавьте в конец excel книги EndTable'}
    cycle_number = int(''.join(re.findall(r'[0-9]', old_cell_cycle_num))) + 1
    cell_cycle = wsheet[f'{cycle_column}{input_row}']
    cell_cycle.value = f'«{cycle_number}» цикл'
    if old_cell_cycle.has_style:
        cell_cycle._style = old_cell_cycle._style

    # вставляем высоту
    old_cell_height = wsheet[f'{height_column}{input_row-1}']
    cell_height = wsheet[f'{height_column}{input_row}']
    cell_height.value = height
    if old_cell_height.has_style:
        cell_height._style = old_cell_height._style
    # вставляем глубину
    old_cell_depth = wsheet[f'{depth_column}{input_row-1}']
    cell_depth = wsheet[f'{depth_column}{input_row}']
    cell_depth.value = depth
    if old_cell_depth.has_style:
        cell_depth._style = old_cell_depth._style

    # вставляем температуры
    # В случае когда температур больше необходимого, берем их с конца

    temp_column_counter = 0
    count_temps = int((height + depth) / 2)
    if len(temperatures) > count_temps:
        temperatures.reverse()
        temperatures = temperatures[:count_temps]
        temperatures.reverse()
    if len(temperatures) < 20:
        none_list = [None] * (20 - len(temperatures))
        temperatures.extend(none_list)

    # вставляем значения
    while temp_column_counter < COUNT_TEMPERATURE_COLUMN:
        if temperatures[temp_column_counter] is None:
            wsheet.cell(input_row, temperatures_column + temp_column_counter).value = '-'
        else:
            wsheet.cell(input_row, temperatures_column + temp_column_counter).value = temperatures[temp_column_counter]
        temp_column_counter += 1


# def main():
#     # нашли ячейку с тс и страницу в файле, которая с ним связана
#     cell_ts_data = find_cell_ts(excel_file, 'Тс 12-4')
#     if cell_ts_data.get('error'):
#         return cell_ts_data['error']
#     else:
#         cell_ts = cell_ts_data['cell_ts']
#         wb = cell_ts_data['wb']
#         wsheet = cell_ts_data['wsheet']

#     # нашли хедер
#     header_cells = find_header_in_sheet(wsheet)

#     if header_cells.get('error'):
#         return header_cells['error']
#     else:
#         # header_validate = validators.header_data_validator(header_cells)
#         # if header_validate.get('error'):
#         #     return header_validate.get('error')
#         # else:
#         date = header_cells['date']

#     merged_cells_list, header_row = unmerge_all_cells_after_header(wsheet, date)

#     input_row = add_new_row(wsheet, cell_ts) # добавили новую строку и вернули ее
#     if input_row is None:
#         return {'error': 'Не удалось добавить строку'}

#     make_style_for_new_row(wsheet, input_row)

#     make_merged_cells(wsheet, merged_cells_list, cell_ts, header_row)

#     sensor_data = get_sensor_data('n0')
#     if sensor_data.get('error'):
#         return sensor_data.get('error')
#     # sensor_validate = validators.sensor_data_validator(sensor_data)
#     # if sensor_validate.get('error'):
#     #     return sensor_validate.get('error')

#     put_data_to_excel(wsheet, input_row, header_cells, sensor_data)

#     try:
#         wb.save('test.xlsx')
#     except PermissionError:
#         return {'error': 'Закройте файл, в который идет сохранение и попробуй снова.'}

# print(main())