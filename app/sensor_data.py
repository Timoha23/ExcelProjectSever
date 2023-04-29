import openpyxl
from pydantic.error_wrappers import ValidationError

from settings import SENSOR_DATA_PATH
from validators import SensorDataModel
from cutypes import SensorData


def get_sensor_data(number: str) -> SensorData:
    """
    Получаем данные из файла ПКЦД.xlsx
    Данная функция принимает number (n734)
    """

    wb = openpyxl.load_workbook(SENSOR_DATA_PATH)

    ws = wb.active
    rows = ws.max_row
    data = {'temperatures': []}

    for i in range(1, rows+1):
        if ws[f'A{i}'].value == number:
            sensors_count = int(ws[i+1][2].value.split(':')[1])
            data['date'] = ws[i][1].value # дата
            data['gk_name'] = ws[i+1][4+sensors_count].value # имя газового куста
            data['ts_number'] = ws[i+1][5+sensors_count].value # номер ТС
            data['depth'] = ws[i+1][6+sensors_count].value # глубина
            data['height'] = ws[i+1][7+sensors_count].value # высота
            data['cargo_height'] = ws[i+1][8+sensors_count].value # высота груза

            for index, el in enumerate(ws[i+1][4:]):
                if index < sensors_count:
                    data['temperatures'].append(el.value)
            try:
                SensorDataModel(**data)
            except ValidationError as ex:
                field_error = str(ex).split('\n')[1]
                return {'error': f'Ошибка в поле {field_error}. Проверьте данное поле'}
            return data
    return {'error': f'Номер датчика {number} не найден.'}
