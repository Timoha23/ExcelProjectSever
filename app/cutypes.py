from datetime import datetime
from typing import TypedDict

from openpyxl.cell import Cell
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet


class SensorData(TypedDict):
    date: datetime
    gk_name: str
    ts_number: str
    depth: int | float
    height: int | float
    cargo_height: int | float


class FindCellTS(TypedDict):
    cell_ts: Cell
    wb: Workbook
    wsheet: Worksheet


class HeaderColumns(TypedDict):
    date: Cell
    cycle: Cell
    temperatures: Cell
    depth: Cell
    actual_depth: Cell
    height: Cell
    avg_temp: Cell
