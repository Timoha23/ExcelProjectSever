from typing import Union
from datetime import datetime

from pydantic import BaseModel, Field


class SensorDataModel(BaseModel):
    gk_name: str = Field(...)
    date: datetime = Field(...)
    ts_number: str = Field(...)
    depth: Union[int, float] = Field(...)
    height: Union[int, float] = Field(...)
    temperatures: list[Union[int, float]] = Field(..., min_items=1)
    cargo_height: Union[int, float] = Field(...)  # высота груза


class HeaderDataModel(BaseModel):
    date: object = Field(...)
    cycle: object = Field(...)
    height: object = Field(...)
    depth: object = Field(...)
    actual_depth: object = Field(...)
    temperatures: object = Field(...)
