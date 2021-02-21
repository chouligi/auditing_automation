import openpyxl
from openpyxl.workbook.workbook import Workbook
from types import GeneratorType


def load_xl_workbook(path: str) -> Workbook:
    return openpyxl.load_workbook(path)


def get_worksheet_values_from_workbook(wookrbook: Workbook, worksheet_name: str) -> GeneratorType:
    return wookrbook[worksheet_name].values


def get_columns(data: GeneratorType) -> str:
    return next(data)[0:]
