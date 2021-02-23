import openpyxl
from openpyxl.workbook.workbook import Workbook
from types import GeneratorType


def load_xl_workbook(path: str) -> Workbook:
    return openpyxl.load_workbook(path)


def get_worksheet_values_from_workbook(wookrbook: Workbook, worksheet_name: str) -> GeneratorType:
    return wookrbook[worksheet_name].values


def get_columns(data: GeneratorType) -> str:
    return next(data)[0:]


def create_new_workbook(output_path: str):
    # Create a new workbook in the given path
    return


def create_workbook_based_on_template():
    """
    This function reads a sheet from workbook A, and creates a new workbook B with a subset
    of the sheet of workbook A.
    :return:
    """
    return
