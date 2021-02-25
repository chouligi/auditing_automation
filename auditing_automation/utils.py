import openpyxl
from openpyxl.workbook.workbook import Workbook
from types import GeneratorType
import os
import xlwings as xw

THIS_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(THIS_DIR, 'data')


def load_xl_workbook(path: str) -> Workbook:
    return openpyxl.load_workbook(path)


def get_worksheet_values_from_workbook(wookrbook: Workbook, worksheet_name: str) -> GeneratorType:
    return wookrbook[worksheet_name].values


def get_columns(data: GeneratorType) -> str:
    return next(data)[0:]


def copy_sheet_in_same_workbook(workbook_path: str, sheet_to_copy_name: str, name_of_new_sheet: str):
    workbook_to_copy = xw.Book(workbook_path)
    sheet_to_copy = workbook_to_copy.sheets[sheet_to_copy_name]

    # copy within the same sheet
    sheet_to_copy.api.copy_worksheet(after_=sheet_to_copy.api)

    copied_sheet = workbook_to_copy.sheets[1]

    copied_sheet.name = name_of_new_sheet


def create_new_workbook(output_path: str):
    # Create a new workbook in the given path
    # WORKBOOK_TO_COPY_PATH = os.path.join(DATA_DIR, 'workbook_to_copy.xlsx')
    # SHEET_TO_COPY_NAME = 'Trial Balance'

    # workbook_to_copy = xw.Book(WORKBOOK_TO_COPY_PATH)
    # sheet_to_copy = workbook_to_copy.sheets[SHEET_TO_COPY_NAME]
    new_workbook = xw.Book()

    new_workbook.save(f"{output_path}-leadsheet.xlsx")


def create_workbook_based_on_template():
    """
    This function reads a sheet from workbook A, and creates a new workbook B with a subset
    of the sheet of workbook A.
    :return:
    """
    return
